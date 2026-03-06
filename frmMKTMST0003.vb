Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Friend Class frmMKTMST0003
    Inherits System.Windows.Forms.Form
    '---------------------------------------------------------------------------------------------------------
    'Copyright(c)       -   MIND
    'Name of module     -   frmMKTMST0003
    'Created by         -
    'Created Date       -
    'Description        -   Customer Item Master
    'Revision History   -
    'Changes done By nisha on 03/07/2003 for adding Quot in update of upd_userid
    'Changes done By Shambhu on 13/01/2004 Add One column Model Name
    'Changes done by Brij on 22/02/2005 to make Customer,Item Code & internal part no mapping 1-1
    'Changes done by Brij on 21/03/2005 in AlreadyExist function to make Customer,Item Code & internal part no mapping 1-1
    'REVISED BY     :   SHUBHRA VERMA
    'REVISED ON     :   15 APR 2009
    'ISSUE ID       :   eMpro-20090415-30149
    'DECRIPTION     :   ADD Upload Flag and Active Flag Columns in the GRID of Customer Item Master.
    'REVISED BY     :   SHUBHRA VERMA
    'REVISED ON     :   22 APR 2009
    'ISSUE ID       :   eMpro-20090423-30547
    'DECRIPTION     :   in Customer Item Master, though the row is non editable, active flag 
    '                   and schedule upload flag should be editable
    'REVISED BY     :   MANOJ VAISH
    'REVISED ON     :   12 MAY 2009
    'ISSUE ID       :   eMpro-20090512-31252
    'DECRIPTION     :   ADD Commodity Column in the GRID of Customer Item Master and
    '                   while defining new customer drawing number schedule upload Flag is not updated 
    '----------------------------------------------------------------------------------------------------------
    'Revised By     : Manoj Kr. Vaish
    'Issue ID       : eMpro-20090611-32362
    'Revision Date  : 12 Jun 2009
    'History        : Add New field of Container type and Packing Level Code-Hilex Nissan CSV File Genaration
    'Modified By sanchi on 27 April 2011 modified to support MultiUnit functionality
    '----------------------------------------------------------------
    'Revised By     : Prashant Rajpal
    'Issue ID       : 10135085
    'Revision Date  : 08 sep 2011
    'History        : unhandled error comes During saving time in NEW mode ..
    '****************************************************************************************
    'Revised By     : Vinod Singh
    'Issue ID       : 10136194
    'Revision Date  : 16 Sep 2011
    'History        : Cust Part Desc was not updating as cust part no was also editable in Edit Mode
    '****************************************************************************************
    ' Revised By     :   Pankaj Kumar
    ' Revision Date  :   10 Oct 2011
    ' Description    :   Modified for MultiUnit Change Management
    '****************************************************************************************
    'MODIFIED BY AVANISH PATHAK ON 8 NOV FOR MULTIUNIT CHANGE MANAGEMENT
    '****************************************************************************************
    'REVISED BY     : SAURAV KUMAR
    'ISSUE ID       : 10195532
    'REVISION DATE  : 16 FEB 2012
    'HISTORY        : ADDITION OF SEARCH CRITERIA ON PART NUMBER AND DESCRIPTION
    '****************************************************************************************
    'MODIFIED BY AJAY SHUKLA ON 9 MARCH FOR MULTIUNIT CHANGE MANAGEMENT
    '****************************************************************************************
    'REVISED BY     : SAURAV KUMAR
    'ISSUE ID       : 10222401 
    'REVISION DATE  : 05 sep 2012
    'HISTORY        : CHANGES WHILE UPDATING DATA
    '****************************************************************************************
    'Created By     : Parveen Kumar
    'Created On     : 16 FEB 2015
    'Description    : eMPro Vehicle BOM
    'Issue ID       : 10737738 
    '-------------------------------------------------------------------------------------------
    'Created By     : Vinod Singh
    'Created On     : 22 June 2015
    'Issue ID       : 10808160 - EOP Changes 
    '-------------------------------------------------------------------------------------------
    'REVISED BY     -  PRASHANT RAJPAL
    'REVISED ON     -  18 SEP 2015
    'PURPOSE        -  10869290 -SERVICE INVOICE 
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'REVISED BY     -  MILIND MISHRA
    'REVISED ON     -  27 OCT 2016
    'PURPOSE        -  101157667 -Auto Invoice part functionality  
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'MODIFIED BY     -  MILIND MISHRA
    'MODIFIED ON     -  15 JUL 2019
    'PURPOSE         -  1234MILIND (DUMMY CODE FOR BOOKMARK PURPOSE)- OPTIMISING CUSTOMER ITEM MASTER FORM 
    '-------------------------------------------------------------------------------------------------------------------------------------------------------


    Dim mintIndex As Short
    Dim blnCheckDelItems As Boolean
    Dim strDelete As String
    Dim strInsert As String
    Dim strupdate As String
    Dim strUpdateBudget As String
    Dim blnAllowMultiPartNo As Boolean
    Dim blnAllowBudget As Boolean
    Dim strFinalPackingCode As String
    Dim strInsertbudget As String
    Dim strDeletebudget As String
    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
        '----------------------------------------------------------------------------
        'Argument       :   NIL
        'Return Value   :   NIL
        'Function       :   To show help on Account Master
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        Dim varRetVal As Object
        Dim strHelp As String
        Dim rscusthelp As New ClsResultSetDB
        On Error GoTo Err_Handler
        Select Case CmdGrpCustomerITem.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                With Me.txtCustCode
                    strHelp = ShowList(1, .MaxLength, "", "b.customer_code", "b.cust_name", "customer_mst b", "", , , , , "b.UNIT_CODE")
                    .Focus()
                End With
                If Val(strHelp) = "-1" Then ' No record
                    Call ConfirmWindow(10170, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                Else
                    Me.txtCustCode.Text = strHelp
                    rscusthelp.GetResult("SELECT a.customer_Code,a.cust_name FROM Customer_mst a Where a.customer_code = '" & strHelp & "' and a.UNIT_CODE='" & gstrUNITID & "' ")
                    If rscusthelp.GetNoRows > 0 Then
                        Me.lblCustCodeDes.Text = rscusthelp.GetValue("cust_name")
                    End If
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                With Me.txtCustCode
                    strHelp = ShowList(1, .MaxLength, "", "a.Customer_Code", "A.CUST_NAME", "Customer_mst a,CustItem_Mst c", " AND a.CUSTOMER_Code=c.Account_Code AND a.UNIT_CODE=c.UNIT_CODE", "HELP", , , , , "a.UNIT_CODE")
                    .Focus()
                End With
                If Val(strHelp) = -1 Then ' No record
                    Call ConfirmWindow(10170, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                    txtCustCode.Focus()
                Else
                    Me.txtCustCode.Text = strHelp
                    rscusthelp.GetResult("SELECT a.customer_Code,a.cust_NAME FROM Customer_mst a,CustItem_Mst c Where a.CUSTOMER_Code=c.Account_Code and a.UNIT_CODE=c.UNIT_CODE AND a.CUSTOMER_code = '" & strHelp & "' AND a.UNIT_CODE='" & gstrUNITID & "' ")
                    If rscusthelp.GetNoRows > 0 Then 'RECORD FOUND
                        Me.lblCustCodeDes.Text = rscusthelp.GetValue("CUST_NAME")
                    End If
                End If
        End Select
        rscusthelp.ResultSetClose()
        rscusthelp = Nothing
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0003_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo errHandler
        mdifrmMain.CheckFormName = mintIndex
        txtCustCode.Focus()
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTMST0003_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo errHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTMST0003_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Err_Handler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            'Call ctlFormHeader1_Click
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0003_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo Err_Handler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                If Me.CmdGrpCustomerITem.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call Me.CmdGrpCustomerITem.Revert()
                        Call EnableControls(False, Me, True)
                        txtCustCode.Enabled = True : txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelp.Enabled = True
                        CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                        CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                        CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                        txtCustCode.Focus()
                        gblnCancelUnload = False
                        gblnFormAddEdit = False
                        'Get Server Date
                    Else
                        Me.ActiveControl.Focus()
                    End If
                End If
        End Select
        GoTo EventExitSub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub frmMKTMST0003_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo errHandler
        'Add Form Name To Window List
        spdITemDetails.DisplayRowHeaders = False
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        'Fill Lebels From Resource File
        Call FillLabelFromResFile(Me) 'Fill Labels >From Resource File
        Call FitToClient(Me, Frame1, ctlFormHeader1, CmdGrpCustomerITem, 200)
        'Set Help Pictures At Command Button
        cmdHelp.Image = My.Resources.ico111.ToBitmap
        'Initially Disable All Controls
        Call EnableControls(False, Me, True)
        Call AddlabelToGrid()
        Dim rsAllowBudget_flag As ClsResultSetDB
        rsAllowBudget_flag = New ClsResultSetDB
        rsAllowBudget_flag.GetResult("SELECT AllowBudget_flag from Sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsAllowBudget_flag.RowCount > 0 Then
            blnAllowBudget = IIf(rsAllowBudget_flag.GetValue("AllowBudget_flag") = True, True, False)
        End If
        rsAllowBudget_flag = Nothing
        If blnAllowBudget = True Then
            mP_Connection.Execute("delete from tmp_budgetitem_mst where ip_address = '" & gstrIpaddressWinSck & "' AND UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        End If
        txtCustCode.Enabled = True
        txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        cmdHelp.Enabled = True
        CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
        CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
        CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0003_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        On Error GoTo errHandler
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
            If Me.CmdGrpCustomerITem.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                enmValue = ConfirmWindow(10055, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Or enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call CmdGrpCustomerITem_ButtonClick(eventSender, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE))
                    ElseIf enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Then
                        gblnCancelUnload = False
                        gblnFormAddEdit = False
                    End If
                Else
                    gblnCancelUnload = True
                    gblnFormAddEdit = True
                    Me.CmdGrpCustomerITem.Focus()
                End If
            Else
                Me.Dispose()
                Exit Sub
            End If
        End If
        'Checking The Status
        If gblnCancelUnload = True Then eventArgs.Cancel = True
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTMST0003_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo errHandler
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Me.Dispose()
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub AddlabelToGrid()
        On Error GoTo errHandler
        spdITemDetails.MaxRows = 0
        spdITemDetails.MaxCols = 25  'change
        With spdITemDetails
            .PrintRowHeaders = False
            .set_RowHeight(0, 13)
            .Row = 0 : .Col = 1 : .Text = " Delete"
            .Row = 0 : .Col = 2 : .Text = "Internal Part No." : .set_ColWidth(2, 15)
            .Row = 0 : .Col = 3 : .Text = " " : .set_ColWidth(3, 3)
            .Row = 0 : .Col = 4 : .Text = "Internal Part Description" : .set_ColWidth(4, 20)
            .Row = 0 : .Col = 5 : .Text = "Customer Part No" : .set_ColWidth(5, 12)
            .Row = 0 : .Col = 6 : .Text = "Customer Part Description" : .set_ColWidth(6, 20)
            .Row = 0 : .Col = 7 : .Text = " Edit/New "
            .Col = 1 : .Col2 = 1 : .BlockMode = True : .ColHidden = True : .BlockMode = False
            .Col = 7 : .Col2 = 7 : .BlockMode = True : .ColHidden = True : .BlockMode = False
            .Row = 0 : .Col = 8 : .Text = " Bin Quantity" : .set_ColWidth(10, 5)
            '.Row = 0 : .Col = 9 : .Text = "Model Name" : .ColHidden = True : .set_ColWidth(9, 18)
            .Row = 0 : .Col = 9 : .Text = "Model Name" : .set_ColWidth(9, 18)
            .Col = 9 : .Col2 = 9 : .BlockMode = True : .BlockMode = False
            .Row = 0 : .Col = 10 : .Text = "Active" : .set_ColWidth(10, 3)
            .Row = 0 : .Col = 11 : .Text = "Upload Flag" : .set_ColWidth(11, 3)
            .Row = 0 : .Col = 12 : .Text = "Colour" : .set_ColWidth(12, 8)
            .Row = 0 : .Col = 13 : .Text = " " : .set_ColWidth(13, 3)
            .Row = 0 : .Col = 14 : .Text = "Category" : .set_ColWidth(14, 8)
            .Row = 0 : .Col = 15 : .Text = " " : .set_ColWidth(15, 3)
            .Row = 0 : .Col = 16 : .Text = "Commodity" : .set_ColWidth(16, 8)
            .Row = 0 : .Col = 17 : .Text = " " : .set_ColWidth(17, 3)
            .Row = 0 : .Col = 18 : .Text = "Container Type" : .set_ColWidth(18, 8)
            .Row = 0 : .Col = 19 : .Text = "Packing Level" : .set_ColWidth(19, 8)
            .Row = 0 : .Col = 20 : .Text = "Model" : .set_ColWidth(20, 4)
            'change New Column Auto Inv Part
            .Row = 0 : .Col = 21 : .Text = "Auto Inv Part" : .set_ColWidth(10, 3)
            .Row = 0 : .Col = 22 : .Text = "Shop Code" : .set_ColWidth(5, 3)
            .Row = 0 : .Col = 23 : .Text = "Gate No" : .set_ColWidth(12, 5)
            .Row = 0 : .Col = 24 : .Text = "End Date" : .set_ColWidth(12, 5)
            .Row = 0 : .Col = 25 : .Text = "Declaration No" : .set_ColWidth(4, 20)
            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            .EditEnterAction = FPSpreadADO.EditEnterActionConstants.EditEnterActionNext
            '.ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsHorizontal
        End With
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Function IsExistFieldValue(ByRef pstrFieldValue As String, ByRef pstrFeildName As String, ByRef pstrTableName As String, Optional ByRef pstrCondition As String = "") As Boolean
        '****************************************************
        'Created By     -  Nisha
        'Description    -  To Check Validity Of Field Data Whethet it Exists In The
        '                  Database Or Not
        'Arguments      -  pstrFieldValue - Field Text, pstrFeildName - Column Name
        '               -  pstrTableName - Table Name, pstrCondition - Optional Parameter For Condition
        '****************************************************
        On Error GoTo errHandler
        IsExistFieldValue = False
        Dim strTableSql As String 'Declared To Make Select Query
        Dim rsExistData As ClsResultSetDB
        If Len(Trim(pstrCondition)) > 0 Then
            strTableSql = "select " & Trim(pstrFeildName) & " from " & Trim(pstrTableName) & " where " & Trim(pstrFeildName) & "='" & Replace(Trim(pstrFieldValue), "'", "") & "' and " & Trim(pstrCondition)
        Else
            strTableSql = "select " & Trim(pstrFeildName) & " from " & Trim(pstrTableName) & " where " & Trim(pstrFeildName) & "='" & Replace(Trim(pstrFieldValue), "'", "") & "'"
        End If
        rsExistData = New ClsResultSetDB
        rsExistData.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsExistData.GetNoRows > 0 Then
            IsExistFieldValue = True
        Else
            IsExistFieldValue = False
        End If
        rsExistData.ResultSetClose()
        Exit Function
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub txtCustCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustCode.TextChanged
        On Error GoTo errHandler
        If Len(Trim(txtCustCode.Text)) = 0 Then
            lblCustCodeDes.Text = ""
            spdITemDetails.MaxRows = 0
        End If
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtCustCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo errHandler
        If KeyCode = 112 Then
            Call cmdHelp_Click(cmdHelp, New System.EventArgs())
        End If
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtCustCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo errHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
            Case 39, 34, 96
                KeyAscii = 0
            Case 13
        End Select
        GoTo EventExitSub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
        On Error GoTo errHandler
        Select Case KeyCode
            Case System.Windows.Forms.Keys.Return
                CmdGrpCustomerITem.Focus()
                If CmdGrpCustomerITem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    If spdITemDetails.Enabled = True Then spdITemDetails.Focus()
                End If
        End Select
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtCustCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo errHandler
        If Len(Trim(txtCustCode.Text)) = 0 Then GoTo EventExitSub
        Select Case CmdGrpCustomerITem.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If IsExistFieldValue(Trim(txtCustCode.Text), "customer_code", "Customer_Mst", "UNIT_CODE='" & gstrUNITID & "'") = True Then
                    lblCustCodeDes.Text = ReturnDescription("cust_name", "customer_mst", "customer_Code ='" & Trim(txtCustCode.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'")
                    If Not spdITemDetails.MaxRows > 0 Then
                        Call AddNewRowType()
                        spdITemDetails.MaxRows = 1 : spdITemDetails.Enabled = True ' : spdITemDetails.Focus()
                        spdITemDetails.Col = 2 : spdITemDetails.Lock = True
                        spdITemDetails.Col = 9 : spdITemDetails.Lock = True
                    End If
                Else
                    MsgBox("Invalid Customer Code", MsgBoxStyle.Information, "eMPro")
                    txtCustCode.Text = "" : txtCustCode.Focus()
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                If IsExistFieldValue(Trim(txtCustCode.Text), "b.customer_code", "CustItem_MSt a,Customer_Mst b", "a.ACCOUNT_Code = b.CUSTOMER_Code AND A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" & gstrUNITID & "'") = True Then
                    spdITemDetails.MaxRows = 0
                    txtItemCode.Enabled = True
                    cmdItemHelp.Enabled = True
                    btnExportExcel.Enabled = True
                    'Call DisplayDetailsinGrid()
                    'Call DisplayDetailsinGridNew()
                    'With spdITemDetails
                    '    .BlockMode = True
                    '    .Col = 1
                    '    .Row = 1
                    '    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    '    .BlockMode = False
                    'End With
                Else
                    MsgBox("Invalid Customer Code", MsgBoxStyle.Information, "eMPro")
                    txtCustCode.Text = "" : txtCustCode.Focus()
                End If
        End Select
        GoTo EventExitSub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Public Function DisplayDetailsinGrid() As Object
        On Error GoTo errHandler
        Dim rsCustItem As ClsResultSetDB
        Dim rsItemMst As ClsResultSetDB
        Dim intLoopCounter As Short
        Dim intMaxCounter As Short
        Dim StrItemCode As String
        Dim NoOfRed As Short
        Dim intPackingCode As Short
        Dim rsPackingCode As ClsResultSetDB
        Dim rsCustItemBudget As ClsResultSetDB
        Dim rsColourBudget As ClsResultSetDB
        Dim rsCommodityBudget As ClsResultSetDB
        rsCustItem = New ClsResultSetDB
        rsCustItemBudget = New ClsResultSetDB
        rsCustItem.GetResult("Select * from CustITem_Mst  Where UNIT_CODE='" & gstrUNITID & "' AND Account_code ='" & txtCustCode.Text.Trim & "' ORDER BY Cust_drgNo")
        If rsCustItem.GetNoRows > 0 Then
            intMaxCounter = rsCustItem.GetNoRows
            rsCustItem.MoveFirst()
            With spdITemDetails
                For intLoopCounter = 1 To intMaxCounter
                    If spdITemDetails.MaxRows < intLoopCounter Then
                        'Call addNewInSpread()
                    End If
                    Call spdITemDetails.SetText(2, intLoopCounter, rsCustItem.GetValue("Item_code"))
                    StrItemCode = rsCustItem.GetValue("Item_code")
                    rsItemMst = New ClsResultSetDB
                    rsItemMst.GetResult("Select Description,AUTO_INVOICE_PART from Item_Mst Where ITem_code ='" & StrItemCode & "' AND UNIT_CODE='" & gstrUNITID & "'")
                    Call spdITemDetails.SetText(4, intLoopCounter, rsItemMst.GetValue("Description"))
                    If rsItemMst.GetValue("AUTO_INVOICE_PART") = True Then       'to get auto inv part value
                        Call spdITemDetails.SetText(21, intLoopCounter, 1)
                        With spdITemDetails
                            .Col = 21 : .Col2 = 21
                            .Row = .ActiveRow : .Row2 = .ActiveRow
                            .BlockMode = True
                            .Lock = True
                            .BlockMode = False

                        End With
                    End If


                    Call spdITemDetails.SetText(5, intLoopCounter, rsCustItem.GetValue("Cust_DrgNo"))
                    Call spdITemDetails.SetText(6, intLoopCounter, rsCustItem.GetValue("Drg_desc"))
                    Call spdITemDetails.SetText(8, intLoopCounter, rsCustItem.GetValue("BinQuantity"))
                    Call spdITemDetails.SetText(9, intLoopCounter, rsCustItem.GetValue("VarModel"))
                    If rsCustItem.GetValue("Active") = "False" Then
                        Call spdITemDetails.SetText(10, intLoopCounter, 0)
                    Else
                        Call spdITemDetails.SetText(10, intLoopCounter, 1)
                    End If
                    If rsCustItem.GetValue("schupldreqd") = "False" Then
                        Call spdITemDetails.SetText(11, intLoopCounter, 0)
                    Else
                        Call spdITemDetails.SetText(11, intLoopCounter, 1)
                    End If
                    If blnAllowBudget = True Then
                        rsCustItemBudget.GetResult("select colour_code,category_code,commodity_code,EndDate from budgetitem_mst where account_code='" & txtCustCode.Text & "' and item_code = '" & rsCustItem.GetValue("Item_code") & "' and cust_drgno='" & rsCustItem.GetValue("Cust_drgno") & "' AND UNIT_CODE='" & gstrUNITID & "'")
                        If rsCustItemBudget.GetNoRows > 0 Then
                            rsColourBudget = New ClsResultSetDB
                            Call spdITemDetails.SetText(12, intLoopCounter, rsCustItemBudget.GetValue("Colour_code"))
                            Call spdITemDetails.SetText(14, intLoopCounter, rsCustItemBudget.GetValue("category_code"))
                            Call spdITemDetails.SetText(16, intLoopCounter, rsCustItemBudget.GetValue("commodity_code"))
                            Call spdITemDetails.SetText(24, intLoopCounter, rsCustItemBudget.GetValue("EndDate"))
                        End If
                    End If
                    Call spdITemDetails.SetText(18, intLoopCounter, rsCustItem.GetValue("Container"))
                    intPackingCode = IIf(IsDBNull(rsCustItem.GetValue("Packing_Code")), 0, Val(rsCustItem.GetValue("packing_Code")))
                    Call GetPackingLevelCode()
                    spdITemDetails.TypeComboBoxClear(19, intLoopCounter)
                    spdITemDetails.TypeComboBoxList = strFinalPackingCode
                    rsPackingCode = New ClsResultSetDB()
                    rsPackingCode.GetResult("Select (Code +'-'+Key2)as Packing_Level from lists where UNIT_CODE='" & gstrUNITID & "' AND key1='Packing Level' and code=" & intPackingCode, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If rsPackingCode.GetNoRows > 0 Then
                        spdITemDetails.SetText(19, intLoopCounter, rsPackingCode.GetValue("Packing_Level"))
                    End If
                    rsPackingCode.ResultSetClose()
                    rsPackingCode = Nothing
                    rsCustItem.MoveNext()
                Next
                rsCustItem.ResultSetClose()
                rsCustItemBudget.ResultSetClose()
                .Enabled = True
                .Row = 1
                .Row2 = .MaxRows
                .Col = 1
                'If blnAllowBudget = True Then
                '    .Col2 = .MaxCols - 1
                'Else
                '    .Col2 = .MaxCols
                'End If
                .Col2 = .MaxCols
                .BlockMode = True
                .Lock = True
                .BlockMode = False
            End With
            spdITemDetails.Col = 1
            spdITemDetails.Col2 = 1
            spdITemDetails.BlockMode = True
            spdITemDetails.ColHidden = False
            spdITemDetails.BlockMode = False
            'If Not blnAllowBudget Then
            '    .Col()
            'End If
            Call CheckUsedItems()
            'To Check if All ITem Are in Used
            intMaxCounter = spdITemDetails.MaxRows
            'blnNotAllRed = True
            For intLoopCounter = 1 To intMaxCounter
                With spdITemDetails
                    .Row = intLoopCounter
                    .Row2 = intLoopCounter
                    If .ForeColor = System.Drawing.Color.Red Then
                        NoOfRed = NoOfRed + 1
                    End If
                End With
            Next
            CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
            CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
        Else
            MsgBox("No Item Defined For this Customer", MsgBoxStyle.Information, "eMPro")
            txtCustCode.Text = "" : txtCustCode.Focus()
        End If
        If blnAllowBudget = True Then
            mP_Connection.Execute("delete from tmp_budgetitem_mst where ip_address='" & gstrIpaddressWinSck & "' AND UNIT_CODE='" & gstrUNITID & "' ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            mP_Connection.Execute("insert into tmp_budgetitem_mst(account_code,cust_drgno,item_code,colour_code,category_code,commodity_code,MODEL_CODE,USAGE_QTY,ENT_DT,ENT_USERID,UPD_DT,UPD_USERID,ip_address,VARIANT_CODE,UNIT_CODE,DefaultModel,EndDate) select account_code,cust_drgno,item_code,colour_code,category_code,commodity_code,MODEL_CODE,USAGE_QTY,ENT_DT,ENT_USERID,UPD_DT,UPD_USERID,'" & gstrIpaddressWinSck & "',VARIANT_CODE,UNIT_CODE,DefaultModel,EndDate from budgetitem_mst where account_code='" & txtCustCode.Text.Trim & "' AND UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        End If
        Exit Function
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function ReturnDescription(ByRef pstrFeildName As String, ByRef pstrTableName As String, ByRef pstrCondition As String) As String
        '****************************************************
        'Created By     -  Nisha
        'Description    -  To Check Validity Of Field Data Whethet it Exists In The
        '                  Database Or Not
        'Arguments      -  pstrFieldValue - Field Text, pstrFeildName - Column Name
        '               -  pstrTableName - Table Name, pstrCondition - Optional Parameter For Condition
        '****************************************************
        On Error GoTo errHandler
        Dim strTableSql As String 'Declared To Make Select Query
        Dim rsReturnDes As ClsResultSetDB
        strTableSql = "select " & Trim(pstrFeildName) & " from " & Trim(pstrTableName) & " where " & Trim(pstrCondition)
        rsReturnDes = New ClsResultSetDB
        rsReturnDes.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsReturnDes.GetNoRows > 0 Then
            ReturnDescription = rsReturnDes.GetValue(pstrFeildName)
        End If
        rsReturnDes.ResultSetClose()
        Exit Function
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub AddNewRowType()
        On Error GoTo errHandler
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        With spdITemDetails
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .BackColorStyle = FPSpreadADO.BackColorStyleConstants.BackColorStyleUnderGrid
            .Col = 1 : .Col2 = .Col : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
            ': .set_ColWidth(1, 520)
            .TypeCheckCenter = True
            .ColsFrozen = 4
            .Row = .MaxRows
            .Col = 2
            .Col2 = .Col
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeMaxEditLen = 16
            '.set_ColWidth(2, 1500)
            .CellTag = False
            .Row = .MaxRows
            .Col = 3
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = "..."
            '.set_ColWidth(3, 315)
            '.TypeButtonPicture = My.Resources.ico111.ToBitmap          '1234MILIND
            .Row = .MaxRows
            .Col = 4
            .Col2 = .Col
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeMaxEditLen = 30
            '.set_ColWidth(4, 2000)
            .Row = .MaxRows
            .Col = 5
            .Col2 = .Col
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeMaxEditLen = 30
            '.set_ColWidth(5, 2000)
            .Row = .MaxRows
            .Col = 6
            .Col2 = .Col
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeMaxEditLen = 50
            '.set_ColWidth(6, 2500)
            .Row = .MaxRows
            .Col = 7
            .Col2 = .Col
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .Row = .MaxRows
            .Col = 8
            .Col2 = .Col
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            '.set_ColWidth(8, 1000)
            .TypeFloatDecimalPlaces = 2
            .TypeFloatMin = 0.0#
            .TypeFloatMax = 999999.99
            .Row = .MaxRows
            .Col = 9
            .Col2 = .Col
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeMaxEditLen = 50
            '.set_ColWidth(9, 1900)
            .Row = .MaxRows
            .Col = 10
            .Col2 = .Col
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
            '.set_ColWidth(10, 600)
            .TypeCheckCenter = True
            .Value = 1
            .ColsFrozen = 4
            .Row = .MaxRows
            .Col = 11
            .Col2 = .Col
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
            '.set_ColWidth(11, 1000)
            .TypeCheckCenter = True
            .Value = 1
            .ColsFrozen = 4
            .Row = .MaxRows
            .Col = 12
            .Col2 = .Col
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeMaxEditLen = 20
            '.set_ColWidth(12, 1500)
            .Row = .MaxRows
            .Col = 13
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = "..."
            '.set_ColWidth(13, 315)
            '.TypeButtonPicture = My.Resources.ico111.ToBitmap          '1234MILIND
            .Row = .MaxRows
            .Col = 14
            .Col2 = .Col
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeMaxEditLen = 20
            '.set_ColWidth(14, 1500)
            .Row = .MaxRows
            .Col = 15
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = "..."
            '.set_ColWidth(15, 315)
            '.TypeButtonPicture = My.Resources.ico111.ToBitmap          '1234MILIND
            .Row = .MaxRows
            .Col = 17
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = "..."
            '.set_ColWidth(17, 315)
            '.TypeButtonPicture = My.Resources.ico111.ToBitmap
            '.TypeButtonText = "VIEW"
            .Row = .MaxRows
            .Col = 20
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = "..."
            '.set_ColWidth(20, 1200)
            .Row = .MaxRows
            .Col = 21                        'check box for auto inv part
            .Col2 = .Col
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
            '.set_ColWidth(1, 520)
            .TypeCheckCenter = True
            If blnAllowBudget = False Then
                .Lock = True
            End If
            .Row = .MaxRows
            .Col = 22
            .Col2 = .Col
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeMaxEditLen = 4

            .Row = .MaxRows
            .Col = 23
            .Col2 = .Col
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeMaxEditLen = 10
            .Row = .MaxRows
            .Col = 24
            .Col2 = .Col
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeDate
            .Text = "01/01/1990"
            .Row = .MaxRows
            .Col = 25
            .Col2 = .Col
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeMaxEditLen = 50
            If Not (gstrUNITID = "MS1" OrElse gstrUNITID = "MS2" Or gstrUNITID = "MK1") Then
                If blnAllowBudget = False Then
                    .Lock = True
                End If
            End If

        End With
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    '    Private Sub addNewInSpread()
    '        On Error GoTo errHandler
    '        Dim intRowHeight As Short
    '        Dim varCurrency As Object
    '        With Me.spdITemDetails
    '            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
    '            .MaxRows = .MaxRows + 1
    '            .Row = .MaxRows
    '            .UnitType = FPSpreadADO.UnitTypeConstants.UnitTypeTwips
    '            .set_RowHeight(.Row, 300)
    '            If CmdGrpCustomerITem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
    '                Call .SetText(7, .MaxRows, "A")
    '            End If
    '            If .MaxRows > 6 Then .ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsBoth
    '        End With
    '        Call AddNewRowType()
    '        Exit Sub
    'errHandler:
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '    End Sub
    Public Function validbudgetData() As Boolean
        On Error GoTo errHandler
        Dim lstrControls As String
        Dim intMaxCount, intCount As Integer
        Dim StrItemCode As String
        Dim rstmp_budgetitem_mst As ClsResultSetDB
        Dim lNo, i As Integer
        Dim varFeild As Object
        validbudgetData = True
        lNo = 1
        lstrControls = ""
        lstrControls = ResolveResString(10059)
        With Me.spdITemDetails
            For i = 1 To spdITemDetails.MaxRows
                varFeild = Nothing
                Call .GetText(2, i, varFeild)
                If CheckForItemMainGroup(varFeild) = True Then
                    varFeild = Nothing
                    Call .GetText(12, i, varFeild) 'colour code not defined
                    If Len(Trim(varFeild)) > 0 Then
                        rstmp_budgetitem_mst = New ClsResultSetDB
                        rstmp_budgetitem_mst.GetResult("select * from Colour_mst where colour_code='" & varFeild & "' and active=1 AND UNIT_CODE='" & gstrUNITID & "'")
                        intMaxCount = rstmp_budgetitem_mst.RowCount
                        If intMaxCount = 0 Then
                            lstrControls = lstrControls & vbCrLf & lNo & ". Invalid Colour Code "
                            lNo = lNo + 1
                            validbudgetData = False
                        End If
                        rstmp_budgetitem_mst.ResultSetClose()
                        rstmp_budgetitem_mst = Nothing
                        varFeild = Nothing
                    End If
                    varFeild = Nothing
                    Call .GetText(14, i, varFeild) 'category code not defined
                    If Len(Trim(varFeild)) > 0 Then
                        rstmp_budgetitem_mst = New ClsResultSetDB
                        rstmp_budgetitem_mst.GetResult("select * from Colour_mst  where category='" & varFeild & "' and active=1 AND UNIT_CODE='" & gstrUNITID & "' ")
                        intMaxCount = rstmp_budgetitem_mst.RowCount
                        If intMaxCount = 0 Then
                            lstrControls = lstrControls & vbCrLf & lNo & ". Invalid Category Code "
                            lNo = lNo + 1
                            validbudgetData = False
                        End If
                        rstmp_budgetitem_mst.ResultSetClose()
                        rstmp_budgetitem_mst = Nothing
                        varFeild = Nothing
                    End If
                    varFeild = Nothing
                    Call .GetText(16, i, varFeild)
                    If Len(Trim(varFeild)) > 0 Then
                        rstmp_budgetitem_mst = New ClsResultSetDB
                        rstmp_budgetitem_mst.GetResult("select * from Commodity_mst  where commodity_code ='" & varFeild & "' and active=1 AND UNIT_CODE='" & gstrUNITID & "' ")
                        intMaxCount = rstmp_budgetitem_mst.RowCount
                        If intMaxCount = 0 Then
                            lstrControls = lstrControls & vbCrLf & lNo & ". Invalid commodity Code "
                            lNo = lNo + 1
                            validbudgetData = False
                        End If
                        rstmp_budgetitem_mst.ResultSetClose()
                        rstmp_budgetitem_mst = Nothing
                        varFeild = Nothing
                    End If
                    rstmp_budgetitem_mst = New ClsResultSetDB
                    rstmp_budgetitem_mst.GetResult("select * from tmp_budgetitem_mst where ip_address = '" & gstrIpaddressWinSck & "' AND UNIT_CODE='" & gstrUNITID & "'")
                    intMaxCount = rstmp_budgetitem_mst.RowCount
                    rstmp_budgetitem_mst.MoveFirst()
                    If intMaxCount = 0 Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". Model Code is not Defined "
                        validbudgetData = False
                    End If
                    rstmp_budgetitem_mst.ResultSetClose()
                    rstmp_budgetitem_mst = Nothing
                End If
                If validbudgetData = False Then
                    GoTo A
                End If
            Next
        End With
A:
        If validbudgetData = False Then
            MsgBox(lstrControls, MsgBoxStyle.Information, ResolveResString(10059))
        End If
        Exit Function
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function ValidateBeforeSave() As Boolean
        On Error GoTo errHandler
        Dim intLoopCounter As Short
        Dim intMaxCounter As Short
        Dim lstrControls As String
        Dim lNo As Integer
        Dim lctrFocus As System.Windows.Forms.Control
        Dim varFeild As Object
        Dim VarItemDRGNo As Object
        Dim rsAllowMultiParts As New ClsResultSetDB
        Dim intPlace As Integer
        Dim varEditFlag As Object
        ValidateBeforeSave = True
        lNo = 1
        lstrControls = ""
        lstrControls = ResolveResString(10059)
        If Len(Me.txtCustCode.Text) = 0 Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Customer Code "
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = Me.txtCustCode
            End If
            ValidateBeforeSave = False
        End If
        If spdITemDetails.MaxRows < 1 Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Add Atleast one Item. "
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = Me.txtCustCode
            End If
            ValidateBeforeSave = False
        End If
        With spdITemDetails
            varFeild = Nothing
            Call .GetText(2, spdITemDetails.MaxRows, varFeild)
            If Len(Trim(varFeild)) = 0 Then
                If spdITemDetails.MaxRows > 1 Then
                    spdITemDetails.MaxRows = spdITemDetails.MaxRows - 1
                    .Col = 2
                    .Row = .MaxRows
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                End If
            End If
            intMaxCounter = spdITemDetails.MaxRows
            For intLoopCounter = 1 To intMaxCounter
                varFeild = Nothing
                Call .GetText(2, intLoopCounter, varFeild)
                If Len(Trim(varFeild)) = 0 Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Internal Part Code in Row - " & intLoopCounter
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = spdITemDetails
                        With spdITemDetails
                            .Row = intLoopCounter
                            .Col = 2
                        End With
                    End If
                    ValidateBeforeSave = False
                End If
                varFeild = Nothing
                Call .GetText(5, intLoopCounter, varFeild)
                If Len(Trim(varFeild)) > 0 Then
                    intPlace = InStr(1, varFeild, "'")
                    If Val(intPlace) > 0 Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". Character ' is not allowed in Customer Part No. in Row - " & intLoopCounter
                        lNo = lNo + 1
                        If lctrFocus Is Nothing Then
                            lctrFocus = spdITemDetails
                            With spdITemDetails
                                .Row = intLoopCounter
                                .Col = 5
                            End With
                        End If
                        ValidateBeforeSave = False
                    End If
                End If
                varFeild = Nothing
                Call .GetText(9, intLoopCounter, varFeild)
                If Len(Trim(varFeild)) > 0 Then
                    intPlace = InStr(1, varFeild, "'")
                    If Val(intPlace) > 0 Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". Character ' is not allowed Model Name in Row - " & intLoopCounter
                        lNo = lNo + 1
                        If lctrFocus Is Nothing Then
                            lctrFocus = spdITemDetails
                            With spdITemDetails
                                .Row = intLoopCounter
                                .Col = 9
                            End With
                        End If
                        ValidateBeforeSave = False
                    End If
                End If
                varFeild = Nothing
                Call .GetText(5, intLoopCounter, varFeild)
                If Len(Trim(varFeild)) = 0 Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Customer Part Code in Row - " & intLoopCounter
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = spdITemDetails
                        With spdITemDetails
                            .Row = intLoopCounter
                            .Col = 5
                        End With
                    End If
                    ValidateBeforeSave = False
                End If
                varFeild = Nothing
                Call .GetText(6, intLoopCounter, varFeild)
                If Len(Trim(varFeild)) = 0 Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Customer Part Description in Row - " & intLoopCounter
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = spdITemDetails
                        With spdITemDetails
                            .Row = intLoopCounter
                            .Col = 6
                        End With
                    End If
                    ValidateBeforeSave = False
                End If
                varFeild = Nothing
                Call .GetText(2, intLoopCounter, varFeild)
                VarItemDRGNo = Nothing
                Call .GetText(5, intLoopCounter, VarItemDRGNo)
                varEditFlag = Nothing
                Call .GetText(7, intLoopCounter, varEditFlag)
                If CmdGrpCustomerITem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                    If varEditFlag = "A" Then
                        If IsExistFieldValue(CStr(Trim(VarItemDRGNo)), "Cust_DrgNo", "CustItem_Mst", "ITem_Code = '" & Trim(varFeild) & "' and Account_code = '" & Trim(txtCustCode.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'") = True Then
                            lstrControls = lstrControls & vbCrLf & lNo & ". This Drawing No Already Exist for this Item" & varFeild
                            lNo = lNo + 1
                            If lctrFocus Is Nothing Then
                                lctrFocus = spdITemDetails
                                With spdITemDetails
                                    .Row = intLoopCounter
                                    .Col = 5
                                End With
                            End If
                            ValidateBeforeSave = False
                            Exit For
                        End If
                        'Praveen on 10 OCT 2017
                        If blnAllowMultiPartNo = True Then
                            If UCase(Trim(GetPlantName)) = "HILEX" Then
                                ValidateBeforeSave = Not funSomeItemsAlreadyinCustItemMSt()
                                If ValidateBeforeSave = False Then
                                    Dim result As DialogResult = MessageBox.Show("Some Internal Part Nos/CustDrgNo already linked with other CustDrgNo/Internal Part Nos for this customer. Would you like to continue..?", "Confirmation", MessageBoxButtons.YesNoCancel)
                                    If result = DialogResult.Yes Then
                                        ValidateBeforeSave = True
                                    ElseIf result = DialogResult.No Then
                                        ValidateBeforeSave = False
                                        lctrFocus = spdITemDetails
                                        lstrControls = lstrControls & vbCrLf & lNo & " some Internal Part Nos/CustDrgNo already linked with other CustDrgNo/Internal Part Nos for this customer "
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    If IsExistFieldValue(CStr(Trim(VarItemDRGNo)), "Cust_DrgNo", "CustItem_Mst", "ITem_Code = '" & Trim(varFeild) & "' and Account_code = '" & Trim(txtCustCode.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'") = True Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". This Drawing No Already Exist for this Item" & varFeild
                        lNo = lNo + 1
                        If lctrFocus Is Nothing Then
                            lctrFocus = spdITemDetails
                            With spdITemDetails
                                .Row = intLoopCounter
                                .Col = 5
                            End With
                        End If
                        ValidateBeforeSave = False
                        Exit For
                    End If
                End If
            Next
        End With
        rsAllowMultiParts.GetResult("SELECT AllowMultiplePartCodes from Sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If rsAllowMultiParts.RowCount > 0 Then
            blnAllowMultiPartNo = IIf(rsAllowMultiParts.GetValue("AllowMultiplePartCodes") = True, True, False)
        End If
        rsAllowMultiParts.ResultSetClose()
        If blnAllowMultiPartNo = False Then
            ValidateBeforeSave = Not funDuplicateItems()
            If ValidateBeforeSave = False Then
                lctrFocus = spdITemDetails
                lstrControls = lstrControls & vbCrLf & lNo & " Duplicate Internal Part Nos "
                MsgBox(lstrControls, MsgBoxStyle.Information, ResolveResString(10059))
                gblnCancelUnload = True
                If TypeOf lctrFocus Is System.Windows.Forms.TextBox Then
                    lctrFocus.Focus()
                Else
                    DirectCast(lctrFocus, AxFPSpreadADO.AxfpSpread).Action = FPSpreadADO.ActionConstants.ActionActiveCell
                End If
                Exit Function
            End If
        End If
        If CmdGrpCustomerITem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            If blnAllowMultiPartNo = False Then
                ValidateBeforeSave = Not funAlreadyinCustItemMSt()
                If ValidateBeforeSave = False Then
                    lctrFocus = spdITemDetails
                    lstrControls = lstrControls & vbCrLf & lNo & " some Internal Part Nos already defined for this customer "
                End If
            Else
                'Praveen on 05OCT 2017
                If UCase(Trim(GetPlantName)) = "HILEX" Then
                    ValidateBeforeSave = Not funSomeItemsAlreadyinCustItemMSt()
                    If ValidateBeforeSave = False Then
                        Dim result As DialogResult = MessageBox.Show("Some Internal Part Nos/CustDrgNo already linked with other CustDrgNo/Internal Part Nos for this customer. Would you like to continue..?", "Confirmation", MessageBoxButtons.YesNoCancel)
                        If result = DialogResult.Yes Then
                            ValidateBeforeSave = True
                        ElseIf result = DialogResult.No Then
                            ValidateBeforeSave = False
                            lctrFocus = spdITemDetails
                            lstrControls = lstrControls & vbCrLf & lNo & " some Internal Part Nos/CustDrgNo already linked with other CustDrgNo/Internal Part Nos for this customer "
                        End If
                    End If
                End If

            End If
        End If

        'Incident # INC1222839 24 Jun 2025
        If (gstrUNITID = "MS1" OrElse gstrUNITID = "MS2" Or gstrUNITID = "MK1") AndAlso (CmdGrpCustomerITem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD OrElse CmdGrpCustomerITem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT) Then
            Dim DeclarationNo As Object
            For intLoop = 1 To spdITemDetails.MaxRows Step 1

                spdITemDetails.Col = 2
                spdITemDetails.Row = intLoop

                Dim IsEdit As Boolean = Convert.ToBoolean(spdITemDetails.CellTag)

                If IsEdit Then
                    Continue For
                End If

                DeclarationNo = Nothing
                spdITemDetails.GetText(25, intLoop, DeclarationNo)
                If DeclarationNo Is Nothing OrElse DeclarationNo.ToString().Length = 0 Then
                    ValidateBeforeSave = False
                    lctrFocus = spdITemDetails
                    lstrControls = lstrControls & vbCrLf & lNo & " Declaration No field can not be blank row : " & intLoop
                    Exit For
                End If
            Next
        End If

        If Not ValidateBeforeSave Then
            MsgBox(lstrControls, MsgBoxStyle.Information, ResolveResString(10059))
            gblnCancelUnload = True
            If TypeOf lctrFocus Is System.Windows.Forms.TextBox Then
                lctrFocus.Focus()
            Else
                DirectCast(lctrFocus, AxFPSpreadADO.AxfpSpread).Action = FPSpreadADO.ActionConstants.ActionActiveCell
            End If
        End If
        Exit Function
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function funDuplicateItems() As Boolean
        '--------------------------------------------------
        'Created by         -       Brij B Bohara
        'Arguments          -       None
        'Purpose            -       To check Duplicate items in grid
        'Output             -       True/False
        '--------------------------------------------------
        Dim intInnerLoop As Short 'For inner loop
        Dim intOuterLoop As Short 'For outer loop
        Dim StrItemCode As Object 'To store item code
        Dim strCurItem As Object 'To store Current Item code
        On Error GoTo errHandler
        'To check item in grid
        For intOuterLoop = 1 To spdITemDetails.MaxRows - 1 Step 1 'Outer loop for picking the item code
            StrItemCode = Nothing
            spdITemDetails.GetText(2, intOuterLoop, StrItemCode)
            For intInnerLoop = intOuterLoop + 1 To spdITemDetails.MaxRows Step 1 'Inner loop for checking the existence of  the item code
                strCurItem = Nothing
                spdITemDetails.GetText(2, intInnerLoop, strCurItem)
                If StrComp(strCurItem, StrItemCode, CompareMethod.Text) = 0 Then
                    funDuplicateItems = True
                    Exit Function
                End If
            Next intInnerLoop
        Next intOuterLoop
        funDuplicateItems = False
        Exit Function 'To avoid the error handler execution
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function funSomeItemsAlreadyinCustItemMSt() As Boolean
        '--------------------------------------------------
        'Created by         -       Praveen Kumar
        'Arguments          -       None
        'Purpose            -       To check more than one relations between itemcode and custDrgNo
        'Output             -       True/False
        '--------------------------------------------------
        Dim intLoop As Short 'For outer loop
        Dim StrItemCode As Object 'To store item code
        Dim strCurItem As Object 'To store Current Item code
        Dim varEditFlag As Object
        Dim rsCustItemMst As ClsResultSetDB
        On Error GoTo errHandler

        rsCustItemMst = New ClsResultSetDB

        For intLoop = 1 To spdITemDetails.MaxRows Step 1 'Outer loop for picking the item code
            varEditFlag = Nothing
            spdITemDetails.GetText(7, intLoop, varEditFlag)
            If (varEditFlag = "A" And CmdGrpCustomerITem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT) Or CmdGrpCustomerITem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                StrItemCode = Nothing
                strCurItem = Nothing
                spdITemDetails.GetText(2, intLoop, StrItemCode)
                spdITemDetails.GetText(5, intLoop, strCurItem)
                rsCustItemMst.GetResult("Select * from CustItem_Mst where UNIT_CODE='" & gstrUNITID & "' AND  (Item_Code='" & StrItemCode & "' or Cust_Drgno = '" & strCurItem & "') AND Active=1 and Account_code = '" & txtCustCode.Text & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rsCustItemMst.GetNoRows > 0 Then
                    funSomeItemsAlreadyinCustItemMSt = True
                    Exit Function
                End If
            End If
        Next intLoop
        rsCustItemMst.ResultSetClose()
        funSomeItemsAlreadyinCustItemMSt = False
        Exit Function 'To avoid the error handler execution
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function funAlreadyinCustItemMSt() As Boolean
        '--------------------------------------------------
        'Created by         -       Brij B Bohara
        'Arguments          -       None
        'Purpose            -       To check Duplicate items in grid
        'Output             -       True/False
        '--------------------------------------------------
        Dim intLoop As Short 'For outer loop
        Dim StrItemCode As Object 'To store item code
        Dim strCurItem As Object 'To store Current Item code
        Dim rsCustItemMst As ClsResultSetDB
        On Error GoTo errHandler
        'To check item in CustItem_mst
        rsCustItemMst = New ClsResultSetDB
        rsCustItemMst.GetResult("Select * from CustItem_Mst where UNIT_CODE='" & gstrUNITID & "' AND Account_code = '" & txtCustCode.Text & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If rsCustItemMst.GetNoRows > 0 Then
            rsCustItemMst.MoveFirst()
            While Not rsCustItemMst.EOFRecord
                strCurItem = rsCustItemMst.GetValue("Item_code")
                For intLoop = 1 To spdITemDetails.MaxRows Step 1 'Outer loop for picking the item code
                    StrItemCode = Nothing
                    spdITemDetails.GetText(2, intLoop, StrItemCode)
                    If StrComp(strCurItem, StrItemCode, CompareMethod.Text) = 0 Then
                        funAlreadyinCustItemMSt = True
                        Exit Function
                    End If
                Next intLoop
                rsCustItemMst.MoveNext()
            End While
        End If
        rsCustItemMst.ResultSetClose()
        funAlreadyinCustItemMSt = False
        Exit Function 'To avoid the error handler execution
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function ValidRowData(ByRef intRow As Integer) As Boolean
        On Error GoTo errHandler
        Dim lstrControls As String
        Dim lNo As Integer
        Dim lctrFocus As System.Windows.Forms.Control
        Dim varFeild As Object
        Dim varItemCode As Object
        ValidRowData = True
        lNo = 1
        lstrControls = ResolveResString(10059)
        '***To Check Blank ItemCode
        Dim VartempDrgNo As Object
        Dim VartempItemCode As Object
        Dim intMaxCounter As Short
        Dim intLoopCounter As Short
        With spdITemDetails
            varFeild = Nothing
            Call .GetText(2, intRow, varFeild)
            If Len(Trim(varFeild)) = 0 Then
                lstrControls = lstrControls & vbCrLf & lNo & ". Internal Part Code in Row - " & intRow
                lNo = lNo + 1
                If lctrFocus Is Nothing Then
                    lctrFocus = spdITemDetails
                    With spdITemDetails
                        .Row = intRow
                        .Col = 2
                    End With
                End If
                ValidRowData = False
            End If
            varFeild = Nothing
            Call .GetText(5, intRow, varFeild)
            If Len(Trim(varFeild)) = 0 Then
                lstrControls = lstrControls & vbCrLf & lNo & ". Customer Part Code in Row - " & intRow
                lNo = lNo + 1
                If lctrFocus Is Nothing Then
                    lctrFocus = spdITemDetails
                    With spdITemDetails
                        .Row = intRow
                        .Col = 5
                    End With
                End If
                ValidRowData = False
            End If
            varFeild = Nothing
            Call .GetText(6, intRow, varFeild)
            If Len(Trim(varFeild)) = 0 Then
                lstrControls = lstrControls & vbCrLf & lNo & ". Customer Part Description in Row - " & intRow
                lNo = lNo + 1
                If lctrFocus Is Nothing Then
                    lctrFocus = spdITemDetails
                    With spdITemDetails
                        .Row = intRow
                        .Col = 6
                    End With
                End If
                ValidRowData = False
            End If
            intMaxCounter = intRow
            VartempDrgNo = Nothing
            Call .GetText(5, intRow, VartempDrgNo)
            VartempItemCode = Nothing
            Call .GetText(2, intRow, VartempItemCode)
            For intLoopCounter = 1 To intRow
                If intLoopCounter <> intRow Then
                    varFeild = Nothing
                    Call .GetText(5, intLoopCounter, varFeild)
                    varItemCode = Nothing
                    Call .GetText(2, intLoopCounter, varItemCode)
                    If (UCase(Trim(varItemCode)) = UCase(Trim(VartempItemCode))) And (UCase(Trim(varFeild)) = UCase(Trim(VartempDrgNo))) Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". This Combination of Internal Part Code and Customer Part Code already Exists."
                        lNo = lNo + 1
                        If lctrFocus Is Nothing Then
                            lctrFocus = spdITemDetails
                            With spdITemDetails
                                .Row = intRow
                                .Col = 5
                            End With
                        End If
                        ValidRowData = False
                    End If
                End If
            Next
            '****
        End With
        If Not ValidRowData Then
            MsgBox(lstrControls, MsgBoxStyle.Information, ResolveResString(10059))
            If TypeOf lctrFocus Is System.Windows.Forms.TextBox Then
                lctrFocus.Focus()
            Else
                lctrFocus.Focus()
                DirectCast(lctrFocus, AxFPSpreadADO.AxfpSpread).Action = FPSpreadADO.ActionConstants.ActionActiveCell
            End If
        End If
        Exit Function
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Public Sub InsertData(ByRef intRow As Integer)
        On Error GoTo errHandler
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim varDesc As Object
        Dim varUOM As Object
        Dim varTariffCode As Object
        Dim varCurrency As Object
        Dim varPrice As Object
        Dim varToolCost As Object
        Dim varCustMtrl As Object
        Dim varPacking As Object
        Dim VarBinQuantity As Object
        Dim VarModel As Object
        Dim varItemDesc As Object
        Dim varCommodity As Object = Nothing
        Dim intUploadFlag As Short
        Dim varContainertype As Object = Nothing
        Dim varPackingCode As Object = Nothing
        Dim intPackingCode As Short
        Dim varAutoInvPart As Object
        Dim varSHOPCode As Object = Nothing
        Dim varGateNo As Object = Nothing
        Dim vardate As Object = Nothing
        Dim decl_no As Object = Nothing
        With spdITemDetails
            varItemCode = Nothing
            Call .GetText(2, intRow, varItemCode)
            varItemDesc = Nothing
            Call .GetText(4, intRow, varItemDesc)
            varDrgNo = Nothing
            Call .GetText(5, intRow, varDrgNo)
            varDesc = Nothing
            Call .GetText(6, intRow, varDesc)
            VarBinQuantity = Nothing
            Call .GetText(8, intRow, VarBinQuantity)
            VarModel = Nothing
            Call .GetText(9, intRow, VarModel)
            .Row = intRow
            .Col = 11
            If .Value = System.Windows.Forms.CheckState.Checked Then
                intUploadFlag = 1
            Else
                intUploadFlag = 0
            End If
            varCommodity = Nothing
            Call .GetText(12, intRow, varCommodity)
            varContainertype = Nothing
            Call .GetText(18, intRow, varContainertype)
            varPackingCode = Nothing
            Call .GetText(19, intRow, varPackingCode)
            If varPackingCode.ToString.Length > 0 Then
                intPackingCode = Val(varPackingCode.ToString.Substring(0, 1))
            End If
            varAutoInvPart = Nothing                     'to update Item_mst if auto inv part is checked
            Call .GetText(21, intRow, varAutoInvPart)
            If varAutoInvPart = "1" Then
                strInsert = Trim(strInsert) & " Update Item_mst set AUTO_INVOICE_PART = 1 where Unit_Code='" & gstrUNITID & "' and Item_code ='" & varItemCode & "' "
            End If
            varSHOPCode = Nothing
            Call .GetText(22, intRow, varSHOPCode)

            varGateNo = Nothing
            Call .GetText(23, intRow, varGateNo)
            vardate = Nothing
            Call .GetText(24, intRow, vardate)
            decl_no = Nothing
            Call .GetText(25, intRow, decl_no)
        End With
        If Len(Trim(strInsert)) > 0 Then
            strInsert = Trim(strInsert) & " Insert into CustItem_Mst(Account_code,Cust_DrgNo,Drg_Desc,Item_Code,BinQuantity,VarModel,"
        Else
            strInsert = " Insert into CustItem_Mst(Account_code,Cust_DrgNo,Drg_Desc,Item_Code,BinQuantity,VarModel,"
        End If
        strInsert = Trim(strInsert) & "Ent_Dt,"
        strInsert = Trim(strInsert) & "Ent_UserId,upd_Dt,Upd_userId, Item_Desc,commodity,schupldreqd,container,packing_code,UNIT_CODE,SHOP_CODE,GATE_NO,Decl_No) Values ("
        strInsert = Trim(strInsert) & "'" & Trim(txtCustCode.Text) & "','" & varDrgNo & "','" & varDesc & "','" & varItemCode & "'," & Val(VarBinQuantity) & ",'" & VarModel & "','" & getDateForDB(GetServerDate()) & "' ,'" & mP_User & "','" & getDateForDB(GetServerDate()) & "' ,'"
        strInsert = Trim(strInsert) & mP_User & "', '" & Replace(Trim(varItemDesc), "'", "") & "','" & varCommodity & "'," & intUploadFlag & ",'" & varContainertype & "'," & intPackingCode & ",'" & gstrUNITID & "','" & varSHOPCode & "','" & varGateNo & "','" & decl_no & "')"
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Function UpdateData(ByRef intRow As Integer) As Object
        On Error GoTo errHandler
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim varDesc As Object
        Dim varUOM As Object
        Dim varTariffCode As Object
        Dim varCurrency As Object
        Dim varPrice As Object
        Dim varToolCost As Object
        Dim varCustMtrl As Object
        Dim varPacking As Object
        Dim VarBinQuantity As Object
        Dim varItemDesc As Object
        Dim VarModel As Object
        Dim varActive As Object = Nothing
        Dim varAutoInvPart As Object = Nothing
        Dim varUploadFlag As Object = Nothing
        Dim varColour As Object = Nothing
        Dim varCategory As Object = Nothing
        Dim varCommodity As Object = Nothing
        Dim varContainertype As Object = Nothing
        Dim varPackingCode As Object = Nothing
        Dim varSHOPCode As Object = Nothing
        Dim intPackingCode As Short
        Dim strInsQuery As String = ""
        Dim strSql As String = ""
        Dim varGateno As Object = Nothing
        Dim decl_no As Object = Nothing

        With spdITemDetails
            varItemCode = Nothing
            Call .GetText(2, intRow, varItemCode)
            varItemDesc = Nothing
            Call .GetText(4, intRow, varItemDesc)
            varDrgNo = Nothing
            Call .GetText(5, intRow, varDrgNo)
            varDesc = Nothing
            Call .GetText(6, intRow, varDesc)
            VarBinQuantity = Nothing
            Call .GetText(8, intRow, VarBinQuantity)
            VarModel = Nothing
            Call .GetText(9, intRow, VarModel)
            varActive = Nothing
            Call .GetText(10, intRow, varActive)
            If varActive = "0" Then
                varActive = "0"
            Else
                varActive = "1"
            End If
            varAutoInvPart = Nothing                        'to update item_mst if auto inv part is checked in edit mode
            Call .GetText(21, intRow, varAutoInvPart)
            If varAutoInvPart = "1" Then
                strSql = " Update Item_mst set AUTO_INVOICE_PART = 1 where Unit_Code='" & gstrUNITID & "' and Item_code ='" & varItemCode & "' "
                mP_Connection.Execute(strSql, ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If

            'ends here 
            varUploadFlag = Nothing
            Call .GetText(11, intRow, varUploadFlag)
            If varUploadFlag = "0" Then
                varUploadFlag = "0"
            Else
                varUploadFlag = "1"
            End If
            varColour = Nothing
            Call .GetText(12, intRow, varColour)
            varCategory = Nothing
            Call .GetText(14, intRow, varCategory)
            varCommodity = Nothing
            Call .GetText(16, intRow, varCommodity)
            varContainertype = Nothing
            Call .GetText(18, intRow, varContainertype)
            varPackingCode = Nothing
            Call .GetText(19, intRow, varPackingCode)
            If varPackingCode.ToString.Length > 0 Then
                intPackingCode = varPackingCode.ToString.Substring(0, 1)
            End If
            varSHOPCode = Nothing
            Call .GetText(22, intRow, varSHOPCode)

            varGateno = Nothing
            Call .GetText(23, intRow, varGateno)
            decl_no = Nothing
            Call .GetText(25, intRow, decl_no)
        End With

        'If Len(Trim(strupdate)) > 0 Then
        '    strupdate = Trim(strupdate) & vbCrLf & "Update CustItem_Mst set Drg_Desc = '" & varDesc & "',VarModel='" & VarModel & "',Active = '" & varActive & "', schupldreqd = '" & varUploadFlag & "',commodity='" & varCommodity & "',container='" & varContainertype & "',packing_code=" & intPackingCode & ","
        'Else
        '    strupdate = "set dateformat 'dmy' Update CustItem_Mst set Drg_Desc = '" & varDesc & "',VarModel='" & VarModel & "',Active = '" & varActive & "', schupldreqd = '" & varUploadFlag & "',commodity='" & varCommodity & "',container='" & varContainertype & "',packing_code=" & intPackingCode & ","
        'End If
        'strupdate = Trim(strupdate) & "Upd_dt = '" & getDateForDB(GetServerDate()) & "', Upd_Userid = '" & mP_User & "', BinQuantity = " & Val(VarBinQuantity) & ", Item_Desc = '" & Replace(Trim(varItemDesc), "'", "") & "'  where account_code ='"
        'strupdate = Trim(strupdate) & Trim(txtCustCode.Text) & "' and Item_code ='" & varItemCode & "' and Cust_DrgNo ='" & varDrgNo & "' AND UNIT_CODE='" & gstrUNITID & "'"

        'ISSUE ID       : 10222401 
        strInsQuery = "Insert into TMPCustItem_Mst(Account_code,Cust_DrgNo,Drg_Desc,Item_Code,Active,BinQuantity,VarModel,Item_Desc,commodity,schupldreqd,container,packing_code,UNIT_CODE,IP_ADDRESS,SHOP_CODE,GATE_NO,DECL_NO)"
        strInsQuery = strInsQuery.Trim & " values('" & txtCustCode.Text.Trim & "','" & varDrgNo & "','" & varDesc & "','" & varItemCode & "'," & varActive & "," & Val(VarBinQuantity) & ",'" & VarModel & "','" & Replace(Trim(varItemDesc), "'", "") & "','" & varCommodity & "','" & varUploadFlag & "','" & varContainertype & "'," & intPackingCode & ",'" & gstrUNITID & "','" & gstrIpaddressWinSck & "','" & varSHOPCode & "','" & varGateno & "','" & decl_no & "' )"

        mP_Connection.Execute(strInsQuery, ADODB.ExecuteOptionEnum.adExecuteNoRecords)

        Exit Function
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function UpdateBudgetData(ByRef intRow As Integer) As Object
        On Error GoTo errHandler
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim varDesc As Object
        Dim VarBinQuantity As Object
        Dim varItemDesc As Object
        Dim VarModel As Object
        Dim varActive As Object = Nothing
        Dim varUploadFlag As Object = Nothing
        Dim varColour As Object = Nothing
        Dim varCategory As Object = Nothing
        Dim varCommodity As Object = Nothing
        Dim varContainertype As Object = Nothing
        Dim rsbudgetdata As ClsResultSetDB
        With spdITemDetails
            varItemCode = Nothing
            Call .GetText(2, intRow, varItemCode)
            varItemDesc = Nothing
            Call .GetText(4, intRow, varItemDesc)
            varDrgNo = Nothing
            Call .GetText(5, intRow, varDrgNo)
            varDesc = Nothing
            Call .GetText(6, intRow, varDesc)
            VarBinQuantity = Nothing
            Call .GetText(8, intRow, VarBinQuantity)
            VarModel = Nothing
            Call .GetText(9, intRow, VarModel)
            varActive = Nothing
            Call .GetText(10, intRow, varActive)
            If varActive = "0" Then
                varActive = "0"
            Else
                varActive = "1"
            End If
            varUploadFlag = Nothing
            Call .GetText(11, intRow, varUploadFlag)
            If varUploadFlag = "0" Then
                varUploadFlag = "0"
            Else
                varUploadFlag = "1"
            End If
            varColour = Nothing
            Call .GetText(12, intRow, varColour)
            varCategory = Nothing
            Call .GetText(14, intRow, varCategory)
            varCommodity = Nothing
            Call .GetText(16, intRow, varCommodity)
            varContainertype = Nothing
            Call .GetText(18, intRow, varContainertype)
        End With
        rsbudgetdata = New ClsResultSetDB
        rsbudgetdata.GetResult("select ACCOUNT_CODE,CUST_DRGNO,ITEM_CODE,COLOUR_CODE,CATEGORY_CODE,COMMODITY_CODE,MODEL_CODE,VARIANT_CODE,USAGE_QTY from budgetitem_mst where account_code ='" & txtCustCode.Text.Trim & "' and item_code='" & varItemCode & "' and cust_drgno='" & varDrgNo & "' AND UNIT_CODE='" & gstrUNITID & "'")
        If Len(Trim(strDeletebudget)) <= 0 Then 'very First Time insertion
            strDeletebudget = "Delete from budgetitem_mst where account_code ='" & txtCustCode.Text.Trim & "' and item_code='" & varItemCode & "' and cust_drgno='" & varDrgNo & "' AND UNIT_CODE='" & gstrUNITID & "'"
        Else
            strDeletebudget = Trim(strDeletebudget) & vbCrLf & "Delete from budgetitem_mst where account_code ='" & txtCustCode.Text.Trim & "' and item_code='" & varItemCode & "' and cust_drgno='" & varDrgNo & "' AND UNIT_CODE='" & gstrUNITID & "'"
        End If
        rsbudgetdata.ResultSetClose()
        rsbudgetdata = Nothing
        Exit Function
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub CheckUsedItems()

        On Error GoTo errHandler
        Dim varItem_Code As Object
        Dim varCustDrgNo As Object
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim intLoopCounter As Short
        Dim intmaxrows As Short
        Dim objVal As Object = Nothing

        intmaxrows = spdITemDetails.MaxRows
        With spdITemDetails
            For intLoopCounter = 1 To intmaxrows
                varItem_Code = Nothing
                Call .GetText(2, intLoopCounter, varItem_Code)
                varCustDrgNo = Nothing
                Call .GetText(5, intLoopCounter, varCustDrgNo)
                rsCustOrdDtl = New ClsResultSetDB
                rsCustOrdDtl.GetResult("Select TOP 1 1 from Cust_ord_dtl where Account_code = '" & txtCustCode.Text & "' and ITem_code = '" & varItem_Code & "' and Cust_drgNo = '" & varCustDrgNo & "' AND UNIT_CODE='" & gstrUNITID & "'")
                If rsCustOrdDtl.GetNoRows >= 1 Then
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = intLoopCounter
                    .Row2 = intLoopCounter
                    .BlockMode = True : .ForeColor = System.Drawing.Color.Red
                    .BlockMode = False
                    .Col = 1
                    .Col2 = 1
                    .Row = intLoopCounter
                    .Row2 = intLoopCounter
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False

                    objVal = Nothing                                      'lock auto inv part checked column
                    .Row = intLoopCounter : .Col = 21 : objVal = .Value
                    If objVal = True Then
                        .Row = intLoopCounter : .Row2 = intLoopCounter
                        .Col = 21 : .Col2 = 21
                        .BlockMode = True : .Lock = True : .BlockMode = False
                    End If                                                ' ends here 
                Else
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = intLoopCounter
                    .Row2 = intLoopCounter
                    .BlockMode = True : .ForeColor = System.Drawing.Color.Black
                    .BlockMode = False
                    .Col = 1
                    .Col2 = 1
                    .Row = intLoopCounter
                    .Row2 = intLoopCounter
                    .BlockMode = True
                    .Lock = False
                    .BlockMode = False

                    .Col = 6

                    .Col2 = .MaxCols - 1
                    .Row = intLoopCounter
                    .Row2 = intLoopCounter
                    .BlockMode = True
                    .Lock = False
                    .BlockMode = False
                End If
                rsCustOrdDtl.ResultSetClose()
            Next
        End With
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub DeleteRecord(ByRef pintRow As Integer)
        On Error GoTo errHandler
        Dim varItemCode As Object
        Dim varDrgNo As Object
        With spdITemDetails
            varItemCode = Nothing
            Call .GetText(2, pintRow, varItemCode)
            varDrgNo = Nothing
            Call .GetText(5, pintRow, varDrgNo)
            If Len(Trim(strDelete)) > 0 Then
                strDelete = Trim(strDelete) & vbCrLf & "Delete CustItem_Mst Where Account_code = '" & Trim(txtCustCode.Text) & "'"
                strDelete = Trim(strDelete) & " and Item_Code ='" & varItemCode & "' and Cust_drgNo ='" & varDrgNo & "' AND UNIT_CODE='" & gstrUNITID & "'"
            Else
                strDelete = Trim(strDelete) & "Delete CustItem_Mst Where Account_code = '" & Trim(txtCustCode.Text) & "'"
                strDelete = Trim(strDelete) & " and Item_Code ='" & varItemCode & "' and Cust_drgNo ='" & varDrgNo & "' AND UNIT_CODE='" & gstrUNITID & "'"
            End If
        End With
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub CheckDeleteMark()
        On Error GoTo errHandler
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim VarDelete As Object
        blnCheckDelItems = False
        intMaxLoop = spdITemDetails.MaxRows
        For intLoopCounter = 1 To intMaxLoop
            VarDelete = Nothing
            Call spdITemDetails.GetText(1, intLoopCounter, VarDelete)
            If VarDelete = "1" Then
                blnCheckDelItems = True
            End If
        Next
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function AlreadyExist(ByRef pstrCustCode As String, ByRef pstrInternalCode As String, ByRef plngRow As Integer) As Boolean
        '----------------------------------------------------------------------------------------------------
        'Created by         -       Brij B Bohara
        'Arguments          -       Customer code and Selected Internal Code
        'Purpose            -       To check Duplicate Internal Item code selected
        'Output             -       True/False
        '                       -       Changes Done by nisha on 21 march added new Parameter plngRow to not to
        '                      check duplicacy in same row
        '----------------------------------------------------------------------------------------------------
        On Error GoTo errHandler
        Dim rsCustItem As ClsResultSetDB
        Dim intItems As Short
        Dim varItemCode As Object
        'Already in Customer Item Master
        rsCustItem = New ClsResultSetDB
        rsCustItem.GetResult("SELECT Cust_Drgno From CustItem_Mst WHERE Account_Code='" & Trim(pstrCustCode) & "' AND Item_code='" & Trim(pstrInternalCode) & "' AND UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsCustItem.EOFRecord = True Then
            AlreadyExist = False
        Else
            AlreadyExist = True
            Exit Function
        End If
        rsCustItem.ResultSetClose()
        'Already selected in Grid
        With spdITemDetails
            If spdITemDetails.MaxRows > 0 Then
                'Creating the string of already used Items in the grid
                For intItems = 1 To spdITemDetails.MaxRows
                    varItemCode = Nothing
                    spdITemDetails.GetText(2, intItems, varItemCode)
                    If intItems <> plngRow Then
                        If StrComp(Trim(pstrInternalCode), Trim(varItemCode), CompareMethod.Text) = 0 Then
                            AlreadyExist = True
                            Exit Function
                        End If
                    End If
                Next
            End If
        End With
        Exit Function
errHandler:
        rsCustItem.ResultSetClose()
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub CmdGrpCustomerITem_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles CmdGrpCustomerITem.ButtonClick
        On Error GoTo errHandler
        Dim intLoopCounter As Integer
        Dim intmaxrows As Short
        Dim varMode As Object
        Dim varDeleteFlag As Object
        Dim intcnt As Integer
        Dim intMaxCount As Integer
        Dim intCount As Integer
        Dim varColour As Object
        Dim varcategory As Object
        Dim varcommodity As Object
        Dim varItemcode As Object
        Dim varCustdrgno As Object
        Dim varDate As Object
        Dim rsbudgetmodel As ClsResultSetDB
        Dim str_Renamed As String
        Dim strupdateCustItemMst As String

        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                spdITemDetails.Col = 1
                spdITemDetails.Col2 = 1
                spdITemDetails.BlockMode = True
                spdITemDetails.ColHidden = True
                spdITemDetails.BlockMode = False
                txtCustCode.Enabled = True : txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtCustCode.Text = ""
                txtItemCode.Text = ""
                txtItemDesc.Text = ""
                Call GetPackingLevelCode()
                mP_Connection.Execute("delete from tmp_budgetitem_mst where ip_address = '" & gstrIpaddressWinSck & "' AND UNIT_CODE='" & gstrUNITID & "'", ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                txtCustCode.Focus()
                With spdITemDetails
                    .Col = 2 : .Lock = True
                    .Col = 9 : .Lock = True
                End With
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE

                Select Case CmdGrpCustomerITem.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Not ValidateBeforeSave() = False Then
                            '10737738
                            If ValidateVEHBOM() = False Then
                                Exit Sub
                            End If
                            With spdITemDetails
                                .Col = 1
                                .Col2 = 1
                                .BlockMode = True
                                .ColHidden = True
                                .BlockMode = False
                            End With
                            intmaxrows = spdITemDetails.MaxRows
                            strInsert = ""
                            For intLoopCounter = 1 To intmaxrows
                                If ValidRowData(intLoopCounter) = True Then
                                    Call InsertData(intLoopCounter)
                                Else
                                    Exit Sub
                                End If
                            Next
                            Dim srtDELOldData As String = String.Empty
                            strInsertbudget = ""
                            If blnAllowBudget = True Then
                                If validbudgetData() = True Then
                                    For intLoopCounter = 1 To intmaxrows

                                        varItemcode = Nothing
                                        Call spdITemDetails.GetText(2, intLoopCounter, varItemcode)

                                        varCustdrgno = Nothing
                                        Call spdITemDetails.GetText(5, intLoopCounter, varCustdrgno)
                                        varDate = Nothing
                                        spdITemDetails.GetText(24, intLoopCounter, varDate)
                                        strInsertbudget = strInsertbudget & "Insert into budgetitem_mst(account_code,cust_drgno,item_code,colour_code,category_code,commodity_code,model_code,usage_qty,"
                                        strInsertbudget = strInsertbudget & " ent_dt,ent_userid,upd_dt,upd_userid,Variant_Code,UNIT_CODE,DefaultModel,EndDate) Select distinct account_code,cust_drgno,item_code,colour_code,category_code,"
                                        strInsertbudget = strInsertbudget & " commodity_code,model_code,usage_qty,ent_dt,ent_userid,ent_dt,upd_userid,Variant_Code,UNIT_CODE,DefaultModel,'" & varDate & "' from tmp_budgetitem_mst where ip_address='" & gstrIpaddressWinSck & "' and account_code='" & txtCustCode.Text.Trim & "'And item_code='" & varItemcode & "' And cust_drgno='" & varCustdrgno & "'  AND UNIT_CODE='" & gstrUNITID & "'"

                                    Next
                                Else
                                    Exit Sub
                                End If
                            End If
                            ResetDatabaseConnection()
                            mP_Connection.BeginTrans()
                            mP_Connection.Execute("set dateformat dmy", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            mP_Connection.Execute(strInsert, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            If blnAllowBudget = True Then
                                mP_Connection.Execute(strInsertbudget, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                            mP_Connection.CommitTrans()
                            MsgBox("Transaction Completed Successfully.", MsgBoxStyle.Information, "eMPro")
                            CmdGrpCustomerITem.Revert()
                            CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                            CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                            CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                            txtCustCode.Text = ""
                            txtItemCode.Text = ""
                            txtItemDesc.Text = ""
                            spdITemDetails.Enabled = False
                            CmdGrpCustomerITem.Focus()
                        End If
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        Dim srtDELOldData As String = String.Empty

                        If blnAllowBudget = True Then
                            If validbudgetData() = False Then
                                Exit Sub
                            End If
                        End If
                        If Not ValidateBeforeSave() = False Then
                            '10737738
                            If ValidateVEHBOM() = False Then
                                Exit Sub
                            End If
                            strInsert = "" : strupdate = "" : strInsertbudget = "" : strDeletebudget = ""

                            'ISSUE ID       : 10222401 
                            mP_Connection.Execute("DELETE FROM TMPCustItem_Mst WHERE UNIT_CODE='" & gstrUNITID & "' AND IP_ADDRESS='" & gstrIpaddressWinSck & "'", ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                            intmaxrows = spdITemDetails.MaxRows
                            For intLoopCounter = 1 To intmaxrows
                                varMode = Nothing
                                varItemcode = Nothing
                                varColour = Nothing
                                varcategory = Nothing
                                varcommodity = Nothing
                                varCustdrgno = Nothing
                                varDate = Nothing

                                Call spdITemDetails.GetText(2, intLoopCounter, varItemcode)
                                Call spdITemDetails.GetText(5, intLoopCounter, varCustdrgno)
                                Call spdITemDetails.GetText(7, intLoopCounter, varMode)
                                Call spdITemDetails.GetText(12, intLoopCounter, varColour)
                                varcategory = Nothing
                                Call spdITemDetails.GetText(14, intLoopCounter, varcategory)
                                varcommodity = Nothing
                                Call spdITemDetails.GetText(16, intLoopCounter, varcommodity)

                                If CheckForItemMainGroup(varItemcode) = True Then
                                    If blnAllowBudget = True Then
                                        If Len(varColour) > 0 Then
                                            If Len(varcategory) <= 0 Or Len(varcommodity) <= 0 Then
                                                MsgBox("Enter Category and Commodity Code.", MsgBoxStyle.Information, ResolveResString(100))
                                                Exit Sub
                                            End If
                                            rsbudgetmodel = New ClsResultSetDB

                                            rsbudgetmodel.GetResult("select ACCOUNT_CODE,CUST_DRGNO,ITEM_CODE,COLOUR_CODE,CATEGORY_CODE,COMMODITY_CODE,MODEL_CODE,USAGE_QTY,VARIANT_CODE from budgetitem_mst where account_code ='" & txtCustCode.Text.Trim & "' and item_code='" & varItemcode & "' and cust_drgno='" & varCustdrgno & "' AND UNIT_CODE='" & gstrUNITID & "' union select ACCOUNT_CODE,CUST_DRGNO,ITEM_CODE,COLOUR_CODE,CATEGORY_CODE,COMMODITY_CODE,MODEL_CODE,USAGE_QTY,VARIANT_CODE from tmp_budgetitem_mst where ip_address= '" & gstrIpaddressWinSck & "'and  account_code ='" & txtCustCode.Text.Trim & "' and item_code='" & varItemcode & "' and cust_drgno='" & varCustdrgno & "' AND UNIT_CODE='" & gstrUNITID & "'")
                                            If rsbudgetmodel.GetNoRows <= 0 Then
                                                MsgBox("Enter Model related Information.", MsgBoxStyle.Information, ResolveResString(100))
                                                Exit Sub
                                            End If
                                            rsbudgetmodel.ResultSetClose()
                                            rsbudgetmodel = Nothing
                                        End If
                                    End If
                                    If blnAllowBudget = True And Len(varColour) > 0 Then
                                        UpdateBudgetData(intLoopCounter)
                                        varDate = Nothing
                                        Call spdITemDetails.GetText(24, intLoopCounter, varDate)
                                        'Call spdITemDetails.GetRowItemData(2)
                                        strInsertbudget = strInsertbudget & "Insert into budgetitem_mst(account_code,cust_drgno,item_code,colour_code,category_code,commodity_code,model_code,usage_qty,"
                                        strInsertbudget = strInsertbudget & " ent_dt,ent_userid,upd_dt,upd_userid,Variant_Code,UNIT_CODE,DefaultModel,EndDate) select distinct account_code,cust_drgno,item_code,colour_code,category_code,"
                                        strInsertbudget = strInsertbudget & " commodity_code,model_code,usage_qty,ent_dt,ent_userid,ent_dt,upd_userid,Variant_Code,UNIT_CODE,DefaultModel,'" & varDate & "' from tmp_budgetitem_mst where ip_address='" & gstrIpaddressWinSck & "' and account_code='" & txtCustCode.Text.Trim & "' And item_code='" & varItemcode & "' And cust_drgno='" & varCustdrgno & "' AND UNIT_CODE='" & gstrUNITID & "'"
                                    End If
                                End If
                                UpdateData(intLoopCounter)
                                If Trim(varMode) = "E" Then
                                    'eMpro-20090423-30547
                                    'commented by shubhra
                                    'UpdateData(intLoopCounter)
                                ElseIf Trim(varMode) = "A" Then
                                    If ValidRowData(intLoopCounter) = True Then
                                        Call InsertData(intLoopCounter)
                                    Else
                                        Exit Sub
                                    End If
                                End If
                            Next
                            'If Len(Trim(strupdate)) > 0 Then
                            ResetDatabaseConnection()
                            mP_Connection.BeginTrans()
                            'mP_Connection.Execute(strupdate, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                            'ISSUE ID       : 10222401 
                            strupdateCustItemMst = "UPDATE X SET X.Drg_Desc = Y.Drg_Desc,X.VarModel=Y.VarModel,X.Active=Y.Active,X.schupldreqd=Y.schupldreqd,X.commodity=Y.commodity,X.container=Y.container,X.packing_code=Y.packing_code,X.Upd_dt='" & getDateForDB(GetServerDate()) & "',X.Upd_Userid='" & mP_User & "',X.BinQuantity=Y.BinQuantity,X.Item_Desc = Y.Item_Desc ,X.SHOP_CODE=Y.SHOP_CODE ,X.GATE_NO=Y.GATE_NO,X.DECL_NO=Y.DECL_NO FROM CustItem_Mst X, TMPCustItem_Mst Y WHERE X.UNIT_CODE=Y.UNIT_CODE AND X.Account_Code=Y.Account_Code AND X.Cust_Drgno=Y.Cust_Drgno AND X.Item_code=Y.Item_code AND Y.IP_ADDRESS='" & gstrIpaddressWinSck & "' AND Y.UNIT_CODE='" & gstrUNITID & "'"
                            mP_Connection.Execute(strupdateCustItemMst, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)


                            If Len(Trim(strDeletebudget)) > 0 Then
                                mP_Connection.Execute(strDeletebudget, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                            If Len(Trim(strInsertbudget)) > 0 Then
                                mP_Connection.Execute(strInsertbudget, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                            If Len(Trim(strInsert)) > 0 Then
                                mP_Connection.Execute(strInsert, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                            mP_Connection.CommitTrans()
                            MsgBox(" Transaction Completed Successfully.", MsgBoxStyle.Information, "eMPro")
                            mP_Connection.Execute("delete from tmp_budgetitem_mst where ip_address='" & gstrIpaddressWinSck & "' AND UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            Me.spdITemDetails.MaxRows = 0
                            CmdGrpCustomerITem.Revert()
                            CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                            CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                            CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                            txtCustCode.Text = ""
                            txtItemCode.Text = ""
                            txtItemDesc.Text = ""
                            txtCustCode.Enabled = True : txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            spdITemDetails.Enabled = False
                            'End If
                            CmdGrpCustomerITem.Focus()
                        End If
                End Select

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                With spdITemDetails
                    .Enabled = True
                    .Col = 1
                    .Col2 = 1
                    .BlockMode = True
                    .ColHidden = True
                    .BlockMode = False
                    For intcnt = 1 To .MaxRows
                        .Row = intcnt
                        .Col = 1
                        If .ForeColor = System.Drawing.Color.Red Then
                            MsgBox("Customer Part Code and description are not editable for Items in Red Color.", MsgBoxStyle.Information, ResolveResString(100))
                            Exit For
                        End If
                    Next
                    intmaxrows = spdITemDetails.MaxRows
                    For intLoopCounter = 1 To intmaxrows
                        .Row = intLoopCounter
                        .Row2 = intLoopCounter
                        If .ForeColor = System.Drawing.Color.Black Then
                            Call .SetText(7, intLoopCounter, "E")
                            .Col = 6
                            .Col2 = 9
                            .Row = intLoopCounter
                            .Row2 = intLoopCounter
                            .BlockMode = True
                            .Lock = False
                            .BlockMode = False
                            .Col = 10
                            .Col2 = 14
                            .Row = intLoopCounter
                            .Row2 = intLoopCounter
                            .BlockMode = True
                            .Lock = False
                            .BlockMode = False

                        Else
                            Call .SetText(7, intLoopCounter, "")
                            .Col = 7
                            .Col2 = 20
                            .Row = intLoopCounter
                            .Row2 = intLoopCounter
                            .BlockMode = True
                            .Lock = False
                            .BlockMode = False
                        End If
                    Next
                    .Col = 2 : .Lock = True
                End With
                txtCustCode.Enabled = False : txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                Call frmMKTMST0003_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
                Call CheckDeleteMark()
                If blnCheckDelItems = True Then
                    With spdITemDetails
                        .Enabled = True
                        intmaxrows = spdITemDetails.MaxRows
                        strDelete = ""
                        For intLoopCounter = 1 To intmaxrows
                            varDeleteFlag = Nothing
                            Call .GetText(1, intLoopCounter, varDeleteFlag)
                            If varDeleteFlag = "1" Then
                                Call DeleteRecord(intLoopCounter)
                            End If
                        Next
                        If ConfirmWindow(10054, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_CRITICAL, 60096) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                            mP_Connection.Execute(strDelete, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            If spdITemDetails.MaxRows > 1 Then
                                spdITemDetails.MaxRows = 0
                                Call DisplayDetailsinGrid()
                                CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                                CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                            Else
                                txtCustCode.Text = ""
                                CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                                CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                                spdITemDetails.Enabled = False
                            End If
                        End If
                    End With
                Else
                    MsgBox("No Item is Selected For Deletion", MsgBoxStyle.Information, "eMPro")
                    CmdGrpCustomerITem.Focus()
                End If
        End Select
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub spdITemDetails_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles spdITemDetails.ButtonClicked
        On Error GoTo ErrHandler
        Dim varItemCode As Object
        Dim varHelpItem As Object
        Dim stritemdesc As String
        Dim StrItemCode As String
        Dim strConsMeasureCode As String
        Dim strItems As String
        Dim intLoopCounter As Integer
        Dim intMaxCounter As Integer
        Dim rsDesc As ClsResultSetDB
        Dim rsCustItemMst As ClsResultSetDB
        Dim rscategory As ClsResultSetDB
        Dim varHelpCategory As Object
        Dim varHelpColour As Object
        Dim strItemList As String
        Dim strStoredItemList As String
        Dim intItems As Integer
        Dim strcolour As String
        Dim StrItemCodes As String
        Dim StrCustdrgno As String
        Dim strColourcode As String
        Dim strCategory As String
        Dim strCommodity As String
        Dim objVal As Object = Nothing

        Dim rsAllowMultiParts As ClsResultSetDB
        rsAllowMultiParts = New ClsResultSetDB
        rsAllowMultiParts.GetResult("SELECT AllowMultiplePartCodes from Sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsAllowMultiParts.RowCount > 0 Then
            blnAllowMultiPartNo = If(rsAllowMultiParts.GetValue("AllowMultiplePartCodes") = True, True, False)
        End If
        rsAllowMultiParts = Nothing
        If e.col = 3 Then
            If blnAllowMultiPartNo = False Then
                strItemList = ""
                With spdITemDetails
                    If spdITemDetails.MaxRows > 0 Then
                        For intItems = 1 To spdITemDetails.MaxRows
                            varItemCode = Nothing
                            spdITemDetails.GetText(2, intItems, varItemCode)
                            If Len(Trim(varItemCode)) > 0 Then
                                If Len(Trim(strItemList)) > 0 Then strItemList = strItemList & ","
                                strItemList = strItemList & "'" & varItemCode & "'"
                            End If
                        Next
                    End If
                End With
                'Check for Item Already Exist in CustITemMSt For That Customer
                rsCustItemMst = New ClsResultSetDB
                rsCustItemMst.GetResult("Select Item_code from CustItem_Mst where Account_code = '" & Trim(txtCustCode.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rsCustItemMst.GetNoRows > 0 Then
                    intMaxCounter = rsCustItemMst.GetNoRows
                    rsCustItemMst.MoveFirst()
                    While Not rsCustItemMst.EOFRecord
                        varItemCode = rsCustItemMst.GetValue("Item_code")
                        If Len(Trim(varItemCode)) > 0 Then
                            If Len(Trim(strItems)) > 0 Then strItems = Trim(strItems) & ","
                            strItems = strItems & "'" & Trim(varItemCode) & "'"
                        End If
                        rsCustItemMst.MoveNext()
                    End While
                End If
                rsCustItemMst.ResultSetClose()
                varItemCode = Nothing
                Call spdITemDetails.GetText(2, spdITemDetails.ActiveRow, varItemCode)
                If Len(strItems) > 0 And Len(strItemList) > 0 Then 'Filter for Grid Item codes as well in CustItem_mst
                    If Len(Trim(varItemCode)) > 0 Then 'For the already entered value for filter
                        '10869290
                        'varHelpItem = ShowList(1, 6, CStr(varItemCode), "Item_code", "Description", "Item_mst", "and Item_Main_grp not in ('M') and STATUS = 'A' and Hold_Flag = 0 and Item_Code NOT IN (" & strItems & "," & strItemList & ")")
                        varHelpItem = ShowList(1, 6, CStr(varItemCode), "Item_code", "Description", "Item_mst", "and STATUS = 'A' and Hold_Flag = 0 and Item_Code NOT IN (" & strItems & "," & strItemList & ")")
                        '10869290
                    Else 'New item entry(itemcode text is blank)
                        '10869290
                        'varHelpItem = ShowList(1, 6, "", "Item_code", "Description", "Item_mst", "and Item_Main_grp not in ('M') and STATUS = 'A' and Hold_Flag = 0 and Item_Code NOT IN (" & strItems & "," & strItemList & ")")
                        varHelpItem = ShowList(1, 6, "", "Item_code", "Description", "Item_mst", "and STATUS = 'A' and Hold_Flag = 0 and Item_Code NOT IN (" & strItems & "," & strItemList & ")")
                        '10869290
                    End If
                ElseIf Len(strItems) > 0 Then 'Filter for CustItem_mst Items
                    If Len(Trim(varItemCode)) > 0 Then 'For the already entered value for filter
                        '10869290
                        'varHelpItem = ShowList(1, 6, CStr(varItemCode), "Item_code", "Description", "Item_mst", "and Item_Main_grp not in ('M') and STATUS = 'A' and Hold_Flag = 0  And Item_Code NOT IN (" & strItems & ")")
                        varHelpItem = ShowList(1, 6, CStr(varItemCode), "Item_code", "Description", "Item_mst", " and STATUS = 'A' and Hold_Flag = 0  And Item_Code NOT IN (" & strItems & ")")
                    Else 'New item entry(itemcode text is blank)
                        'varHelpItem = ShowList(1, 6, "", "Item_code", "Description", "Item_mst", "and Item_Main_grp not in ('M') and STATUS = 'A' and Hold_Flag = 0 and Item_Code NOT IN (" & strItems & ")")
                        varHelpItem = ShowList(1, 6, "", "Item_code", "Description", "Item_mst", "and STATUS = 'A' and Hold_Flag = 0 and Item_Code NOT IN (" & strItems & ")")
                        '10869290
                    End If
                    '10869290
                ElseIf Len(strItemList) > 0 Then 'Filter for Items in grid
                    If Len(Trim(varItemCode)) > 0 Then 'For the already entered value for filter
                        'varHelpItem = ShowList(1, 6, CStr(varItemCode), "Item_code", "Description", "Item_mst", "and Item_Main_grp not in ('M') and STATUS = 'A' and Hold_Flag = 0  Item_Code NOT IN (" & strItemList & ")")
                        varHelpItem = ShowList(1, 6, CStr(varItemCode), "Item_code", "Description", "Item_mst", "and STATUS = 'A' and Hold_Flag = 0  Item_Code NOT IN (" & strItemList & ")")
                    Else 'New item entry(itemcode text is blank)
                        'varHelpItem = ShowList(1, 6, "", "Item_code", "Description", "Item_mst", "and Item_Main_grp not in ('M') and STATUS = 'A' and Hold_Flag = 0 and Item_Code NOT IN (" & strItemList & ")")
                        varHelpItem = ShowList(1, 6, "", "Item_code", "Description", "Item_mst", "and STATUS = 'A' and Hold_Flag = 0 and Item_Code NOT IN (" & strItemList & ")")
                    End If
                Else ' No Filter criteria
                    If Len(Trim(varItemCode)) > 0 Then 'For the already entered value for filter
                        'varHelpItem = ShowList(1, 6, CStr(varItemCode), "Item_code", "Description", "Item_mst", "and Item_Main_grp not in ('M') and STATUS = 'A' and Hold_Flag = 0 ")
                        varHelpItem = ShowList(1, 6, CStr(varItemCode), "Item_code", "Description", "Item_mst", "and STATUS = 'A' and Hold_Flag = 0 ")
                    Else 'New item entry(itemcode text is blank)
                        'varHelpItem = ShowList(1, 6, "", "Item_code", "Description", "Item_mst", "and Item_Main_grp not in ('M') and STATUS = 'A' and Hold_Flag = 0 ")
                        varHelpItem = ShowList(1, 6, "", "Item_code", "Description", "Item_mst", "and STATUS = 'A' and Hold_Flag = 0 ")
                    End If
                End If
                '10869290
            Else 'Show all Item codes
                If Len(Trim(varItemCode)) > 0 Then
                    'varHelpItem = ShowList(1, 6, CStr(varItemCode), "Item_code", "Description", "Item_mst", "and Item_Main_grp not in ('M') and STATUS = 'A' and Hold_Flag = 0 ")
                    varHelpItem = ShowList(1, 6, CStr(varItemCode), "Item_code", "Description", "Item_mst", "and STATUS = 'A' and Hold_Flag = 0 ")
                Else
                    'varHelpItem = ShowList(1, 6, "", "Item_code", "Description", "Item_mst", "and Item_Main_grp not in ('M') and STATUS = 'A' and Hold_Flag = 0")
                    varHelpItem = ShowList(1, 6, "", "Item_code", "Description", "Item_mst", " and STATUS = 'A' and Hold_Flag = 0")
                End If
            End If
            '10869290
            If varHelpItem = "-1" Then
                Call ConfirmWindow(10013, ConfirmWindowButtonsEnum.BUTTON_OK, ConfirmWindowImagesEnum.IMG_INFO)
            ElseIf varHelpItem = "" Or varHelpItem = String.Empty Then
                Call spdITemDetails.SetText(4, spdITemDetails.ActiveRow, "")
                Call spdITemDetails.SetText(6, spdITemDetails.ActiveRow, "")
                Exit Sub
            Else 'New Item
                Call spdITemDetails.SetText(2, spdITemDetails.ActiveRow, varHelpItem)
                rsDesc = New ClsResultSetDB
                '10869290
                'rsDesc.GetResult("Select Item_code,Description,Cons_Measure_Code from Item_Mst where Item_Main_grp not in ('M') and sTATUS = 'A' and Hold_Flag = 0 " & " and Item_Code ='" & varHelpItem & "' AND UNIT_CODE='" & gstrUNITID & "'")
                rsDesc.GetResult("Select Item_code,Description,Cons_Measure_Code from Item_Mst where  sTATUS = 'A' and Hold_Flag = 0 " & " and Item_Code ='" & varHelpItem & "' AND UNIT_CODE='" & gstrUNITID & "'")
                '10869290
                stritemdesc = rsDesc.GetValue("Description")
                Call spdITemDetails.SetText(4, spdITemDetails.ActiveRow, stritemdesc)
                Call spdITemDetails.SetText(6, spdITemDetails.ActiveRow, stritemdesc)
                StrItemCode = rsDesc.GetValue("Item_code")
                strConsMeasureCode = rsDesc.GetValue("Cons_Measure_Code")
                Call spdITemDetails.SetText(2, spdITemDetails.ActiveRow, StrItemCode)
                rsDesc.ResultSetClose()
                If CheckForItemMainGroup(StrItemCode) = False Then
                    With Me.spdITemDetails
                        .Row = .ActiveRow : .Row2 = .ActiveRow : .Col = 12 : .Col2 = 13 : .BlockMode = True : .Lock = True : .BlockMode = False
                        .Row = .ActiveRow : .Row2 = .ActiveRow : .Col = 14 : .Col2 = 15 : .BlockMode = True : .Lock = True : .BlockMode = False
                        .Row = .ActiveRow : .Row2 = .ActiveRow : .Col = 16 : .Col2 = 17 : .BlockMode = True : .Lock = True : .BlockMode = False
                        .Row = .ActiveRow : .Row2 = .ActiveRow : .Col = 20 : .BlockMode = True : .Lock = True : .BlockMode = False
                    End With
                End If
            End If
        End If
        If e.col = 9 Then
            varHelpColour = ShowList(1, 6, "", "Model_code", "Model_desc", "Budget_Model_mst", "")
            If varHelpColour = "-1" Then
                Call ConfirmWindow(10013, ConfirmWindowButtonsEnum.BUTTON_OK, ConfirmWindowImagesEnum.IMG_INFO)
            ElseIf varHelpColour = "" Or varHelpColour = String.Empty Then
                Call spdITemDetails.SetText(9, spdITemDetails.ActiveRow, "")
                Exit Sub
            Else
                Call spdITemDetails.SetText(9, spdITemDetails.ActiveRow, varHelpColour)
            End If
            'Dim strHelp As String()
            'Dim strSQL = "SELECT DISTINCT A.VARMODEL,A.ACCOUNT_CODE,B.CUST_NAME FROM CUSTITEM_MST A,CUSTOMER_MST B WHERE A.unit_code=B.unit_Code and A.unit_Code='" & gstrUNITID & "' and  A.ACCOUNT_CODE = B.CUSTOMER_CODE AND A.VARMODEL IS NOT NULL AND A.VARMODEL <> '' ORDER BY A.VARMODEL,A.ACCOUNT_CODE"
            'strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select Model", 0, 0, "")
            'If Not IsNothing(strHelp) AndAlso IsNothing(strHelp(1)) Then
            '    MessageBox.Show("Warehouse not defined for selected vendor.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            '    Return
            'End If
            'If Not IsNothing(strHelp) Then
            '    Dim strModel = strHelp(0).Trim
            'End If
        End If
        If e.col = 13 Then
            varHelpColour = ShowList(1, 6, "", "colour_code", "colour_desc", "colour_mst", "and active=1 ")
            If varHelpColour = "-1" Then
                Call ConfirmWindow(10013, ConfirmWindowButtonsEnum.BUTTON_OK, ConfirmWindowImagesEnum.IMG_INFO)
            ElseIf varHelpColour = "" Or varHelpColour = String.Empty Then
                Call spdITemDetails.SetText(12, spdITemDetails.ActiveRow, "")
                Call spdITemDetails.SetText(14, spdITemDetails.ActiveRow, "")
                Call spdITemDetails.SetText(16, spdITemDetails.ActiveRow, "")
                Exit Sub
            Else
                Call spdITemDetails.SetText(12, spdITemDetails.ActiveRow, varHelpColour)
                Call spdITemDetails.SetText(14, spdITemDetails.ActiveRow, "")
                Call spdITemDetails.SetText(16, spdITemDetails.ActiveRow, "")
            End If
        End If
        If e.col = 15 Then
            With spdITemDetails
                .Col = 12
                .Row = .ActiveRow
                strcolour = Trim(.Value)
            End With
            varHelpCategory = ShowList(1, 6, "", "Category", "colour_desc", "colour_mst", " and colour_code ='" & strcolour & "' and active=1 ")
            If varHelpCategory = "-1" Then
                Call ConfirmWindow(10013, ConfirmWindowButtonsEnum.BUTTON_OK, ConfirmWindowImagesEnum.IMG_INFO)
            ElseIf varHelpCategory = "" Or varHelpCategory = String.Empty Then
                Call spdITemDetails.SetText(14, spdITemDetails.ActiveRow, "")
                Exit Sub
            Else
                Call spdITemDetails.SetText(14, spdITemDetails.ActiveRow, varHelpCategory)
            End If
        End If
        If e.col = 17 Then
            varHelpColour = ShowList(1, 6, "", "commodity_code", "commodity_desc", "commodity_mst", "and active=1 ")
            If varHelpColour = "-1" Then
                Call ConfirmWindow(10013, ConfirmWindowButtonsEnum.BUTTON_OK, ConfirmWindowImagesEnum.IMG_INFO)
            ElseIf varHelpColour = "" Or varHelpColour = String.Empty Then
                Call spdITemDetails.SetText(16, spdITemDetails.ActiveRow, "")
                Exit Sub
            Else
                Call spdITemDetails.SetText(16, spdITemDetails.ActiveRow, varHelpColour)
            End If
        End If
        If e.col = 20 Then
            Dim frmForm_modeldetails As New FrmModelDetails
            frmForm_modeldetails.Customercode = txtCustCode.Text.Trim
            With spdITemDetails
                .Col = 2
                .Row = .ActiveRow
                StrItemCodes = Trim(.Value)
                .Col = 5
                StrCustdrgno = Trim(.Value)
                .Col = 12
                strColourcode = Trim(.Value)
                .Col = 14
                strCategory = Trim(.Value)
                .Col = 16
                strCommodity = Trim(.Value)
            End With
            If StrCustdrgno = "" Or strColourcode = "" Or strCategory = "" Or strCommodity = "" Then
                MsgBox("Enter the all fields (Cust. Part No/Colour/Category/Commodity) ", vbInformation + vbOKOnly, "eMPro")
                Exit Sub
            End If
            frmForm_modeldetails.Itemcode = StrItemCodes
            frmForm_modeldetails.Custdrgno = StrCustdrgno
            frmForm_modeldetails.Colourcode = strColourcode
            frmForm_modeldetails.Categorycode = strCategory
            frmForm_modeldetails.Commoditycode = strCommodity
            frmForm_modeldetails.Mode = CmdGrpCustomerITem.Mode
            frmForm_modeldetails.ShowDialog()
        End If

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub spdITemDetails_KeyDownEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles spdITemDetails.KeyDownEvent
        On Error GoTo ErrHandler
        Select Case e.keyCode
            Case Keys.F1
                If spdITemDetails.ActiveCol = 2 Then
                    If CmdGrpCustomerITem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT And spdITemDetails.ActiveRow < spdITemDetails.MaxRows Then Exit Sub
                    Call spdITemDetails_ButtonClicked(sender, New AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent(3, spdITemDetails.ActiveRow, 0))
                End If
                If spdITemDetails.ActiveCol = 9 Then
                    If CmdGrpCustomerITem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT And spdITemDetails.ActiveRow < spdITemDetails.MaxRows Then Exit Sub
                    Call spdITemDetails_ButtonClicked(sender, New AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent(9, spdITemDetails.ActiveRow, 0))
                End If
                If spdITemDetails.ActiveCol = 12 Then
                    If CmdGrpCustomerITem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT And spdITemDetails.ActiveRow < spdITemDetails.MaxRows Then Exit Sub
                    Call spdITemDetails_ButtonClicked(sender, New AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent(13, spdITemDetails.ActiveRow, 0))
                End If
                If spdITemDetails.ActiveCol = 14 Then
                    If CmdGrpCustomerITem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT And spdITemDetails.ActiveRow < spdITemDetails.MaxRows Then Exit Sub
                    Call spdITemDetails_ButtonClicked(sender, New AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent(15, spdITemDetails.ActiveRow, 0))
                End If
                If spdITemDetails.ActiveCol = 16 Then
                    If CmdGrpCustomerITem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT And spdITemDetails.ActiveRow < spdITemDetails.MaxRows Then Exit Sub
                    Call spdITemDetails_ButtonClicked(sender, New AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent(17, spdITemDetails.ActiveRow, 0))
                End If
            Case 13
                If spdITemDetails.ActiveCol = 14 Then
                    If CmdGrpCustomerITem.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                        '***If this is Last Row & Last Column
                        If spdITemDetails.ActiveRow = spdITemDetails.MaxRows Then
                            ' Call addNewInSpread()
                            AddNewRowType()
                        End If
                        '***
                    End If
                End If
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub spdITemDetails_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles spdITemDetails.KeyPressEvent
        On Error GoTo ErrHandler
        Select Case e.keyAscii
            Case Keys.Return
            Case 39, 34, 96
                e.keyAscii = 0
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub spdITemDetails_KeyUpEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyUpEvent) Handles spdITemDetails.KeyUpEvent
        On Error GoTo ErrHandler
        If ((e.shift = 2) And (e.keyCode = Keys.N)) Then
            If CmdGrpCustomerITem.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                'Call addNewInSpread()
                AddNewRowType()
                Call spdITemDetails_LeaveCell(sender, New AxFPSpreadADO._DSpreadEvents_LeaveCellEvent(spdITemDetails.ActiveCol, spdITemDetails.MaxRows - 1, 3, spdITemDetails.MaxRows, False))
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub spdITemDetails_LeaveCell(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles spdITemDetails.LeaveCell
        On Error GoTo ErrHandler
        Dim varFeild As Object
        Dim varItemCode As Object
        Dim varCurrency As Object
        Dim varCurRowOne As Object
        Dim intMaxCounter As Integer
        Dim intLoopCounter As Integer
        Dim strDescription As String
        Dim rsAllowMultiParts As ClsResultSetDB
        If e.newCol = -1 Then
            Exit Sub
        End If
        If CmdGrpCustomerITem.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            With spdITemDetails
                .Row = e.row : .Row2 = e.row : .Col = e.col : .Col2 = e.col : .BlockMode = True
                If .Lock = False Then
                    '****Valid ItemCode
                    If e.col = 2 Then
                        varFeild = Nothing
                        Call .GetText(e.col, e.row, varFeild)
                        If Len(Trim(CStr(varFeild))) > 0 Then
                            If Len(Trim(varFeild)) > 0 Then
                                '10869290
                                'If IsExistFieldValue(CStr(Trim(varFeild)), "Item_code", "Item_Mst", "Item_Main_Grp Not in ('M') and Status = 'A' and Hold_Flag =0 AND UNIT_CODE='" & gstrUNITID & "'") = False Then
                                If IsExistFieldValue(CStr(Trim(varFeild)), "Item_code", "Item_Mst", "Status = 'A' and Hold_Flag =0 AND UNIT_CODE='" & gstrUNITID & "'") = False Then
                                    '10869290
                                    'MsgBox("Invalid Item Code Press F1 For Help", vbInformation, "eMPro")
                                    MsgBox("Selected Item Code is either On Hold or Inactive.", vbInformation, "eMPro")
                                    .Col = e.col : .Row = .ActiveRow : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                Else
                                    rsAllowMultiParts = New ClsResultSetDB
                                    rsAllowMultiParts.GetResult("SELECT AllowMultiplePartCodes from Sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                                    If rsAllowMultiParts.RowCount > 0 Then
                                        blnAllowMultiPartNo = IIf(rsAllowMultiParts.GetValue("AllowMultiplePartCodes") = True, True, False)
                                    End If
                                    rsAllowMultiParts.ResultSetClose()
                                    If blnAllowMultiPartNo = False Then
                                        If AlreadyExist(Trim(txtCustCode.Text), Trim(varFeild), e.row) = True Then
                                            MsgBox("Internal Part No for this Customer " & vbCrLf & " Either already defined or already selected in Grid", vbInformation + vbOKOnly, "eMPro")
                                            Call .SetText(2, e.row, "")
                                            Call .SetText(4, e.row, "")
                                            Call .SetText(6, e.row, "")
                                            Exit Sub
                                        End If
                                    End If
                                    '10869290
                                    'strDescription = ReturnDescription("Description", "ITem_Mst", "Item_Main_Grp not in ('M') and Status = 'A' and Hold_Flag =0 and Item_Code ='" & CStr(varFeild) & "' AND UNIT_CODE='" & gstrUNITID & "'")
                                    strDescription = ReturnDescription("Description", "ITem_Mst", "Status = 'A' and Hold_Flag =0 and Item_Code ='" & CStr(varFeild) & "' AND UNIT_CODE='" & gstrUNITID & "'")
                                    '10869290
                                    Call .SetText(4, e.row, strDescription)
                                    Call .SetText(6, e.row, strDescription)
                                End If
                            End If
                        Else
                            Call .SetText(6, e.row, "")
                        End If
                    End If
                    '****
                    If CmdGrpCustomerITem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                        If e.col = 5 Then
                            varFeild = Nothing
                            Call .GetText(e.col, e.row, varFeild)
                            varItemCode = Nothing
                            Call .GetText(2, e.row, varItemCode)
                            If IsExistFieldValue(CStr(Trim(varFeild)), "Cust_DrgNo", "CustItem_Mst", "ITem_Code = '" & Trim(varItemCode) & "' and Account_Code = '" & Trim(txtCustCode.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'") = True Then
                                MsgBox("This Drawing No Already Exist for this Item", vbInformation, "eMPro")
                                .Col = e.col : .Row = .ActiveRow : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            End If
                        End If
                    End If
                    If e.col = 12 Then
                        varFeild = Nothing
                        Call .GetText(e.col, .ActiveRow, varFeild)
                        If Len(Trim(CStr(varFeild))) > 0 Then
                            If Len(Trim(varFeild)) > 0 Then
                                If IsExistFieldValue(CStr(Trim(varFeild)), "colour_code", "colour_Mst", "UNIT_CODE='" & gstrUNITID & "'") = False Then
                                    MsgBox("Invalid Colour  Press F1 For Help", vbInformation, "eMPro")
                                    Call .SetText(e.col, .ActiveRow, "")
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                End If
                            End If
                        Else
                            Call .SetText(14, .ActiveRow, "")
                            Call .SetText(16, .ActiveRow, "")
                        End If
                    End If
                    If e.col = 14 Then
                        varFeild = Nothing
                        Call .GetText(e.col, .ActiveRow, varFeild)
                        If Len(Trim(CStr(varFeild))) > 0 Then
                            If Len(Trim(varFeild)) > 0 Then
                                If IsExistFieldValue(CStr(Trim(varFeild)), "category", "colour_Mst", "UNIT_CODE='" & gstrUNITID & "'") = False Then
                                    MsgBox("Invalid Category  Code Press F1 For Help", vbInformation, "eMPro")
                                    Call .SetText(e.col, .ActiveRow, "")
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                End If
                            End If
                        End If
                    End If
                    If e.col = 16 Then 'commodity
                        varFeild = Nothing
                        Call .GetText(e.col, .ActiveRow, varFeild)
                        If Len(Trim(CStr(varFeild))) > 0 Then
                            If Len(Trim(varFeild)) > 0 Then
                                If IsExistFieldValue(CStr(Trim(varFeild)), "commodity_code", "commodity_Mst", "UNIT_CODE='" & gstrUNITID & "'") = False Then
                                    MsgBox("Invalid Commodity Code Press F1 For Help", vbInformation, "eMPro")
                                    Call .SetText(e.col, .ActiveRow, "")
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                End If
                            End If
                        End If
                    End If
                    '****Valid row check in case of last column
                    If Not e.newRow = -1 Then
                        If e.row <> e.newRow Then
                            varFeild = Nothing
                            Call .GetText(2, e.row, varFeild)
                            If Len(Trim(varFeild)) = 0 Then
                                If spdITemDetails.MaxRows > 1 Then
                                    .MaxRows = .MaxRows - 1
                                    .Col = 2 : .Row = .MaxRows : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                End If
                            Else
                                If ValidRowData(e.row) = False Then
                                    If spdITemDetails.MaxRows > 1 Then
                                        .MaxRows = .MaxRows - 1
                                        .Col = 2 : .Row = .MaxRows : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    End If
                                End If
                            End If
                        End If
                    End If
                    '****
                End If
                .BlockMode = False
            End With
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")  '("machine_master.htm")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub GetPackingLevelCode()
        'Revised By     : Manoj Kr. Vaish
        'Issue ID       : eMpro-20090611-32362
        'Revision Date  : 14 Jun 2009
        'History        : Get the Packing Level from the Lists table-Key='Packing Level'
        '****************************************************************************************
        Dim rspackinglevel As New ClsResultSetDB
        Dim strquery As String
        Dim strPackingCode As String
        strFinalPackingCode = ""
        strquery = "Select (Code +'-'+Key2)as Packing_Level from lists where UNIT_CODE='" & gstrUNITID & "' and key1='Packing Level'"
        rspackinglevel.GetResult(strquery)
        If rspackinglevel.GetNoRows > 0 Then
            rspackinglevel.MoveFirst()
            While Not rspackinglevel.EOFRecord
                strPackingCode = IIf((rspackinglevel.GetValue("Packing_Level") = "Unknown"), "", rspackinglevel.GetValue("Packing_Level"))
                strFinalPackingCode = strFinalPackingCode & strPackingCode & Chr(9) '& "[V]alue":
                rspackinglevel.MoveNext()
            End While
            strFinalPackingCode = VB.Left(strFinalPackingCode, Len(strFinalPackingCode) - 1)
            rspackinglevel.ResultSetClose()
        End If
        rspackinglevel = Nothing
    End Sub
    Private Function CheckForItemMainGroup(ByVal Item_code As String) As Boolean
        Dim rstHelpDb As ClsResultSetDB
        CheckForItemMainGroup = False
        Try
            rstHelpDb = New ClsResultSetDB
            Call rstHelpDb.GetResult("Select * from item_mst where item_main_grp in ('F','S') and item_code = '" & Item_code & "' and UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rstHelpDb.GetNoRows >= 1 Then
                CheckForItemMainGroup = True
            End If
            rstHelpDb.ResultSetClose()
            rstHelpDb = Nothing
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    Private Function ValidateVEHBOM() As Boolean

        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim intRow As Integer
        Dim strItem As String = String.Empty
        Try

            With spdITemDetails

                Using sqlCmd As SqlCommand = New SqlCommand("USP_VCHBOM_MODEL_DTL_VALIDATE")
                    sqlCmd.CommandType = CommandType.StoredProcedure
                    sqlCmd.Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    sqlCmd.Parameters.Add("@ITEM_CODE", SqlDbType.VarChar, 16)
                    sqlCmd.Parameters.Add("@CUST_CODE", SqlDbType.VarChar, 8).Value = Trim(txtCustCode.Text)
                    sqlCmd.Parameters.Add("@CUST_DRGNO", SqlDbType.VarChar, 30)
                    sqlCmd.Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 50).Value = gstrIpaddressWinSck
                    If (CmdGrpCustomerITem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD) Then
                        sqlCmd.Parameters.Add("@MODE", SqlDbType.Bit).Value = 1
                    Else
                        sqlCmd.Parameters.Add("@MODE", SqlDbType.Bit).Value = 0
                    End If
                    sqlCmd.Parameters.Add("@MSG", SqlDbType.VarChar, 5000).Direction = ParameterDirection.Output

                    For intRow = 1 To .MaxRows
                        varItemCode = Nothing
                        Call .GetText(2, intRow, varItemCode)
                        varDrgNo = Nothing
                        Call .GetText(5, intRow, varDrgNo)

                        sqlCmd.Parameters("@ITEM_CODE").Value = varItemCode
                        sqlCmd.Parameters("@CUST_DRGNO").Value = varDrgNo
                        SqlConnectionclass.ExecuteNonQuery(sqlCmd)

                        If sqlCmd.Parameters("@MSG").Value.ToString.Trim.Length > 0 Then
                            strItem += varItemCode + ","
                        End If
                        'Return True
                    Next

                End Using

            End With
            If strItem.ToString.Trim.Length > 0 Then
                MsgBox("Model Details is mandatory for the following items as these items are linked for the Customer for which Marketing Schedule (Vehicle BOM) is enabled." + vbCrLf + strItem, MsgBoxStyle.Exclamation, ResolveResString(100))
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Public Function DisplayDetailsinGridNew() As Object
        'ADDED AGAINST 1234MILIND
        ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
        Try
            Dim intLoopCounter As Short
            Dim intMaxCounter As Short
            Dim StrItemCode As String
            Dim NoOfRed As Short
            Dim intPackingCode As Short
            Dim strSQL As String = String.Empty
            Dim boolRow As Boolean = False
            Dim oDr As SqlDataReader = Nothing
            Dim oDS As DataSet
            Dim intRow As Integer
            strSQL = "Select TOP 1 1 from CustITem_Mst  Where UNIT_CODE='" & gstrUNITID & "' AND Account_code ='" & txtCustCode.Text.Trim & "' ORDER BY Cust_drgNo"
            boolRow = SqlConnectionclass.ExecuteScalar(strSQL)

            If boolRow = True Then
                Using sqlCmd As SqlCommand = New SqlCommand
                    With sqlCmd
                        .Connection = SqlConnectionclass.GetConnection
                        .CommandText = "USP_DISPLAY_CUSTITEM_MST_IN_GRID"
                        .CommandTimeout = 0
                        .CommandType = CommandType.StoredProcedure
                        .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                        .Parameters.Add("@CUSTOMER_CODE", SqlDbType.VarChar, 20).Value = txtCustCode.Text.Trim
                        .Parameters.Add("@MODE", SqlDbType.Char, 1).Value = "V"
                        .Parameters.Add("@ITEM_CODE", SqlDbType.VarChar, 16).Value = txtItemCode.Text.Trim
                        .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                        Using dt As DataTable = SqlConnectionclass.GetDataTable(sqlCmd)
                            If dt.Rows.Count > 0 Then
                                With spdITemDetails
                                    .MaxRows = 0
                                    For Each row As DataRow In dt.Rows
                                        'addNewInSpread()
                                        AddNewRowType()
                                        .Row = .MaxRows
                                        .Col = 2 : .Text = Convert.ToString(row("ITEM_CODE"))
                                        .Col = 2 : .CellTag = Convert.ToBoolean(row("CUSTORD_STATUS"))
                                        .Col = 4 : .Text = Convert.ToString(row("DESCRIPTION"))
                                        .Col = 5 : .Text = Convert.ToString(row("Cust_DrgNo"))
                                        .Col = 6 : .Text = Convert.ToString(row("Drg_desc"))
                                        .Col = 8 : .Text = Convert.ToString(row("BinQuantity"))
                                        .Col = 9 : .Text = Convert.ToString(row("VarModel"))

                                        If Convert.ToBoolean(row("AUTO_INVOICE_PART")) = True Then
                                            .Col = 21 : .Col2 = 21
                                            .Row = .ActiveRow : .Row2 = .ActiveRow
                                            .BlockMode = True : .Lock = True : .BlockMode = False
                                        End If

                                        If Convert.ToString(row("Active")) = False Then
                                            .Col = 10 : .Value = 0
                                        Else
                                            .Col = 10 : .Value = 1
                                        End If
                                        If Convert.ToString(row("schupldreqd")) = False Then
                                            .Col = 11 : .Value = 0
                                        Else
                                            .Col = 11 : .Value = 1
                                        End If
                                        .Col = 12 : .Text = Convert.ToString(row("Colour_code"))
                                        .Col = 14 : .Text = Convert.ToString(row("category_code"))
                                        .Col = 16 : .Text = Convert.ToString(row("commodity_code"))
                                        .Col = 18 : .Text = Convert.ToString(row("Container"))
                                        Call GetPackingLevelCode()
                                        spdITemDetails.TypeComboBoxClear(19, intLoopCounter)
                                        spdITemDetails.TypeComboBoxList = strFinalPackingCode
                                        .Col = 19 : .Text = Convert.ToString(row("Packing_level"))
                                        If Convert.ToBoolean(row("CUSTORD_STATUS")) = True Then
                                            .Col = 1 : .Col2 = .MaxCols
                                            .Row = .MaxRows : .Row2 = .MaxRows
                                            .BlockMode = True : .ForeColor = System.Drawing.Color.Red : .BlockMode = False
                                            .Col = 1 : .Col2 = 1
                                            .Row = .MaxRows : .Row2 = .MaxRows
                                            .BlockMode = True : .Lock = True : .BlockMode = False
                                        Else
                                            .Col = 1 : .Col2 = .MaxCols
                                            .Row = .MaxRows : .Row2 = .MaxRows
                                            .BlockMode = True : .ForeColor = System.Drawing.Color.Black : .BlockMode = False
                                            .Col = 1 : .Col2 = 1
                                            .Row = .MaxRows : .Row2 = .MaxRows
                                            .BlockMode = True : .Lock = False : .BlockMode = False
                                            .Col = 6 : .Col2 = .MaxCols - 1
                                            .Row = .MaxRows : .Row2 = .MaxRows
                                            .BlockMode = True : .Lock = False : .BlockMode = False
                                        End If
                                        Dim endDt As String = VB6.Format(Convert.ToString(row("EndDate")), "dd/mm/yyyy")

                                        .Col = 22 : .Text = Convert.ToString(row("Shop_code"))
                                        .Col = 23 : .Text = Convert.ToString(row("GATE_NO"))

                                        .Col = 24 : .Text = endDt
                                        .Col = 25 : .Text = Convert.ToString(row("decl_no"))
                                        '.Lock
                                    Next

                                End With
                            End If
                        End Using
                    End With
                End Using
                CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                CmdGrpCustomerITem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
            Else
                MsgBox("No Item Defined For this Customer", MsgBoxStyle.Information, "eMPro")
                txtCustCode.Text = "" : txtCustCode.Focus()
            End If
            If blnAllowBudget = True Then
                mP_Connection.Execute("delete from tmp_budgetitem_mst where ip_address='" & gstrIpaddressWinSck & "' AND UNIT_CODE='" & gstrUNITID & "' ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                mP_Connection.Execute("insert into tmp_budgetitem_mst(account_code,cust_drgno,item_code,colour_code,category_code,commodity_code,MODEL_CODE,USAGE_QTY,ENT_DT,ENT_USERID,UPD_DT,UPD_USERID,ip_address,VARIANT_CODE,UNIT_CODE,DefaultModel,EndDate) select account_code,cust_drgno,item_code,colour_code,category_code,commodity_code,MODEL_CODE,USAGE_QTY,ENT_DT,ENT_USERID,UPD_DT,UPD_USERID,'" & gstrIpaddressWinSck & "',VARIANT_CODE,UNIT_CODE,DefaultModel,EndDate from budgetitem_mst where account_code='" & txtCustCode.Text.Trim & "' AND UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            Exit Function

        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Function

    Private Sub cmdItemHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdItemHelp.Click
        ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
        Dim strQry As String
        Dim strHelp() As String
        Try
            If txtCustCode.Text = "" Then
                MessageBox.Show("Select Customer Code first..!!!")
                Exit Sub
            End If
            strQry = "SELECT ITEM_CODE, ITEM_DESC, Cust_Drgno, Drg_Desc FROM CUSTITEM_MST WHERE UNIT_CODE ='" & gstrUNITID & "' AND ACCOUNT_CODE='" & txtCustCode.Text & "'"
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
            If Not (UBound(strHelp) = -1) Then
                If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                    MsgBox("No Record Found !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                    Exit Sub
                Else
                    txtItemCode.Text = strHelp(0)
                    txtItemDesc.Text = strHelp(1)
                    spdITemDetails.MaxRows = 0
                    DisplayDetailsinGridNew()
                End If

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub btnExportExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportExcel.Click
        ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
        Try
            txtItemCode.Text = ""
            txtItemDesc.Text = ""
            spdITemDetails.MaxRows = 0
            ExportInExcel()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub
    Public Sub ExportInExcel()
        ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)

        Dim strSQL As String = String.Empty
        Dim i, j As Integer
        Dim strVar As String
        Dim xlApp As Microsoft.Office.Interop.Excel.Application
        Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim raXL As Microsoft.Office.Interop.Excel.Range
        Dim misValue As Object = System.Reflection.Missing.Value
        Try
            If txtCustCode.Text.Trim = "" Then
                MessageBox.Show("Select Customer first...!!!")
                Exit Sub
            End If
            strSQL = "SELECT TOP 1 1 FROM CUSTITEM_MST WHERE ACCOUNT_CODE='" & txtCustCode.Text & "' AND UNIT_CODE='" & gstrUNITID & "'"
            If DataExist(strSQL) Then
                Using sqlCmd As SqlCommand = New SqlCommand
                    With sqlCmd
                        .Connection = SqlConnectionclass.GetConnection
                        .CommandText = "USP_DISPLAY_CUSTITEM_MST_IN_GRID"
                        .CommandTimeout = 0
                        .CommandType = CommandType.StoredProcedure
                        .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                        .Parameters.Add("@CUSTOMER_CODE", SqlDbType.VarChar, 20).Value = txtCustCode.Text.Trim
                        .Parameters.Add("@MODE", SqlDbType.Char, 1).Value = "E"
                        .Parameters.Add("@ITEM_CODE", SqlDbType.VarChar, 16).Value = ""
                        .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                        Using dt As DataTable = SqlConnectionclass.GetDataTable(sqlCmd)
                            If dt.Rows.Count > 0 Then
                                xlApp = New Microsoft.Office.Interop.Excel.ApplicationClass
                                'xlApp.Visible = True
                                xlWorkBook = xlApp.Workbooks.Add(misValue)
                                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                                lblGenerate.Text = "Generating Excel Report...."
                                With xlWorkSheet.Range("A1", "P1")
                                    .Font.Bold = True
                                    .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                                    .EntireColumn.ColumnWidth = 20
                                End With
                                xlWorkSheet.Range("A1:P1").Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
                                xlWorkSheet.Cells(1, 1).Value = "CUSTOMER CODE"
                                xlWorkSheet.Cells(1, 2).Value = "ITEM CODE"
                                xlWorkSheet.Cells(1, 3).Value = "ITEM DESCRIPTION"
                                xlWorkSheet.Cells(1, 4).Value = "AUTO INVOICE PART"
                                xlWorkSheet.Cells(1, 5).Value = "CUST DRGNO."
                                xlWorkSheet.Cells(1, 6).Value = "DRAWING DESC"
                                xlWorkSheet.Cells(1, 7).Value = "BIN QUANTITY"
                                xlWorkSheet.Cells(1, 8).Value = "MODEL"
                                xlWorkSheet.Cells(1, 9).Value = "ACTIVE"
                                xlWorkSheet.Cells(1, 10).Value = "SCHEDULE UPLOAD REQUIRED"
                                xlWorkSheet.Cells(1, 11).Value = "COLOUR CODE"
                                xlWorkSheet.Cells(1, 12).Value = "CATEGORY CODE"
                                xlWorkSheet.Cells(1, 13).Value = "COMMODITY CODE"
                                xlWorkSheet.Cells(1, 14).Value = "CONTAINER"
                                xlWorkSheet.Cells(1, 15).Value = "PACKING LEVEL"
                                xlWorkSheet.Cells(1, 16).Value = "CUSTORD_STATUS"
                                xlWorkSheet.Cells(1, 17).Value = "UNIT CODE"
                                xlWorkSheet.Cells(1, 18).Value = "IP ADDRESS"
                                If gstrUNITID = "MS1" Or gstrUNITID = "MS2" Or gstrUNITID = "MK1 " Then  ''Anupam Kumar 
                                    xlWorkSheet.Cells(1, 19).Value = "DECLARATION NO"
                                End If

                                For i = 0 To dt.Rows.Count - 1
                                    For j = 0 To dt.Columns.Count - 3
                                        xlWorkSheet.Cells(i + 2, j + 1) =
                                    dt.Rows(i).Item(j)
                                    Next
                                    If dt.Rows(i).Item(15) = "TRUE" Then
                                        xlWorkSheet.Range("A" & i + 2 & "", "P" & i + 2 & "").Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
                                    Else
                                        xlWorkSheet.Range("A" & i + 2 & "", "P" & i + 2 & "").Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
                                    End If
                                Next

                                raXL = xlWorkSheet.UsedRange
                                raXL.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                raXL.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

                                'raXL = xlWorkSheet.Columns(6)
                                'raXL.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

                                With xlWorkSheet.Range("A1", "P1")
                                    .Font.Bold = True
                                    .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                                    .Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Cyan)
                                    .Borders.LineStyle = Excel.XlLineStyle.xlContinuous
                                End With
                                With xlWorkSheet.Range("P1", "P1")
                                    .EntireColumn.Hidden = True
                                End With
                                xlApp.Visible = True
                            End If
                        End Using
                    End With
                End Using
            Else
                MessageBox.Show("Invalid Customer...!!!")
                Exit Sub
            End If
            System.Windows.Forms.Cursor.Current = Cursors.Arrow
            xlApp.Quit()
            lblGenerate.Text = ""
            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
        End Try
    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub CmdGrpCustomerITem_Load(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub
End Class
