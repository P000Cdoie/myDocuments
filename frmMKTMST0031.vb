Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.Data.SqlClient
Friend Class frmMKTMST0031
    Inherits System.Windows.Forms.Form
    '---------------------------------------------------------------------------------------------------------
    'Copyright(c)       -   MIND
    'Name of module     -   frmMKTMST0031
    'Created by         -   Priyanka
    'Created Date       -   23/Dec/2019
    'Description        -   End Customer Linkage Master
    '---------------------------------------------------------------------------------------------------------


    Dim mintIndex As Short
    Dim blnCheckDelItems As Boolean
    Dim strDelete As String
    Dim strInsert As String
    Dim strupdate As String
    Private Enum GridEnum
        Sel = 1
        ITEMCODE
        Description
        Customer_Part_No
        Customer_Part_Description
        End_Customer_Name
        btnselect = 7
    End Enum

    Dim DocFrm As Form
    Dim dtDocTable As DataTable

    Private Sub AddlabelToGrid()
        On Error GoTo errHandler
        Me.spdITemDetails.MaxCols = 0
        Me.spdITemDetails.Row = FPSpreadADO.CoordConstants.SpreadHeader
        Me.spdITemDetails.set_RowHeight(Me.spdITemDetails.Row, 20)

        Me.spdITemDetails.MaxCols = Me.spdITemDetails.MaxCols + 1
        Me.spdITemDetails.Col = Me.spdITemDetails.MaxCols
        Me.spdITemDetails.Text = "Select"
        Me.spdITemDetails.set_ColWidth(Me.spdITemDetails.MaxCols, 3)

        Me.spdITemDetails.MaxCols = Me.spdITemDetails.MaxCols + 1
        Me.spdITemDetails.Col = Me.spdITemDetails.MaxCols
        Me.spdITemDetails.Text = "ITEM CODE"
        Me.spdITemDetails.set_ColWidth(Me.spdITemDetails.MaxCols, 15)

        Me.spdITemDetails.MaxCols = Me.spdITemDetails.MaxCols + 1
        Me.spdITemDetails.MaxCols = Me.spdITemDetails.MaxCols
        Me.spdITemDetails.Col = Me.spdITemDetails.MaxCols
        Me.spdITemDetails.Text = "Description"
        Me.spdITemDetails.set_ColWidth(Me.spdITemDetails.MaxCols, 20)

        Me.spdITemDetails.MaxCols = Me.spdITemDetails.MaxCols + 1
        Me.spdITemDetails.MaxCols = Me.spdITemDetails.MaxCols
        Me.spdITemDetails.Col = Me.spdITemDetails.MaxCols
        Me.spdITemDetails.Text = "Customer Part No"
        Me.spdITemDetails.set_ColWidth(Me.spdITemDetails.MaxCols, 10)

        Me.spdITemDetails.MaxCols = Me.spdITemDetails.MaxCols + 1
        Me.spdITemDetails.Col = Me.spdITemDetails.MaxCols
        Me.spdITemDetails.Text = "Customer Part Description"
        Me.spdITemDetails.set_ColWidth(Me.spdITemDetails.MaxCols, 20)

        Me.spdITemDetails.MaxCols = Me.spdITemDetails.MaxCols + 1
        Me.spdITemDetails.Col = Me.spdITemDetails.MaxCols
        Me.spdITemDetails.Text = "End Customer Name"
        Me.spdITemDetails.set_ColWidth(Me.spdITemDetails.MaxCols, 10)

        Me.spdITemDetails.MaxCols = Me.spdITemDetails.MaxCols + 1
        Me.spdITemDetails.Col = Me.spdITemDetails.MaxCols
        Me.spdITemDetails.Text = " "
        Me.spdITemDetails.set_ColWidth(Me.spdITemDetails.MaxCols, 4)

        Me.spdITemDetails.CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
        Me.spdITemDetails.EditEnterAction = FPSpreadADO.EditEnterActionConstants.EditEnterActionNext
        Me.spdITemDetails.ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsHorizontal
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub

    Private Sub AddNewRowType()
        On Error GoTo errHandler
        Me.spdITemDetails.MaxRows = Me.spdITemDetails.MaxRows + 1
        Me.spdITemDetails.Row = Me.spdITemDetails.MaxRows

        Me.spdITemDetails.Col = GridEnum.Sel
        Me.spdITemDetails.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox

        Me.spdITemDetails.Col = GridEnum.ITEMCODE
        Me.spdITemDetails.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
        Me.spdITemDetails.TypeEditLen = 50

        Me.spdITemDetails.Col = GridEnum.Description
        Me.spdITemDetails.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
        Me.spdITemDetails.TypeEditLen = 150

        Me.spdITemDetails.Col = GridEnum.Customer_Part_No
        Me.spdITemDetails.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
        Me.spdITemDetails.TypeEditLen = 150

        Me.spdITemDetails.Col = GridEnum.Customer_Part_Description
        Me.spdITemDetails.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
        Me.spdITemDetails.TypeEditLen = 150

        Me.spdITemDetails.Col = GridEnum.End_Customer_Name
        Me.spdITemDetails.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
        Me.spdITemDetails.TypeEditLen = 100

        Me.spdITemDetails.Col = GridEnum.btnselect
        Me.spdITemDetails.CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
        Me.spdITemDetails.TypeButtonPicture = My.Resources.ico111.ToBitmap
        'Me.spdITemDetails.Focus()
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
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
        Select Case Btngroup.Mode
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
                        Call txtCustCode_Validating(txtCustCode, New System.ComponentModel.CancelEventArgs(False))
                    End If
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                With Me.txtCustCode
                    strHelp = ShowList(1, .MaxLength, "", "a.Customer_Code", "b.CUST_NAME", "End_Cust_Item_Linkage a,Customer_mst b", " AND a.CUSTOMER_Code=b.Customer_Code AND a.UNIT_CODE=b.UNIT_CODE", "HELP", , , , , "a.UNIT_CODE")
                    .Focus()
                End With
                If Val(strHelp) = -1 Then ' No record
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                    txtCustCode.Focus()
                Else
                    Me.txtCustCode.Text = strHelp
                    rscusthelp.GetResult("SELECT distinct a.customer_Code,b.cust_NAME FROM End_Cust_Item_Linkage a inner join Customer_mst b on a.CUSTOMER_Code=b.Customer_Code and a.UNIT_CODE=b.UNIT_CODE AND a.CUSTOMER_code= '" & strHelp & "' AND a.UNIT_CODE='" & gstrUNITID & "' ")
                    If rscusthelp.GetNoRows > 0 Then 'RECORD FOUND
                        Me.lblCustCodeDes.Text = rscusthelp.GetValue("CUST_NAME")
                        Btngroup.Enabled(1) = True
                        ' Btngroup.Enabled(2) = True
                        Call txtCustCode_Validating(txtCustCode, New System.ComponentModel.CancelEventArgs(False))
                    End If
                End If
                'Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                '    With Me.txtCustCode
                '        strHelp = ShowList(1, .MaxLength, "", "a.Customer_Code", "b.CUST_NAME", "End_Cust_Item_Linkage a,Customer_mst b", " AND a.CUSTOMER_Code=b.Customer_Code AND a.UNIT_CODE=b.UNIT_CODE", "HELP", , , , , "a.UNIT_CODE")
                '        .Focus()
                '    End With
                '    If Val(strHelp) = -1 Then ' No record
                '        Call ConfirmWindow(10170, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                '        txtCustCode.Focus()
                '    Else
                '        Me.txtCustCode.Text = strHelp
                '        rscusthelp.GetResult("SELECT distinct a.customer_Code,b.cust_NAME FROM End_Cust_Item_Linkage a inner join Customer_mst b on a.CUSTOMER_Code=b.Customer_Code and a.UNIT_CODE=b.UNIT_CODE AND a.CUSTOMER_code= '" & strHelp & "' AND a.UNIT_CODE='" & gstrUNITID & "' ")
                '        If rscusthelp.GetNoRows > 0 Then 'RECORD FOUND
                '            Me.lblCustCodeDes.Text = rscusthelp.GetValue("CUST_NAME")
                '            Call txtCustCode_Validating(txtCustCode, New System.ComponentModel.CancelEventArgs(False))
                '        End If
                '    End If
        End Select
        rscusthelp.ResultSetClose()
        rscusthelp = Nothing
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0031_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo errHandler
        mdifrmMain.CheckFormName = mintIndex
        txtCustCode.Focus()
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTMST0031_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo errHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTMST0031_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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
    Private Sub frmMKTMST0031_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Try
            Select Case KeyAscii
                Case System.Windows.Forms.Keys.Return
                    System.Windows.Forms.SendKeys.SendWait("{TAB}") 'If user press the Enter Key ,the focus will be advanced
                Case System.Windows.Forms.Keys.Escape 'If user press Escape than valCancel will be callked.
                    Call EnableControls(True, Me, True)
                    txtCustCode.Enabled = True : txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelp.Enabled = True
                    txtCustCode.Focus()
                    If Btngroup.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                        Call Btngroup_ButtonClick(Btngroup, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL))
                    End If
            End Select

        Catch ex As Exception
            RaiseException(ex)
            gblnCancelUnload = True : gblnFormAddEdit = True
        Finally
            eventArgs.KeyChar = Chr(KeyAscii)
            If KeyAscii = 0 Then
                eventArgs.Handled = True
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub frmMKTMST0031_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo errHandler
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FitToClient(Me, Frame1, ctlFormHeader1, Btngroup, 200)
        'Set Help Pictures At Command Button
        cmdHelp.Image = My.Resources.ico111.ToBitmap
        'Initially Disable All Controls
        Call EnableControls(True, Me, True)
        Call AddlabelToGrid()
        Clearscreen()
      
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0031_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum

        Try
            If UnloadMode >= 0 And UnloadMode <= 5 Then
                If Me.Btngroup.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or Me.Btngroup.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then 'If not View Mode
                    enmValue = ConfirmWindow(10055, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) 'Confirm before unloading the FORM
                    If enmValue <> eMPowerFunctions.ConfirmWindowReturnEnum.VAL_CANCEL Then 'If  'YES' or 'NO'
                        If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then 'If YES
                            Call Btngroup_ButtonClick(Btngroup, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE))
                        ElseIf enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Then  'If NO than Unload the Form
                            gblnCancelUnload = False 'Variable used in MDI Form before unloading
                            gblnFormAddEdit = False
                            Me.Btngroup.Focus()
                        Else
                            gblnCancelUnload = True : gblnFormAddEdit = True ' If Cancel than Focus will be set on in the first field.
                            Me.Btngroup.Focus()
                        End If
                    Else
                        txtCustCode.Focus()
                        gblnCancelUnload = True
                        gblnFormAddEdit = True
                    End If
                Else
                    Me.Dispose()
                    Exit Sub
                End If
            End If
            If gblnCancelUnload Then eventArgs.Cancel = True 'Do not unload FORM, if the value of gblncancelUnload is False
        Catch ex As Exception
            RaiseException(ex)
            gblnCancelUnload = True : gblnFormAddEdit = True
            eventArgs.Cancel = Cancel
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub frmMKTMST0031_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo errHandler
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Me.Dispose()
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
            TxtSearch.Text = ""
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
                If spdITemDetails.Enabled = True Then spdITemDetails.Focus()
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
        Select Case Btngroup.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                CmbSearchBy.Enabled = True
                TxtSearch.Enabled = True

                CmbSearchBy.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                TxtSearch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)

                CmbSearchBy.SelectedIndex = 0
                TxtSearch.Text = ""

                If IsExistFieldValue(Trim(txtCustCode.Text), "customer_code", "Customer_Mst", "UNIT_CODE='" & gstrUNITID & "'") = True Then
                    lblCustCodeDes.Text = ReturnDescription("cust_name", "customer_mst", "customer_Code ='" & Trim(txtCustCode.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'")
                    spdITemDetails.MaxRows = 0
                    Call DisplayDetailsinGrid()
                    If spdITemDetails.MaxRows > 0 Then
                        With spdITemDetails
                            .Col = GridEnum.Sel : .ColHidden = False
                            .Col = GridEnum.btnselect : .ColHidden = False
                        End With
                    End If
                Else
                    MsgBox("Invalid Customer Code", MsgBoxStyle.Information, "eMPro")
                    txtCustCode.Text = "" : txtCustCode.Focus()
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                CmbSearchBy.Enabled = True
                TxtSearch.Enabled = True
                CmbSearchBy.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                TxtSearch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                CmbSearchBy.SelectedIndex = 0
                TxtSearch.Text = ""

                If IsExistFieldValue(Trim(txtCustCode.Text), "a.customer_code", "End_Cust_Item_Linkage a,Customer_Mst b", "a.CUSTOMER_Code = b.CUSTOMER_Code AND A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" & gstrUNITID & "'") = True Then
                    spdITemDetails.MaxRows = 0
                    Call DisplayDetailsinGrid_View()
                    Btngroup.Enabled(1) = True
                    If spdITemDetails.MaxRows > 0 Then
                        With spdITemDetails
                            .Col = GridEnum.Sel : .ColHidden = True
                            .Col = GridEnum.btnselect : .ColHidden = True
                        End With
                    End If
                    'spdITemDetails.ColHidden = GridEnum.Sel
                Else
                    MsgBox("Invalid Customer Code", MsgBoxStyle.Information, "eMPro")
                    txtCustCode.Text = "" : txtCustCode.Focus()
                End If
                'Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                '    CmbSearchBy.Enabled = True
                '    TxtSearch.Enabled = True
                '    CmbSearchBy.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                '    TxtSearch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                '    CmbSearchBy.SelectedIndex = 0
                '    TxtSearch.Text = ""

                '    If IsExistFieldValue(Trim(txtCustCode.Text), "a.customer_code", "End_Cust_Item_Linkage a,Customer_Mst b", "a.CUSTOMER_Code = b.CUSTOMER_Code AND A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" & gstrUNITID & "'") = True Then
                '        spdITemDetails.MaxRows = 0
                '        Call DisplayDetailsinGrid_View()
                '        If spdITemDetails.MaxRows > 0 Then
                '            With spdITemDetails
                '                .Col = GridEnum.Sel : .ColHidden = False
                '                .Col = GridEnum.btnselect : .ColHidden = False
                '            End With
                '        End If
                '    Else
                '        MsgBox("Invalid Customer Code", MsgBoxStyle.Information, "eMPro")
                '        txtCustCode.Text = "" : txtCustCode.Focus()
                '    End If
        End Select
                GoTo EventExitSub
errHandler:
                Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
                GoTo EventExitSub
EventExitSub:
                eventArgs.Cancel = Cancel
    End Sub
    Public Function DisplayDetailsinGrid() As Object
        Dim DataRd As SqlDataReader
        Dim Sqlcmd As New SqlCommand
        Dim ObjVal As Object = Nothing
        Dim intLoopCounter As Short
        Dim intMaxCounter As Short
        Dim StrItemCode As String
        Dim strcustcode As String
        Dim strSql As String = String.Empty
        Try
            strcustcode = txtCustCode.Text.Trim
            strSql = "Select * from CustITem_Mst  Where UNIT_CODE=@UNIT_CODE AND Account_code =@cust_code and Item_code not in (select Item_code from End_Cust_Item_Linkage where UNIT_CODE=@UNIT_CODE AND Customer_Code =@cust_code ) ORDER BY Cust_drgNo"
            Sqlcmd.CommandTimeout = 0
            Sqlcmd.Connection = SqlConnectionclass.GetConnection
            Sqlcmd.CommandType = CommandType.Text
            Sqlcmd.Parameters.AddWithValue("@cust_code", txtCustCode.Text.Trim)
            Sqlcmd.Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
            Sqlcmd.CommandText = strSql
            Me.spdITemDetails.MaxRows = 0
            DataRd = Sqlcmd.ExecuteReader()
            If DataRd.HasRows() Then
                While DataRd.Read()
                    AddNewRowType()
                    Me.spdITemDetails.Row = spdITemDetails.MaxRows
                    Me.spdITemDetails.Row2 = spdITemDetails.MaxRows
                    Me.spdITemDetails.Col = 1
                    Me.spdITemDetails.Col2 = spdITemDetails.MaxCols

                    Me.spdITemDetails.Row = spdITemDetails.MaxRows
                    Me.spdITemDetails.Col = GridEnum.Sel
                    Me.spdITemDetails.Value = 0

                    Me.spdITemDetails.Col = GridEnum.ITEMCODE
                    Me.spdITemDetails.Text = DataRd("Item_code").ToString
                    StrItemCode = DataRd("Item_code").ToString
                    Me.spdITemDetails.Col = GridEnum.Description
                    Me.spdITemDetails.Text = DataRd("Item_Desc").ToString


                    Me.spdITemDetails.Col = GridEnum.Customer_Part_No
                    Me.spdITemDetails.Text = DataRd("Cust_DrgNo").ToString

                    Me.spdITemDetails.Col = GridEnum.Customer_Part_Description
                    Me.spdITemDetails.Text = DataRd("Drg_desc").ToString

                    Me.spdITemDetails.Col = GridEnum.End_Customer_Name
                    Me.spdITemDetails.Text = " "
                    Me.spdITemDetails.Col = GridEnum.btnselect
                End While

                Me.spdITemDetails.Row = 1
                Me.spdITemDetails.Row2 = Me.spdITemDetails.MaxRows
                Me.spdITemDetails.Col = 1
                Me.spdITemDetails.Col2 = Me.spdITemDetails.MaxCols
            End If
            If DataRd.IsClosed = False Then DataRd.Close()
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            If IsNothing(Sqlcmd.Connection) = False Then
                If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
                Sqlcmd.Connection.Dispose()
            End If
            If IsNothing(Sqlcmd) = False Then
                Sqlcmd.Dispose()
            End If
        End Try  
    End Function

    Public Function DisplayDetailsinGrid_View() As Object
        Dim DataRd As SqlDataReader
        Dim Sqlcmd As New SqlCommand
        Dim ObjVal As Object = Nothing
        Dim intLoopCounter As Short
        Dim intMaxCounter As Short
        Dim StrItemCode As String
        Dim strcustcode As String
        Dim strSql As String = String.Empty
        Try
            strcustcode = txtCustCode.Text.Trim
            strSql = "Select * from End_Cust_Item_Linkage Where UNIT_CODE=@UNIT_CODE AND CUSTOMER_Code =@cust_code ORDER BY Cust_drgNo"
            Sqlcmd.CommandTimeout = 0
            Sqlcmd.Connection = SqlConnectionclass.GetConnection
            Sqlcmd.CommandType = CommandType.Text
            Sqlcmd.Parameters.AddWithValue("@cust_code", txtCustCode.Text.Trim)
            Sqlcmd.Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
            Sqlcmd.CommandText = strSql
            Me.spdITemDetails.MaxRows = 0
            DataRd = Sqlcmd.ExecuteReader()
            If DataRd.HasRows() Then
                While DataRd.Read()
                    AddNewRowType()
                    Me.spdITemDetails.Row = spdITemDetails.MaxRows
                    Me.spdITemDetails.Row2 = spdITemDetails.MaxRows
                    Me.spdITemDetails.Col = 1
                    Me.spdITemDetails.Col2 = spdITemDetails.MaxCols

                    Me.spdITemDetails.Row = spdITemDetails.MaxRows
                    Me.spdITemDetails.Col = GridEnum.Sel
                    Me.spdITemDetails.Value = 0

                    Me.spdITemDetails.Col = GridEnum.ITEMCODE
                    Me.spdITemDetails.Text = DataRd("Item_code").ToString
                    StrItemCode = DataRd("Item_code").ToString
                    Me.spdITemDetails.Col = GridEnum.Description
                    Me.spdITemDetails.Text = DataRd("Description").ToString


                    Me.spdITemDetails.Col = GridEnum.Customer_Part_No
                    Me.spdITemDetails.Text = DataRd("Cust_DrgNo").ToString

                    Me.spdITemDetails.Col = GridEnum.Customer_Part_Description
                    Me.spdITemDetails.Text = DataRd("Drg_desc").ToString

                    Me.spdITemDetails.Col = GridEnum.End_Customer_Name
                    Me.spdITemDetails.Text = DataRd("End_Customer_Name").ToString
                    Me.spdITemDetails.Col = GridEnum.btnselect

                    'Select Case Btngroup.Mode
                    '    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                    '        With spdITemDetails
                    '            .Col = GridEnum.btnselect
                    '        End With
                    'End Select
                End While

                Me.spdITemDetails.Row = 1
                Me.spdITemDetails.Row2 = Me.spdITemDetails.MaxRows
                Me.spdITemDetails.Col = 1
                Me.spdITemDetails.Col2 = Me.spdITemDetails.MaxCols
            End If
            If DataRd.IsClosed = False Then DataRd.Close()
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            If IsNothing(Sqlcmd.Connection) = False Then
                If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
                Sqlcmd.Connection.Dispose()
            End If
            If IsNothing(Sqlcmd) = False Then
                Sqlcmd.Dispose()
            End If
        End Try
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
    '            If Btngroup.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
    '                Call .SetText(5, .MaxRows, "A")
    '            End If
    '            If .MaxRows > 6 Then .ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsBoth
    '        End With
    '        Call AddNewRowType()
    '        Exit Sub
    'errHandler:
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '    End Sub
   
    Private Sub spdITemDetails_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles spdITemDetails.ButtonClicked
        On Error GoTo ErrHandler
        Dim strSQL As String = String.Empty
        If e.col = GridEnum.btnselect Then
            Dim strHelp() As String
            Dim strCustomerLinkage As String = String.Empty
            Dim varHelpCustomerLinkage As Object
            strSQL = "Select Key2 as End_customer_list,Descr as Description from lists where UNIT_CODE='" & gstrUNITID & "' and key1='End Customer'"
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "END Customer Help", 1)
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            If UBound(strHelp) <> -1 Then
                If strHelp(0) <> "0" Then
                    strCustomerLinkage = strHelp(0)
                    With spdITemDetails
                        .Row = e.row
                        .Col = GridEnum.End_Customer_Name
                        .Text = strCustomerLinkage
                        .Col = GridEnum.Sel
                        .Value = 1
                    End With
                Else
                    MsgBox("No Record Available", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                End If
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    '    Private Sub spdITemDetails_KeyDownEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles spdITemDetails.KeyDownEvent
    '        On Error GoTo ErrHandler
    '        Select Case e.keyCode
    '            Case Keys.F1
    '                If spdITemDetails.ActiveCol = 2 Then
    '                    If Btngroup.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT And spdITemDetails.ActiveRow < spdITemDetails.MaxRows Then Exit Sub
    '                    Call spdITemDetails_ButtonClicked(sender, New AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent(3, spdITemDetails.ActiveRow, 0))
    '                End If

    '            Case 13
    '                If spdITemDetails.ActiveCol = 14 Then
    '                    If Btngroup.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
    '                        '***If this is Last Row & Last Column
    '                        If spdITemDetails.ActiveRow = spdITemDetails.MaxRows Then
    '                            Call addNewInSpread()
    '                        End If
    '                        '***
    '                    End If
    '                End If
    '        End Select
    '        Exit Sub
    'ErrHandler:
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '    End Sub
    '    Private Sub spdITemDetails_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles spdITemDetails.KeyPressEvent
    '        On Error GoTo ErrHandler
    '        Select Case e.keyAscii
    '            Case Keys.Return
    '            Case 39, 34, 96
    '                e.keyAscii = 0
    '        End Select
    '        Exit Sub
    'ErrHandler:
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '    End Sub
    '    Private Sub spdITemDetails_KeyUpEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyUpEvent) Handles spdITemDetails.KeyUpEvent
    '        On Error GoTo ErrHandler
    '        If ((e.shift = 2) And (e.keyCode = Keys.N)) Then
    '            Call addNewInSpread()
    '            Call spdITemDetails_LeaveCell(sender, New AxFPSpreadADO._DSpreadEvents_LeaveCellEvent(spdITemDetails.ActiveCol, spdITemDetails.MaxRows - 1, 3, spdITemDetails.MaxRows, False))
    '        End If
    '        Exit Sub
    'ErrHandler:
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '    End Sub
    '    Private Sub spdITemDetails_LeaveCell(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles spdITemDetails.LeaveCell
    '        On Error GoTo ErrHandler
    '        Dim varFeild As Object
    '        Dim varItemCode As Object

    '        Exit Sub
    'ErrHandler:
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '        Exit Sub
    '    End Sub
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")  '("machine_master.htm")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub CmbSearchBy_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbSearchBy.TextChanged

        On Error GoTo ErrHandler

        If TxtSearch.Enabled Then
            TxtSearch.Focus()
        End If

        Exit Sub
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub
    'ISSUE ID       : 10195532
    Private Sub CmbSearchBy_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbSearchBy.SelectedIndexChanged

        On Error GoTo ErrHandler

        TxtSearch.Text = ""
        If TxtSearch.Enabled Then
            TxtSearch.Focus()
        End If

        Exit Sub
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub
    'ISSUE ID       : 10195532
    Private Sub TxtSearch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtSearch.TextChanged

        On Error GoTo ErrHandler

        Call SearchItem()

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub
    'ISSUE ID       : 10195532
    Private Sub SearchItem()

        On Error GoTo ErrHandler

        Dim intRowNo As Long
        Dim strSearchBy As String = ""
        Dim strSearchItem As Object = Nothing

        If Len(Trim(CmbSearchBy.Text)) = 0 Then
            If CmbSearchBy.Enabled Then CmbSearchBy.Focus()
            Exit Sub
        End If

        If Len(Trim(TxtSearch.Text)) = 0 Then
            spdITemDetails.TopRow = 1
            If TxtSearch.Enabled Then TxtSearch.Focus()
            Exit Sub
        End If

        strSearchBy = CmbSearchBy.Text

        For intRowNo = 1 To spdITemDetails.MaxRows
            If strSearchBy = "Item Code" Then
                strSearchItem = Nothing
                Call spdITemDetails.GetText(2, intRowNo, strSearchItem)
            ElseIf strSearchBy = "Item Desc" Then
                strSearchItem = Nothing
                Call spdITemDetails.GetText(3, intRowNo, strSearchItem)
            ElseIf strSearchBy = "Customer Part No." Then
                strSearchItem = Nothing
                Call spdITemDetails.GetText(4, intRowNo, strSearchItem)
            ElseIf strSearchBy = "Customer Part Desc" Then
                strSearchItem = Nothing
                Call spdITemDetails.GetText(5, intRowNo, strSearchItem)
            End If

            If UCase(strSearchItem) Like UCase(TxtSearch.Text) & "*" Then
                spdITemDetails.TopRow = intRowNo
                Exit For
            End If
        Next intRowNo

        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub

    Private Function Validate_Data() As Boolean
        Validate_Data = False
        Dim strsel As String = String.Empty
        Dim strEndcustomer As String = String.Empty
        Dim sel As Object = Nothing
        Dim EndCustomer As Object = Nothing
        Dim intZ As Integer
        For intZ = 1 To Me.spdITemDetails.MaxRows
            strsel = String.Empty
            sel = Nothing
            strEndcustomer = String.Empty
            EndCustomer = Nothing
            Call Me.spdITemDetails.GetText(GridEnum.Sel, intZ, sel)
            Call Me.spdITemDetails.GetText(GridEnum.End_Customer_Name, intZ, EndCustomer)
            strsel = sel.ToString()
            strEndcustomer = EndCustomer.ToString()
            If strsel = "1" And strEndcustomer <> "" Then
                Validate_Data = True
                Exit Function
            End If
        Next intZ
    End Function
    Private Function Save_Data() As Boolean
        Dim Sqlcmd As New SqlCommand
        Dim intY As Integer
        Dim intX As Integer
        Dim varstr_select As String = String.Empty
        Dim EndCustomer_name As String = String.Empty
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim varDesc As Object
        Dim varItemDesc As Object
        Dim VarEndCustomerName As Object = Nothing
        Dim varSelect As Object = Nothing
        Dim SqlTrans As SqlTransaction
        Try
            Sqlcmd.CommandTimeout = 0
            Sqlcmd.Connection = SqlConnectionclass.GetConnection
            Sqlcmd.CommandType = CommandType.StoredProcedure
            SqlTrans = Sqlcmd.Connection.BeginTransaction()
            Sqlcmd.Transaction = SqlTrans

            For intX = 1 To Me.spdITemDetails.MaxRows
                varItemCode = Nothing
                varItemDesc = Nothing
                varDrgNo = Nothing
                varDesc = Nothing
                VarEndCustomerName = Nothing
                varSelect = Nothing
                varstr_select = String.Empty
                EndCustomer_name = String.Empty
                Call Me.spdITemDetails.GetText(GridEnum.ITEMCODE, intX, varItemCode)
                Call Me.spdITemDetails.GetText(GridEnum.Description, intX, varItemDesc)
                Call Me.spdITemDetails.GetText(GridEnum.Customer_Part_No, intX, varDrgNo)
                Call Me.spdITemDetails.GetText(GridEnum.Customer_Part_Description, intX, varDesc)
                Call Me.spdITemDetails.GetText(GridEnum.End_Customer_Name, intX, VarEndCustomerName)
                Call Me.spdITemDetails.GetText(GridEnum.Sel, intX, varSelect)
                varstr_select = varSelect.ToString()
                EndCustomer_name = VarEndCustomerName.ToString()
                Sqlcmd.Parameters.Clear()
                Sqlcmd.CommandTimeout = 0
                Sqlcmd.CommandType = CommandType.StoredProcedure
                If varSelect <> Nothing And varstr_select = "1" And VarEndCustomerName.Trim <> "" And EndCustomer_name.Trim <> String.Empty And VarEndCustomerName <> Nothing Then
                    Sqlcmd.Parameters.AddWithValue("@Customer_Code", txtCustCode.Text.Trim)
                    Sqlcmd.Parameters.AddWithValue("@Unit_Code", gstrUNITID)
                    Sqlcmd.Parameters.AddWithValue("@Item_Code", varItemCode)
                    Sqlcmd.Parameters.AddWithValue("@Description", Replace(Trim(varItemDesc), "'", ""))
                    Sqlcmd.Parameters.AddWithValue("@Cust_drgno", varDrgNo)
                    Sqlcmd.Parameters.AddWithValue("@Drg_Desc", varDesc)
                    Sqlcmd.Parameters.AddWithValue("@End_Customer_Name", VarEndCustomerName)
                    Sqlcmd.Parameters.AddWithValue("@Ent_Dt", getDateForDB(GetServerDate()))
                    Sqlcmd.Parameters.AddWithValue("@Ent_UserID", mP_User)
                    Sqlcmd.Parameters.AddWithValue("@MODE", "SAVE")
                    Sqlcmd.CommandText = "USP_SAVE_END_CUST_ITEM_LINKAGE"
                    intY = Sqlcmd.ExecuteNonQuery()
                End If
            Next intX
            SqlTrans.Commit()
            Save_Data = True
        Catch ex As Exception
            If Not IsNothing(SqlTrans) Then
                SqlTrans.Rollback()
            End If
            RaiseException(ex)
        End Try
    End Function

    Private Function Update_Data() As Boolean
        Dim Sqlcmd As New SqlCommand
        Dim intY As Integer
        Dim intX As Integer
        Dim varstr_select As String = String.Empty
        Dim EndCustomer_name As String = String.Empty
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim varDesc As Object
        Dim varItemDesc As Object
        Dim VarEndCustomerName As Object = Nothing
        Dim varSelect As Object = Nothing
        Dim SqlTrans As SqlTransaction
        Try
            Sqlcmd.CommandTimeout = 0
            Sqlcmd.Connection = SqlConnectionclass.GetConnection
            Sqlcmd.CommandType = CommandType.StoredProcedure
            SqlTrans = Sqlcmd.Connection.BeginTransaction()
            Sqlcmd.Transaction = SqlTrans

            For intX = 1 To Me.spdITemDetails.MaxRows
                varItemCode = Nothing
                varItemDesc = Nothing
                varDrgNo = Nothing
                varDesc = Nothing
                VarEndCustomerName = Nothing
                varSelect = Nothing
                varstr_select = String.Empty
                EndCustomer_name = String.Empty
                Call Me.spdITemDetails.GetText(GridEnum.ITEMCODE, intX, varItemCode)
                Call Me.spdITemDetails.GetText(GridEnum.Description, intX, varItemDesc)
                Call Me.spdITemDetails.GetText(GridEnum.Customer_Part_No, intX, varDrgNo)
                Call Me.spdITemDetails.GetText(GridEnum.Customer_Part_Description, intX, varDesc)
                Call Me.spdITemDetails.GetText(GridEnum.End_Customer_Name, intX, VarEndCustomerName)
                Call Me.spdITemDetails.GetText(GridEnum.Sel, intX, varSelect)
                varstr_select = varSelect.ToString()
                EndCustomer_name = VarEndCustomerName.ToString()
                Sqlcmd.Parameters.Clear()
                Sqlcmd.CommandTimeout = 0
                Sqlcmd.CommandType = CommandType.StoredProcedure
                If varSelect <> Nothing And varstr_select = "1" And VarEndCustomerName.Trim <> "" And EndCustomer_name.Trim <> String.Empty And VarEndCustomerName <> Nothing Then
                    If Not IsRecordExists("SELECT Item_Code from End_Cust_Item_Linkage(nolock) where Customer_Code='" + txtCustCode.Text.Trim + "' and Unit_Code='" + gstrUNITID + "' and Item_Code='" + varItemCode + "'") Then
                        SqlTrans.Rollback()
                        MessageBox.Show("Selected ITEM CODE NOT EXIST FOR UPDATE.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Return False
                    End If
                    Sqlcmd.Parameters.AddWithValue("@Customer_Code", txtCustCode.Text.Trim)
                    Sqlcmd.Parameters.AddWithValue("@Unit_Code", gstrUNITID)
                    Sqlcmd.Parameters.AddWithValue("@Item_Code", varItemCode)
                    Sqlcmd.Parameters.AddWithValue("@Description", Replace(Trim(varItemDesc), "'", ""))
                    Sqlcmd.Parameters.AddWithValue("@Cust_drgno", varDrgNo)
                    Sqlcmd.Parameters.AddWithValue("@Drg_Desc", varDesc)
                    Sqlcmd.Parameters.AddWithValue("@End_Customer_Name", VarEndCustomerName)
                    Sqlcmd.Parameters.AddWithValue("@Ent_Dt", getDateForDB(GetServerDate()))
                    Sqlcmd.Parameters.AddWithValue("@Ent_UserID", mP_User)
                    Sqlcmd.Parameters.AddWithValue("@MODE", "UPDATE")
                    Sqlcmd.CommandText = "USP_SAVE_END_CUST_ITEM_LINKAGE"
                    intY = Sqlcmd.ExecuteNonQuery()
                End If
            Next intX
            SqlTrans.Commit()
            Update_Data = True
        Catch ex As Exception
            If Not IsNothing(SqlTrans) Then
                SqlTrans.Rollback()
            End If
            RaiseException(ex)
        End Try
    End Function
    Private Sub Btngroup_ButtonClick(ByVal Sender As System.Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles Btngroup.ButtonClick
        Try
            Select Case e.Button
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD 'If user selectes Add Button
                    Call Clearscreen()
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT ' If user selects Edit Button
                    'Call Clearscreen()
                    If spdITemDetails.MaxRows > 0 Then
                        With spdITemDetails
                            .Col = GridEnum.Sel : .ColHidden = False
                            .Col = GridEnum.btnselect : .ColHidden = False
                        End With
                    End If

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE ' If user selects Delete Button 
                    MsgBox("Delete Functionality Not Available.", MsgBoxStyle.Information, ResolveResString(100))
                    Btngroup.Revert()
                    Btngroup.Enabled(1) = False
                    Btngroup.Enabled(2) = False
                    Btngroup.Enabled(5) = False
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL 'If user selects Cancel than valCancel will be called.
                    Call Clearscreen()
                    Btngroup.Revert()
                    Btngroup.Enabled(1) = False
                    Btngroup.Enabled(2) = False
                    Btngroup.Enabled(5) = False
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE ' If user selects Close Boutton than Unload the form & Query Unload will be gets activated.
                    Me.Close()
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE ' For saving Addition/Updation
                    Select Case Me.Btngroup.Mode
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD ' For inserting Addrec will be called
                            Try
                                If Me.txtCustCode.Text.Trim = String.Empty Then
                                    MsgBox("Please Select Customer Code", MsgBoxStyle.Information, ResolveResString(100))
                                    txtCustCode.Focus()
                                    Exit Sub
                                ElseIf Me.spdITemDetails.MaxRows = 0 Then
                                    MsgBox("No Items For End Customer Linkage.", MsgBoxStyle.Information, ResolveResString(100))
                                    Exit Sub
                                End If
                                If Validate_Data() = False Then
                                    MsgBox("Please select atleast one record.", MsgBoxStyle.Information, ResolveResString(100))
                                    Exit Sub
                                End If
                            Catch Ex As Exception
                                MsgBox(Ex.ToString, MsgBoxStyle.Information, ResolveResString(100))
                            End Try

                            Try
                                If Save_Data() = True Then
                                    Me.spdITemDetails.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    MsgBox("Record Saved Successfully.", MsgBoxStyle.Information, ResolveResString(100))
                                    txtCustCode.Text = ""
                                    spdITemDetails.MaxRows = 0
                                    Btngroup.Revert()
                                    Btngroup.Enabled(1) = False
                                    Btngroup.Enabled(2) = False
                                    Btngroup.Enabled(5) = False
                                End If
                            Catch Ex As Exception
                                MsgBox(Ex.Message.ToString, MsgBoxStyle.Information, ResolveResString(100))
                            End Try
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT ' For Updation EditRec will ba called.
                            Try
                                If Me.txtCustCode.Text.Trim = String.Empty Then
                                    MsgBox("Please Select Customer Code", MsgBoxStyle.Information, ResolveResString(100))
                                    txtCustCode.Focus()
                                    Exit Sub
                                ElseIf Me.spdITemDetails.MaxRows = 0 Then
                                    MsgBox("No Items For End Customer Linkage.", MsgBoxStyle.Information, ResolveResString(100))
                                    Exit Sub
                                End If
                                If Validate_Data() = False Then
                                    MsgBox("Please select atleast one record.", MsgBoxStyle.Information, ResolveResString(100))
                                    Exit Sub
                                End If
                            Catch Ex As Exception
                                MsgBox(Ex.ToString, MsgBoxStyle.Information, ResolveResString(100))
                            End Try

                            Try
                                If Update_Data() = True Then
                                    Me.spdITemDetails.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    MsgBox("Record Update Successfully.", MsgBoxStyle.Information, ResolveResString(100))
                                    txtCustCode.Text = ""
                                    spdITemDetails.MaxRows = 0
                                    Btngroup.Revert()
                                    Btngroup.Enabled(1) = False
                                    Btngroup.Enabled(2) = False
                                    Btngroup.Enabled(5) = False
                                End If
                            Catch Ex As Exception
                                MsgBox(Ex.Message.ToString, MsgBoxStyle.Information, ResolveResString(100))
                            End Try
                    End Select
            End Select
        Catch ex As Exception
            RaiseException(ex)
            gblnCancelUnload = True : gblnFormAddEdit = True
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub Clearscreen()

        Try
            txtCustCode.Text = ""
            TxtSearch.Text = ""
            spdITemDetails.MaxRows = 0
            txtCustCode.Enabled = True
            txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            cmdHelp.Enabled = True
            Btngroup.Enabled(1) = False
            Btngroup.Enabled(2) = False
            Btngroup.Enabled(5) = False
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub
End Class
