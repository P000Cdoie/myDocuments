Option Strict Off
Option Explicit On
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility.VB6
Imports System.Data.SqlClient

Friend Class frmMKTTRN0080
    Inherits System.Windows.Forms.Form
    '***************************************************************************************
    'COPYRIGHT       : MIND LTD.
    'MODULE          : FRMMKTTRNP0080 - CUMMULATIVE QUANTITY ADJUSTMENT AUTHORIZATION
    'AUTHOR          : PRASHANT RAJPAL
    'CREATION DATE   : 12-JULY 2013- 20 JUL 2013
    ' ISSUE ID       : 10416813 
    'PURPOSE          : CUMULATIVE ASN FUNCTIONLITY -TRANSACTION FORM FOR AUTHOIRZING CDR NOS
    '******************************************************************************************

    Private mlngFormTag As Integer
    Private mServerDate As String
    Private Enum enmGrid
        col_Part_Code = 1
        col_Item_Code = 2
        col_Item_UOM = 3
        col_Item_CummsBeforeAdjustment = 4
        col_Item_CummsDiff = 5
        col_Item_Nature = 6
        col_Item_NewCumms = 7
        col_Invoice = 8
        col_CDR_Reference = 9
    End Enum

    Private Sub cmdDocCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDocCode.Click
        Dim StrDocHelp As String
        On Error GoTo ErrHandler
        StrDocHelp = ShowList(0, Len(txtDocCode.Text), txtDocCode.Text, "Doc_No", "" & DateColumnNameInShowList("Trans_date") & " as Trans_Date", "ASN_CUMMSADJST_HDR", "  AND AUTHORIZED_CODE IS NULL ", "ASN CUMMS ADJUSTED SERIES ")
        If StrDocHelp = "-1" Then

            Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            With Me.txtDocCode
                .Enabled = True
                .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                .Focus()
            End With
            cmdDocCode.Enabled = True

        ElseIf StrDocHelp = "" Then
            Me.txtDocCode.Focus()
        Else
            Me.txtDocCode.Text = StrDocHelp
            txtDocCode_Validating(txtDocCode, New System.ComponentModel.CancelEventArgs(False))
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub ctlStckAdjstHeader_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ctlStckAdjstHeader.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub frmMKTTRN0080_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mlngFormTag
        frmModules.NodeFontBold(Me.Tag) = True
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0080_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0080_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If (KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0) Then Call ctlStckAdjstHeader_ClickEvent(ctlStckAdjstHeader, New System.EventArgs())
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub frmMKTTRN0080_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                If (Me.CmdStckAdjustBttn.Mode) <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call Me.CmdStckAdjustBttn.Revert() : gblnCancelUnload = False : gblnFormAddEdit = False
                        Call EnableControls(False, Me, True)
                        SpStckAdjst.Enabled = True


                        With txtCustomerCode
                            .Enabled = False
                            .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            .Focus()
                        End With

                        With Me.txtDocCode
                            .Enabled = True
                            .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            .Focus()
                        End With
                        cmdDocCode.Enabled = True
                        Call RefreshControls()
                        GoTo EventExitSub
                    Else
                        Me.ActiveControl.Focus()
                        GoTo EventExitSub
                    End If
                End If
        End Select
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmMKTTRN0080_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        mlngFormTag = mdifrmMain.AddFormNameToWindowList(Me.ctlStckAdjstHeader.Tag)
        Call FitToClient(Me, fraMain, ctlStckAdjstHeader, CmdStckAdjustBttn, 400)
        CmdStckAdjustBttn.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(fraMain.Left) + (VB6.PixelsToTwipsX(fraMain.Width) / 4))
        Call EnableControls(False, Me, True)
        mServerDate = getDateForDB(GetServerDate())
        Call InitializeControls()
        cmdauthorize.Enabled = True
        cmdReject.Enabled = True
        CmdClose.Enabled = True

        txtLocationCode.Text = gstrUNITID
        lbllocationdesc.Text = GetQryOutput("SELECT Description FROM Location_Mst WHERE UNIT_CODE='" & gstrUNITID & "' AND LOCATION_CODE='" & gstrUNITID & "'")
        SpStckAdjst.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)

        With txtDocCode
            .Enabled = True
            .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            .Focus()
        End With
        cmdDocCode.Enabled = True
        gblnCancelUnload = False
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub InitializeControls()
        On Error GoTo ErrHandler
        With txtLocationCode
            .Enabled = False
            .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        End With

        lblDisplayDate.Text = VB6.Format(mServerDate, gstrDateFormat)
        With CmdStckAdjustBttn
            .Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        End With
        Call CmdStckAdjustBttn.ShowButtons(True, False, False, True)
        Call SetGridHdrs()
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0080_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        On Error GoTo ErrHandler
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
            If CmdStckAdjustBttn.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                enmValue = ConfirmWindow(10055, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Or enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then

                    ElseIf enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Then
                        gblnCancelUnload = False : gblnFormAddEdit = False
                    End If
                Else

                    gblnCancelUnload = True : gblnFormAddEdit = True
                    Me.ActiveControl.Focus()
                End If
            Else

                Exit Sub
            End If
        End If

        If gblnCancelUnload = True Then eventArgs.Cancel = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTTRN0080_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        Me.Dispose()
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mlngFormTag
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub SpStckAdjst_Change(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ChangeEvent) Handles SpStckAdjst.Change
        On Error GoTo ErrHandler
        Dim intNoOfBatch As Integer
        Dim dblStkBeforeAdjustment As Double
        Dim dblcumms_AdstValue As Double
        Dim strnaturetype As String
        Select Case e.col
            Case enmGrid.col_Item_Nature
                With SpStckAdjst
                    .Row = e.row
                    .Col = enmGrid.col_Item_CummsBeforeAdjustment
                    dblStkBeforeAdjustment = Val(.Text)
                    .Col = enmGrid.col_Item_CummsDiff
                    dblcumms_AdstValue = Val(.Text)
                    .Col = enmGrid.col_Item_Nature
                    strnaturetype = .Text.Trim
                    If (strnaturetype = "ADDITION") Then
                        .Col = enmGrid.col_Item_NewCumms
                        .Text = Format(dblcumms_AdstValue + dblStkBeforeAdjustment, "#.0000")
                    End If
                    If (strnaturetype = "SUBTRACTION") Then
                        .Col = enmGrid.col_Item_NewCumms
                        .Text = Format(dblStkBeforeAdjustment - dblcumms_AdstValue, "#.0000")
                    End If
                End With
        End Select
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub
    Private Sub SpStckAdjst_DblClick(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_DblClickEvent) Handles SpStckAdjst.DblClick
        With SpStckAdjst
            If (e.col = 0 And e.row > 0) And (CmdStckAdjustBttn.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD) Then
                .Row = e.row
                .Action = FPSpreadADO.ActionConstants.ActionDeleteRow
                .MaxRows = .MaxRows - 1
            End If
        End With
    End Sub
    Private Sub SpStckAdjst_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SpStckAdjst.Enter
        On Error GoTo ErrHandler
        Select Case CmdStckAdjustBttn.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                ToolTip1.SetToolTip(SpStckAdjst, "Press Ctrl+N for New Row")
        End Select
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SpStckAdjst_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SpStckAdjst.Leave
        On Error GoTo ErrHandler
        ToolTip1.SetToolTip(SpStckAdjst, "")
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtDocCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDocCode.TextChanged
        On Error GoTo ErrHandler
        Select Case CmdStckAdjustBttn.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                If (txtDocCode.Text.Trim.Length = 0) Then
                    Call RefreshControls()
                End If
        End Select
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtDocCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDocCode.Enter
        On Error GoTo ErrHandler
        With Me.txtDocCode
            .SelectionStart = 0
            .SelectionLength = Len(txtDocCode.Text)
        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtDocCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDocCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then
            Call txtDocCode_Validating(txtDocCode, New System.ComponentModel.CancelEventArgs(False))
        Else
            KeyAscii = KeyAscii
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
    Private Sub txtDocCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtDocCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            Call cmdDocCode_Click(cmdDocCode, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtDocCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtDocCode.Validating
        On Error GoTo ErrHandler
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim strsql As String
        Dim oRS As ADODB.Recordset
        If (txtDocCode.Text.Trim.Length = 0) Then Exit Sub
        strsql = "Select Doc_No,customer_code From ASN_CummsAdjst_hdr" & _
            " WHERE UNIT_CODE='" & gstrUNITID & "' AND Doc_Type = 9997" & _
            " AND DOC_NO = '" & txtDocCode.Text.Trim & "' AND AUTHORIZED_CODE IS NULL"
        oRS = mP_Connection.Execute(strsql)

        SpStckAdjst.MaxRows = 0
        txtDocCode.Text = String.Empty
        txtCustomerCode.Text = String.Empty


        If Not (oRS.BOF And oRS.EOF) Then
            'lblDisplayDate.Text = VB6.Format(oRS.Fields("Doc_date").Value, gstrDateFormat)


            txtDocCode.Text = oRS.Fields("Doc_No").Value
            txtCustomerCode.Text = oRS.Fields("customer_code").Value

            Call FillDataInSpread()

            With CmdStckAdjustBttn
                .Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
                .Focus()
            End With
        Else
            Cancel = True
            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_EXCLAMATION) ' flashes a message showing that record doesn't exist.
            Call RefreshControls()
            txtDocCode.Text = String.Empty
            txtCustomerCode.Text = String.Empty

        End If
        oRS.Close()
        oRS = Nothing
        GoTo EventExitSub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtLocationCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLocationCode.TextChanged
        On Error GoTo ErrHandler
        Select Case CmdStckAdjustBttn.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                If (txtLocationCode.Text.Trim.Length = 0) Then
                    Call RefreshControls()
                    lbllocationdesc.Text = String.Empty
                    With txtDocCode
                        .Clear()
                        .Enabled = False
                        .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    End With
                    cmdDocCode.Enabled = False
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                Call RefreshControls()
                lbllocationdesc.Text = String.Empty
        End Select
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub RefreshControls()
        On Error GoTo ErrHandler
        With SpStckAdjst
            .MaxRows = 0
        End With
        lblDisplayDate.Text = VB6.Format(mServerDate, gstrDateFormat)
        lblCustomerdesc.Text = String.Empty
        With CmdStckAdjustBttn
            .Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtLocationCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLocationCode.Enter
        On Error GoTo ErrHandler
        With Me.txtLocationCode
            .SelectionStart = 0
            .SelectionLength = Len(txtLocationCode.Text)
        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub FillDataInSpread()
        Dim strsql As String
        Dim StrItemCode As String
        Dim strUOM As String
        Dim oRS As ADODB.Recordset
        Dim dblStkThatDay As Double
        Dim dblAdjustedStk As Double
        Dim StrPartCode As String
        Dim StrPlantCode As String
        Dim StrPartDescription As String
        Dim StrNature As String
        Dim StrFirstPartCode As String
        Dim StrFirstitemCode As String

        strsql = "SELECT * FROM ASN_CUMMSADJST_HDR HD INNER JOIN ASN_CUMMSADJST_DTL DT ON HD.UNIT_CODE=DT.UNIT_CODE " & _
        " AND HD.DOC_TYPE=DT.DOC_TYPE AND HD.DOC_NO=DT.DOC_NO AND " & _
    " HD.UNIT_CODE='" & gstrUNITID & "' AND HD.DOC_TYPE = 9997" & _
    " AND HD.DOC_NO = '" & txtDocCode.Text.Trim & "' AND HD.AUTHORIZED_CODE IS NULL "

        oRS = mP_Connection.Execute(strsql)
        If Not (oRS.EOF And oRS.BOF) Then
            With SpStckAdjst
                While Not oRS.EOF
                    Call AddBlankRowInGrid()
                    .Row = .MaxRows
                    .Col = enmGrid.col_Part_Code
                    .Text = oRS.Fields("CUST_ITEM_CODE").Value
                    StrPartCode = .Text.Trim
                    .Col = enmGrid.col_Item_UOM
                    .Text = oRS.Fields("UOM").Value
                    .Col = enmGrid.col_Item_Code
                    .Text = oRS.Fields("ITEM_CODE").Value
                    StrItemCode = .Text.Trim
                    .Col = enmGrid.col_Item_CummsBeforeAdjustment
                    .Text = oRS.Fields("CURRENT_CUMMS").Value
                    dblStkThatDay = .Text
                    .Col = enmGrid.col_Item_CummsDiff
                    .Text = oRS.Fields("Adjusted_CUMMS").Value
                    dblAdjustedStk = .Text
                    .Col = enmGrid.col_Item_NewCumms
                    .Text = oRS.Fields("AfterAdjst_Cumms").Value
                    .Col = enmGrid.col_Item_Nature
                    .Text = oRS.Fields("nature").Value
                    StrNature = .Text.Trim
                    .Col = enmGrid.col_Invoice
                    .Text = oRS.Fields("Invoice").Value
                    .Col = enmGrid.col_CDR_Reference
                    .Text = oRS.Fields("CDR_Reference").Value

                    oRS.MoveNext()
                End While
                .Col = enmGrid.col_Item_Code
                .Row = 1
                StrFirstitemCode = .Text


                .Col = enmGrid.col_Part_Code
                .Row = 1
                StrFirstPartCode = .Text
                If (StrFirstPartCode.Length > 0 And StrFirstitemCode.Length > 0) Then
                    txtcustpartdesc.Text = CStr(Find_Value("SELECT DRG_DESC FROM CUSTITEM_MST  WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & txtCustomerCode.Text.Trim & "' AND ACTIVE=1 AND ITEM_CODE='" & StrFirstitemCode & "' AND CUST_DRGNO='" & StrFirstPartCode & "'"))
                End If

            End With
        End If

        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function CheckPreviousRecords() As Boolean
        On Error GoTo ErrHandler
        Dim intCheckRow As Short
        Dim varCheckPrevItemCode As Object
        Dim varCheckNextItemCode As Object
        CheckPreviousRecords = False
        With Me.SpStckAdjst
            For intCheckRow = 1 To .MaxRows
                varCheckPrevItemCode = Nothing
                Call .GetText(1, intCheckRow, varCheckPrevItemCode)
                varCheckNextItemCode = Nothing
                Call .GetText(1, .ActiveRow, varCheckNextItemCode)
                If varCheckNextItemCode <> "" Then
                    If Not (intCheckRow = .ActiveRow) Then
                        If Trim(varCheckNextItemCode) = Trim(varCheckPrevItemCode) Then
                            CheckPreviousRecords = True
                            Call .SetText(1, .ActiveRow, "")
                            Call MsgBox("Item Code [" & varCheckNextItemCode & "] Is Already Entered", MsgBoxStyle.Information, ResolveResString(100))
                            .Row = .ActiveRow
                            .Col = 1
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            Exit Function
                        Else
                            CheckPreviousRecords = False
                        End If
                    End If
                End If
            Next intCheckRow
        End With
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function ValidatebeforeSave() As Boolean
        On Error GoTo ErrHandler
        Dim intLoopCounter As Short
        Dim lstrControls As String
        Dim lNo As Integer
        Dim lctrFocus As System.Windows.Forms.Control
        Dim intmLoopCounter As Short
        Dim SumItemQty As Double
        ValidatebeforeSave = False
        lNo = 1
        lstrControls = ResolveResString(10059)
        lctrFocus = Nothing
        With SpStckAdjst
            Call DeleteBlankRows()
            If (txtLocationCode.Text.Trim.Length = 0) Then
                lstrControls = lstrControls & vbCrLf & lNo & ". Location Code."
                lNo = lNo + 1
                If lctrFocus Is Nothing Then
                    lctrFocus = Me.txtLocationCode
                End If
                ValidatebeforeSave = True
            End If

            If (txtCustomerCode.Text.Trim.Length = 0) Then
                lstrControls = lstrControls & vbCrLf & lNo & ". Customer Code."
                lNo = lNo + 1
                If lctrFocus Is Nothing Then
                    lctrFocus = Me.txtCustomerCode
                End If
                ValidatebeforeSave = True
            End If

            

            If .MaxRows <= 0 Then
                lstrControls = lstrControls & vbCrLf & lNo & ". At least one Part code must be there."
                lNo = lNo + 1
                If lctrFocus Is Nothing Then
                    lctrFocus = Me.txtCustomerCode
                End If
                ValidatebeforeSave = True
            End If




        End With
        If (ValidatebeforeSave = True) Then
            MsgBox(lstrControls, MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
            lctrFocus.Focus()
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function updateCummsAdjustTable(ByVal strtype As String) As Boolean
        On Error GoTo ErrHandler
        Dim strInsertStck_Hdr As String = ""
        Dim strInsertStck_Dtl As String = ""
        Dim strStocUpdate As String = ""
        Dim strDocumentNumber As String = ""
        Dim intLoopCounter As Short
        Dim StrPartCode As String = ""
        Dim dblCums_Qty As Double
        Dim dblAfteradusted_cumms As Double
        Dim dblAdusted_cumms As Double
        Dim strSQL As String = ""
        Dim STRINVOICETYPE As String
        Dim StrNature As String
        Dim StrUOM As String
        Dim strInvoice As String
        Dim strPartDesc As String
        Dim stritemcode As String

        updateCummsAdjustTable = False
        STRINVOICETYPE = "INV"
        ResetDatabaseConnection()
        mP_Connection.BeginTrans()
        If UCase(strtype) = "AUTHORIZE" Then
            strSQL = "UPDATE ASN_CUMMSADJST_HDR SET AUTH_STATUS=1, AUTHORIZED_CODE='" & Trim(mP_User) & "',AUTH_DATE=CONVERT(DATETIME, CONVERT(VARCHAR(11), GETDATE(), 106), 106), " & _
            " AUTH_TIME= SUBSTRING(CONVERT(VARCHAR(20),GETDATE()),13,LEN(GETDATE())),UPD_DT=GETDATE(),UPD_USERID='" & Trim(mP_User) & "' WHERE DOC_TYPE=9997 AND DOC_NO='" & Me.txtDocCode.Text.Trim & "' AND UNIT_CODE='" & gstrUNITID & "' "
        Else
            strSQL = "UPDATE ASN_CUMMSADJST_HDR SET AUTH_STATUS=2, AUTHORIZED_CODE='" & Trim(mP_User) & "',AUTH_DATE=CONVERT(DATETIME, CONVERT(VARCHAR(11), GETDATE(), 106), 106), " & _
            " AUTH_TIME= SUBSTRING(CONVERT(VARCHAR(20),GETDATE()),13,LEN(GETDATE())),UPD_DT=GETDATE(),UPD_USERID='" & Trim(mP_User) & "' WHERE DOC_TYPE=9997 AND DOC_NO='" & Me.txtDocCode.Text.Trim & "' AND UNIT_CODE='" & gstrUNITID & "' "
        End If
        mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.CommitTrans()

        updateCummsAdjustTable = True
        Exit Function
ErrHandler:
        mP_Connection.RollbackTrans()
        updateCummsAdjustTable = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub DeleteBlankRows()
        On Error GoTo ErrHandler
        Dim lngRowCount As Integer
        With Me.SpStckAdjst
StartHere:
            For lngRowCount = 1 To .MaxRows
                .Row = lngRowCount
                .Col = enmGrid.col_Part_Code
                If (.Text.Trim.Length = 0) Then
                    .Action = FPSpreadADO.ActionConstants.ActionDeleteRow
                    .MaxRows = .MaxRows - 1
                    GoTo StartHere
                End If
            Next lngRowCount
        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    
    
    Private Sub SpStckAdjst_EditChange(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles SpStckAdjst.EditChange
        On Error GoTo ErrHandler
        Dim intNoOfBatch As Integer
        Dim dblStkBeforeAdjustment As Double
        Dim dblcumms_AdstValue As Double
        Dim strnaturetype As String
        Select Case e.col
            Case enmGrid.col_Item_CummsDiff
                With SpStckAdjst
                    .Row = e.row
                    .Col = enmGrid.col_Item_CummsBeforeAdjustment
                    dblStkBeforeAdjustment = Val(.Text)
                    .Col = enmGrid.col_Item_CummsDiff
                    dblcumms_AdstValue = Val(.Text)
                    .Col = enmGrid.col_Item_Nature
                    strnaturetype = .Text.Trim
                    If (strnaturetype = "ADDITION") Then
                        .Col = enmGrid.col_Item_NewCumms
                        .Text = Format(dblcumms_AdstValue + dblStkBeforeAdjustment, "#.0000")
                    End If
                    If (strnaturetype = "SUBTRACTION") Then
                        .Col = enmGrid.col_Item_NewCumms
                        .Text = Format(dblStkBeforeAdjustment - dblcumms_AdstValue, "#.0000")
                    End If
                End With

        End Select
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SpStckAdjst_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles SpStckAdjst.KeyPressEvent
        On Error GoTo ErrHandler
        Select Case CmdStckAdjustBttn.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                Select Case e.keyAscii
                    Case 34, 39, 96
                        e.keyAscii = 0
                End Select
        End Select
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SpStckAdjst_KeyUpEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyUpEvent) Handles SpStckAdjst.KeyUpEvent
        On Error GoTo ErrHandler
        Select Case CmdStckAdjustBttn.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If ((e.keyCode = Keys.N) And (e.shift = ShiftConstants.CtrlMask)) Then
                    With SpStckAdjst
                        .Row = .MaxRows
                        .Col = enmGrid.col_Part_Code
                        If (.Text.Trim.Length > 0) Then
                            Call AddBlankRowInGrid()
                        End If
                    End With
                ElseIf ((e.keyCode = Keys.F1) And (e.shift = 0)) Then
                    With SpStckAdjst
                        Select Case .ActiveCol
                            Case enmGrid.col_Part_Code
                                Dim intActiveRow As Integer
                                intActiveRow = SpStckAdjst.ActiveRow

                        End Select
                    End With
                End If
        End Select
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SpStckAdjst_LeaveCell(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles SpStckAdjst.LeaveCell
        On Error GoTo ErrHandler
        Dim blnStatus As Boolean
        Dim strSql As String
        Dim strItemcode As String
        Dim strPartcode As String

        If e.newCol = -1 Then
            Exit Sub
        End If
        Select Case CmdStckAdjustBttn.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                With SpStckAdjst
                    Select Case e.col
                        Case enmGrid.col_Part_Code
                            .Row = .ActiveRow
                            .Col = enmGrid.col_Part_Code
                            strPartcode = .Text.Trim
                            .Col = enmGrid.col_Item_Code
                            strItemcode = .Text.Trim
                            If (strPartcode.Length > 0) Then
                                'txtcustpartdesc.text = Find_Value("SELECT DRG_DESC FROM CUSTITEM_MST  WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustomerCode.Text.Trim & "' AND ITEM_CODE='" & strItemcode & "' AND CUST_DRGNO ='" & strPartcode & "' AND ACTIVE=1"))
                                txtcustpartdesc.Text = CStr(Find_Value("SELECT DRG_DESC FROM CUSTITEM_MST  WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & txtCustomerCode.Text.Trim & "' AND ACTIVE=1 AND ITEM_CODE='" & strItemcode & "' AND CUST_DRGNO='" & strPartcode & "'"))
                            End If
                    End Select
                End With
        End Select
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SpStckAdjst_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles SpStckAdjst.Validating
        DeleteBlankRows()
    End Sub
    
    Private Sub SetGridHdrs()
        On Error GoTo ErrHandler
        With SpStckAdjst
            .MaxRows = 0
            .MaxCols = enmGrid.col_CDR_Reference
            .Enabled = True
            .Appearance = FPSpreadADO.AppearanceConstants.Appearance3DWithBorder
            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            .EditEnterAction = FPSpreadADO.EditEnterActionConstants.EditEnterActionNext
            .ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsBoth
            .ProcessTab = True
            .UserResizeCol = FPSpreadADO.UserResizeConstants2.UserResizeOn
            .UserResizeRow = FPSpreadADO.UserResizeConstants2.UserResizeOff
            .ColsFrozen = 1
            .set_RowHeight(0, 800)
            .SetText(enmGrid.col_Part_Code, 0, "Part Code")
            .set_ColWidth(enmGrid.col_Part_Code, 1800)

            .SetText(enmGrid.col_Item_Code, 0, "Item Code")
            .set_ColWidth(enmGrid.col_Item_Code, 1200)

            .SetText(enmGrid.col_Item_UOM, 0, "UOM")
            .set_ColWidth(enmGrid.col_Item_UOM, 500)

            .SetText(enmGrid.col_Item_CummsBeforeAdjustment, 0, "Current Cumm. Qty")
            .set_ColWidth(enmGrid.col_Item_CummsBeforeAdjustment, 800)

            .SetText(enmGrid.col_Item_CummsDiff, 0, "Adjust Cumms ")
            .set_ColWidth(enmGrid.col_Item_CummsDiff, 800)

            .SetText(enmGrid.col_Item_Nature, 0, "Nature ")
            .set_ColWidth(enmGrid.col_Item_Nature, 1000)

            .SetText(enmGrid.col_Item_NewCumms, 0, "New Cumms After Adjustment")
            .set_ColWidth(enmGrid.col_Item_NewCumms, 1000)

            .SetText(enmGrid.col_Invoice, 0, "Invoice No ")
            .set_ColWidth(enmGrid.col_Invoice, 1000)

            .SetText(enmGrid.col_CDR_Reference, 0, "CDR Reference ")
            .set_ColWidth(enmGrid.col_CDR_Reference, 2500)

        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub AddBlankRowInGrid()
        On Error GoTo ErrHandler
        With SpStckAdjst
            If (.MaxRows > 0) Then
                .Row = .MaxRows
                .Col = enmGrid.col_Part_Code
                If (.Text.Trim.Length = 0) Then
                    Exit Sub
                End If
            End If
            .Enabled = True
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .set_RowHeight(.Row, 300)

            .Col = enmGrid.col_Part_Code
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft

            .Col = enmGrid.col_Item_Code
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft


            .Col = enmGrid.col_Item_UOM
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter

            .Col = enmGrid.col_Item_CummsBeforeAdjustment
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .Lock = True

            .Col = enmGrid.col_Item_Nature
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            '.TypeComboBoxList = "ADDITION" & Chr(9) & "SUBTRACTION"
            '.TypeComboBoxCurSel = 0

            .Col = enmGrid.col_Item_CummsDiff
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
            .TypeIntegerMin = 0
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight

            .Col = enmGrid.col_Item_NewCumms
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .Lock = True


            .Col = enmGrid.col_Invoice
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
            .TypeIntegerMin = 0
            .TypeIntegerMax = 99999999
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight

            .Col = enmGrid.col_CDR_Reference
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft


            .Row = 1
            .Row2 = .MaxRows
            .Col = 1
            .Col2 = .MaxCols
            .BlockMode = True
            .Lock = True
            .BlockMode = False

        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtCustomerCode_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
        On Error GoTo ErrHandler
        Dim strLocCondition As String
        Dim Cancel As Boolean = e.Cancel
        Dim strSql As String

        Select Case CmdStckAdjustBttn.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If (txtCustomerCode.Text.Trim.Length > 0) Then
                    strSql = "SELECT Cust_name FROM CUSTOMER_MST " & _
                    " WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustomerCode.Text.Trim & "' AND ALLOWASNTEXTGENERATION=1 "
                    lblCustomerdesc.Text = GetQryOutput(strSql)
                    If (lblCustomerdesc.Text.Trim.Length > 0) Then

                    Else
                        lblCustomerdesc.Text = String.Empty
                        Call ConfirmWindow(10010)
                        txtCustomerCode.Clear()
                        Cancel = True
                    End If
                End If
        End Select
        GoTo EventExitSub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        e.Cancel = Cancel
    End Sub

    Private Sub TxtCDRreferences_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Dim Cancel As Boolean = e.Cancel
        On Error GoTo ErrHandler
        If Me.SpStckAdjst.MaxRows = 0 Then
            Call AddBlankRowInGrid()
        End If
        GoTo EventExitSub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        e.Cancel = Cancel
    End Sub
    Public Function Find_Value(ByRef strField As String) As String
        '----------------------------------------------------------------------------
        'Author         :   Arshad Ali
        'Argument       :   Sql query string as strField
        'Return Value   :   selected table field value as String
        'Function       :   Return a field value from a table
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim Rs As New ADODB.Recordset
        Rs = New ADODB.Recordset
        Rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        Rs.Open(strField, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
        If Rs.RecordCount > 0 Then
            If IsDBNull(Rs.Fields(0).Value) = False Then
                Find_Value = Rs.Fields(0).Value
            Else
                Find_Value = ""
            End If
        Else
            Find_Value = ""
        End If
        Rs.Close()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Sub SpStckAdjst_ClickEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles SpStckAdjst.ClickEvent

        On Error GoTo ErrHandler

        Dim blnStatus As Boolean
        Dim strSql As String
        Dim strItemcode As String
        Dim strPartcode As String

        Select Case CmdStckAdjustBttn.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                With SpStckAdjst
                    Select Case e.col
                        Case enmGrid.col_Part_Code
                            .Row = .ActiveRow
                            .Col = enmGrid.col_Part_Code
                            strPartcode = .Text.Trim
                            .Col = enmGrid.col_Item_Code
                            strItemcode = .Text.Trim
                            If (strPartcode.Length > 0) Then
                                txtcustpartdesc.Text = CStr(Find_Value("SELECT DRG_DESC FROM CUSTITEM_MST  WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & txtCustomerCode.Text.Trim & "' AND ACTIVE=1 AND ITEM_CODE='" & strItemcode & "' AND CUST_DRGNO='" & strPartcode & "'"))
                            End If
                    End Select
                End With
        End Select
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdauthorize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdauthorize.Click
        On Error GoTo ErrHandler

        If (ValidatebeforeSave() = False) Then
            If ConfirmWindow(10177, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                If updateCummsAdjustTable("AUTHORIZE") Then
                    MsgBox("ASN Cumms No " & txtDocCode.Text & " has been Authorized ", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                    Call RefreshControls()
                    SpStckAdjst.Enabled = True
                    With Me.txtDocCode
                        .Enabled = True
                        .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        .Focus()
                    End With
                    cmdDocCode.Enabled = True
                    txtcustpartdesc.Text = String.Empty
                    txtDocCode.Text = String.Empty
                    txtCustomerCode.Text = String.Empty
                End If
            End If
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub cmdReject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReject.Click
        On Error GoTo ErrHandler

        If (ValidatebeforeSave() = False) Then
            If ConfirmWindow(10177, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                If updateCummsAdjustTable("REJECT") Then
                    MsgBox("ASN Cumms No  " & txtDocCode.Text & " has been Rejected ", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                    SpStckAdjst.Enabled = True
                    With Me.txtDocCode
                        .Enabled = True
                        .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        .Focus()
                    End With
                    cmdDocCode.Enabled = True
                    txtcustpartdesc.Text = String.Empty
                    txtDocCode.Text = String.Empty
                    txtCustomerCode.Text = String.Empty
                End If
            End If
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub CmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdClose.Click
        On Error GoTo ErrHandler
        Me.Close()
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
End Class