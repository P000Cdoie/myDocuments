Option Strict Off
Option Explicit On
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility.VB6
Imports System.Data.SqlClient

Friend Class frmMKTTRN0079
    Inherits System.Windows.Forms.Form
    '***************************************************************************************
    'COPYRIGHT       : MIND LTD.
    'MODULE          : FRMMKTTRNP0079 - CUMULATIVE QUANTITY ADJUSTMENT 
    'AUTHOR          : PRASHANT RAJPAL
    'CREATION DATE   : 10-JULY 2013- 20 JUL 2013
    ' ISSUE ID       : 10416813 
    'PURPOSE          : CUMULATIVE ASN FUNCTIONLITY -TRANSACTION FORM FOR CREATING CDR NO
    '***************************************************************************************

    Private mlngFormTag As Integer
    Private mServerDate As String
    Dim mstrUpdateDocumentNoSQL As String
    Private Enum enmGrid
        col_Part_Code = 1
        col_Item_Hlp = 2
        col_Part_Code_Description = 3
        col_Item_Code = 4
        col_Item_UOM = 5
        col_Item_CummsBeforeAdjustment = 6
        col_Item_CummsDiff = 7
        col_Item_Nature = 8
        col_Item_NewCumms = 9
        col_Invoice = 10
        col_CDR_Reference = 11
    End Enum

    Private Sub cmdDocCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDocCode.Click
        Dim StrDocHelp As String
        On Error GoTo ErrHandler
        StrDocHelp = ShowList(0, Len(txtDocCode.Text), txtDocCode.Text, "Doc_No", "" & DateColumnNameInShowList("Trans_date") & " as Trans_Date", "ASN_CUMMSADJST_HDR", " ", "ASN ADUSTED SERIES ")
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

    Private Sub FRMMKTTRN0079_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mlngFormTag
        frmModules.NodeFontBold(Me.Tag) = True
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub FRMMKTTRN0079_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub FRMMKTTRN0079_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If (KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0) Then Call ctlStckAdjstHeader_ClickEvent(ctlStckAdjstHeader, New System.EventArgs())
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub FRMMKTTRN0079_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                If (Me.CmdASNAdjustBttn.Mode) <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call Me.CmdASNAdjustBttn.Revert() : gblnCancelUnload = False : gblnFormAddEdit = False
                        Call EnableControls(False, Me, True)
                        SpASNAdjst.Enabled = True

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

                        With Me.txtinvoicetype
                            .Enabled = False
                            .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
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

    Private Sub FRMMKTTRN0079_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler

        mlngFormTag = mdifrmMain.AddFormNameToWindowList(Me.ctlStckAdjstHeader.Tag)
        Call FitToClient(Me, fraMain, ctlStckAdjstHeader, CmdASNAdjustBttn, 400)
        CmdASNAdjustBttn.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(fraMain.Left) + (VB6.PixelsToTwipsX(fraMain.Width) / 4))
        Call EnableControls(False, Me, True)
        mServerDate = getDateForDB(GetServerDate())
        Call InitializeControls()
        txtLocationCode.Text = gstrUNITID
        lbllocationdesc.Text = GetQryOutput("SELECT Description FROM Location_Mst WHERE UNIT_CODE='" & gstrUNITID & "' AND LOCATION_CODE='" & gstrUNITID & "'")

        With txtDocCode
            .Enabled = True
            .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            .Focus()
        End With
        With Me.txtinvoicetype
            .Enabled = False
            .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            .Focus()
        End With
        cmdDocCode.Enabled = True
        cmdInvoicetype.Enabled = False
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
        With CmdASNAdjustBttn
            .Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        End With
        Call CmdASNAdjustBttn.ShowButtons(True, False, False, True)
        Call SetGridHdrs()
        Exit Sub

ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub FRMMKTTRN0079_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        On Error GoTo ErrHandler

        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
            If CmdASNAdjustBttn.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
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
    Private Sub FRMMKTTRN0079_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler

        Me.Dispose()
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mlngFormTag
        Exit Sub

ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub SpASNAdjst_Change(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ChangeEvent) Handles SpASNAdjst.Change
        On Error GoTo ErrHandler

        Dim intNoOfBatch As Integer
        Dim dblStkBeforeAdjustment As Double
        Dim dblcumms_AdstValue As Double
        Dim strnaturetype As String
        Select Case e.col
            Case enmGrid.col_Item_Nature, enmGrid.col_Part_Code
                With SpASNAdjst
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
    Private Sub SpASNAdjst_DblClick(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_DblClickEvent) Handles SpASNAdjst.DblClick
        With SpASNAdjst
            If (e.col = 0 And e.row > 0) And (CmdASNAdjustBttn.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD) Then
                .Row = e.row
                .Action = FPSpreadADO.ActionConstants.ActionDeleteRow
                .MaxRows = .MaxRows - 1
            End If
        End With
    End Sub
    Private Sub SpASNAdjst_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SpASNAdjst.Enter
        On Error GoTo ErrHandler
        Select Case CmdASNAdjustBttn.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                ToolTip1.SetToolTip(SpASNAdjst, "Press Ctrl+N for New Row")
        End Select
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SpASNAdjst_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SpASNAdjst.Leave
        On Error GoTo ErrHandler
        ToolTip1.SetToolTip(SpASNAdjst, "")
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtDocCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDocCode.TextChanged
        On Error GoTo ErrHandler
        Select Case CmdASNAdjustBttn.Mode
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
        If (txtDocCode.Text.Trim.Length = 0) Then
            SpASNAdjst.MaxRows = 0
            txtDocCode.Text = String.Empty
            txtCustomerCode.Text = String.Empty
            txtcustpartdesc.Text = String.Empty
            txtinvoicetype.Text = String.Empty
            cmdcustcodehelp.Enabled = False
            cmdInvoicetype.Enabled = False
            txtcustpartdesc.Text = String.Empty
            Exit Sub
        End If

        strsql = "Select Doc_No,customer_code From ASN_CummsAdjst_hdr" & _
            " WHERE UNIT_CODE='" & gstrUNITID & "' AND Doc_Type = 9997" & _
            " AND Doc_No = '" & txtDocCode.Text.Trim & "'"
        oRS = mP_Connection.Execute(strsql)

        SpASNAdjst.MaxRows = 0
        txtDocCode.Text = String.Empty
        txtCustomerCode.Text = String.Empty
        txtcustpartdesc.Text = String.Empty
        txtinvoicetype.Text = String.Empty
        If Not (oRS.BOF And oRS.EOF) Then

            txtDocCode.Text = oRS.Fields("Doc_No").Value
            txtCustomerCode.Text = oRS.Fields("customer_code").Value

            Call FillDataInSpread()

            With CmdASNAdjustBttn
                .Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
                .Focus()
            End With
        Else
            Cancel = True
            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_EXCLAMATION) ' flashes a message showing that record doesn't exist.
            Call RefreshControls()
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

        Select Case CmdASNAdjustBttn.Mode
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

        With SpASNAdjst
            .MaxRows = 0
        End With
        lblDisplayDate.Text = VB6.Format(mServerDate, gstrDateFormat)
        txtCustomerCode.Text = String.Empty
        lblCustomerdesc.Text = String.Empty
        txtcustpartdesc.Text = String.Empty

        With CmdASNAdjustBttn
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

        Select Case CmdASNAdjustBttn.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                strsql = "Select * From ASN_CummsAdjst_hdr HD INNER JOIN ASN_CummsAdjst_DTL DT ON HD.UNIT_CODE=DT.UNIT_CODE " & _
                " AND HD.DOC_TYPE=DT.DOC_TYPE AND HD.DOC_NO=DT.DOC_NO AND " & _
            " HD.UNIT_CODE='" & gstrUNITID & "' AND HD.Doc_Type = 9997" & _
            " AND HD.Doc_No = '" & txtDocCode.Text.Trim & "'"

                oRS = mP_Connection.Execute(strsql)
                If Not (oRS.EOF And oRS.BOF) Then
                    With SpASNAdjst
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
                            .Col = enmGrid.col_Invoice
                            .Text = oRS.Fields("Invoice").Value
                            .Col = enmGrid.col_CDR_Reference
                            .Text = oRS.Fields("CDR_REFERENCE").Value
                            oRS.MoveNext()
                        End While
                        
                        .Col = enmGrid.col_Item_Code
                        .Row = 1
                        StrFirstitemCode = .Text


                        .Col = enmGrid.col_Part_Code
                        .Row = 1
                        StrFirstPartCode = .Text
                        If (StrFirstPartCode.Length > 0) Then
                            txtcustpartdesc.Text = CStr(Find_Value("SELECT DRG_DESC FROM CUSTITEM_MST  WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & txtCustomerCode.Text.Trim & "' AND ACTIVE=1 AND ITEM_CODE='" & StrFirstitemCode & "' AND CUST_DRGNO='" & StrFirstPartCode & "'"))
                        End If
                    End With
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                With SpASNAdjst

                    .Row = .ActiveRow
                    .Col = enmGrid.col_Part_Code
                    StrPartCode = .Text.Trim

                    .Col = enmGrid.col_Part_Code_Description
                    StrPartDescription = .Text.Trim

                    .Col = enmGrid.col_Item_UOM
                    strUOM = GetItemUOM(StrItemCode)
                    .Text = strUOM

                    StrPlantCode = CStr(Find_Value("SELECT PLANT_CODE FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustomerCode.Text.Trim & "'"))
                    .Col = enmGrid.col_Item_CummsBeforeAdjustment
                    .Text = CInt(Find_Value("SELECT DBO.UDF_GET_CUMMULATIVEQTY_ADJUSTMENT('" & gstrUNITID & "','" & StrPlantCode & "','" & StrPartCode & "','" & Me.txtCustomerCode.Text.Trim & "','" & Me.txtinvoicetype.Text & "')"))


                    .Col = enmGrid.col_Item_NewCumms
                    .Text = CInt(Find_Value("SELECT DBO.UDF_GET_CUMMULATIVEQTY_ADJUSTMENT('" & gstrUNITID & "','" & StrPlantCode & "','" & StrPartCode & "','" & Me.txtCustomerCode.Text.Trim & "','" & Me.txtinvoicetype.Text & "')"))

                    .Col = enmGrid.col_Invoice
                    .Text = ""
                    .Col = enmGrid.col_CDR_Reference
                    .Text = ""
                    txtcustpartdesc.Text = String.Empty
                    .Col = enmGrid.col_Item_CummsDiff
                    .Value=0 
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    .Focus()
                End With
        End Select
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
        With Me.SpASNAdjst
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
        Dim strquery As String
        Dim stritemcode As String
        Dim strPartcode As String
        Dim CDRno As String

        ValidatebeforeSave = False
        lNo = 1
        lstrControls = ResolveResString(10059)
        lctrFocus = Nothing
        With SpASNAdjst
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

            If (txtinvoicetype.Text.Trim.Length = 0) Then
                lstrControls = lstrControls & vbCrLf & lNo & ". Invoice Type ."
                lNo = lNo + 1
                If lctrFocus Is Nothing Then
                    lctrFocus = Me.txtLocationCode
                End If
                ValidatebeforeSave = True
            End If


            If .MaxRows <= 0 Then
                lstrControls = lstrControls & vbCrLf & lNo & ". Define at least one Part code ."
                lNo = lNo + 1
                If lctrFocus Is Nothing Then
                    lctrFocus = txtCustomerCode
                End If
                ValidatebeforeSave = True
            End If

            For intLoopCounter = 1 To .MaxRows
                .Row = intLoopCounter
                .Col = enmGrid.col_Invoice
                If Val(.Text) = CDbl("0") Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Invoice no can't be ZERO Or Blank ."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.SpASNAdjst
                        .Col = enmGrid.col_Invoice
                        .Row = intLoopCounter
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        .Focus()
                    End If
                    ValidatebeforeSave = True
                    Exit For
                End If
            Next

            For intLoopCounter = 1 To .MaxRows
                .Row = intLoopCounter
                .Col = enmGrid.col_Item_NewCumms
                If Val(.Text) < 0 Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". New Cumms After Adustment cannot be Negative."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.SpASNAdjst
                        .Row = intLoopCounter
                        .Col = enmGrid.col_Item_CummsDiff
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        .Focus()
                    End If
                    ValidatebeforeSave = True
                    Exit For
                End If
            Next

            For intLoopCounter = 1 To .MaxRows
                .Row = intLoopCounter
                .Col = enmGrid.col_CDR_Reference
                If Len(.Text.Trim) <= 0 Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". CDR Reference cant be Blank ."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.SpASNAdjst
                        .Col = enmGrid.col_CDR_Reference
                        .Row = intLoopCounter
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        .Focus()
                    End If
                    ValidatebeforeSave = True
                    Exit For
                End If
            Next

            For intLoopCounter = 1 To .MaxRows
                .Row = intLoopCounter
                .Col = enmGrid.col_Item_CummsDiff
                If Val(.Text) = CDbl("0") Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Adusted Cumms can't be ZERO."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.SpASNAdjst
                        .Col = enmGrid.col_Item_CummsDiff
                        .Row = intLoopCounter
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        .Focus()
                    End If
                    ValidatebeforeSave = True
                    Exit For
                End If
            Next

            
            For intLoopCounter = 1 To .MaxRows

                .Row = intLoopCounter
                .Col = enmGrid.col_Item_Code
                stritemcode = .Text

                .Row = intLoopCounter
                .Col = enmGrid.col_Part_Code
                strPartcode = .Text

                strquery = "SELECT TOP 1 1 FROM ASN_CUMMSADJST_HDR H INNER JOIN ASN_CUMMSADJST_DTL D " & _
                                             " ON H.UNIT_CODE=D.UNIT_CODE AND H.DOC_NO=D.DOC_NO AND H.AUTHORIZED_CODE IS NULL AND " & _
                                             " H.UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustomerCode.Text.Trim & "' AND ITEM_CODE='" & stritemcode & "' AND CUST_ITEM_CODE='" & strPartcode & "'"
                If DataExist(strquery) = True Then

                    CDRno = CStr(Find_Value("SELECT DISTINCT H.DOC_NO FROM ASN_CUMMSADJST_HDR H INNER JOIN ASN_CUMMSADJST_DTL D " & _
                                             " ON H.UNIT_CODE=D.UNIT_CODE AND H.DOC_NO=D.DOC_NO AND H.AUTHORIZED_CODE IS NULL AND " & _
                                             " H.UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustomerCode.Text.Trim & "' AND ITEM_CODE='" & stritemcode & "' AND CUST_ITEM_CODE='" & strPartcode & "'"))

                    lstrControls = lstrControls & vbCrLf & lNo & ". UNAUTHORIZED CDR (CDR NO : " & CDRno & " ) PENDING FOR ITEM CODE ." & stritemcode
                    lNo = lNo + 1

                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtCustomerCode
                    End If
                    ValidatebeforeSave = True
                End If
            Next

        End With
        If (ValidatebeforeSave = True) Then
            MsgBox(lstrControls, MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
            lctrFocus.Focus()
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function InsertIntoCummsAdjustTable() As Boolean
        On Error GoTo ErrHandler
        Dim strReference As String = ""
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

        InsertIntoCummsAdjustTable = False
        'STRINVOICETYPE = "INV"
        STRINVOICETYPE = txtinvoicetype.Text.Trim
        ResetDatabaseConnection()
        mP_Connection.BeginTrans()
        strDocumentNumber = ""
        strDocumentNumber = Generate_docNo()
        If strDocumentNumber Is Nothing Then
            MsgBox("Document number cannot be generated. Document series not defined", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
            mP_Connection.RollbackTrans()
            txtLocationCode.Focus()
            InsertIntoCummsAdjustTable = False
            Exit Function
        End If
        If strDocumentNumber.Trim.Length = 0 Then
            MsgBox("Document number cannot be generated. Document series not defined", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
            mP_Connection.RollbackTrans()
            txtLocationCode.Focus()
            InsertIntoCummsAdjustTable = False
            Exit Function
        Else
            While Len(strDocumentNumber) < 6
                strDocumentNumber = "0" + strDocumentNumber
            End While
            strDocumentNumber = "CDR" + strDocumentNumber
            txtDocCode.Text = Trim(strDocumentNumber)
        End If
        strSQL = "INSERT INTO ASN_CUMMSADJST_HDR(UNIT_CODE,DOC_TYPE,DOC_NO,INVOICE_TYPE,CUSTOMER_CODE,TRANS_DATE,TRANS_TIME,ENT_DT,ENT_USERID,UPD_DT,UPD_USERID ) " & _
                        " VALUES ('" & gstrUNITID & "',9997 " & ",'" & strDocumentNumber & "','" & STRINVOICETYPE & "','" & Me.txtCustomerCode.Text.Trim & "'," & _
                        " CONVERT(DATETIME, CONVERT(VARCHAR(11), GETDATE(), 106), 106) ,SUBSTRING(CONVERT(VARCHAR(20),GETDATE()),13,LEN(GETDATE()))," & _
                        "GETDATE(),'" & Trim(mP_User) & "', GETDATE(),'" & Trim(mP_User) & "')"
        mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        With SpASNAdjst
            For intLoopCounter = 1 To .MaxRows
                .Row = intLoopCounter
                .Col = enmGrid.col_Part_Code
                StrPartCode = .Text.Trim
                .Col = enmGrid.col_Part_Code_Description
                strPartDesc = .Text.Trim
                .Col = enmGrid.col_Item_Code
                stritemcode = .Text.Trim
                .Col = enmGrid.col_Item_CummsBeforeAdjustment
                dblCums_Qty = Val(.Text)
                .Col = enmGrid.col_Item_CummsDiff
                dblAdusted_cumms = Val(.Text)
                .Col = enmGrid.col_Item_NewCumms
                dblAfteradusted_cumms = Val(.Text)
                .Col = enmGrid.col_Item_Nature
                StrNature = .Text.Trim
                .Col = enmGrid.col_Item_UOM
                StrUOM = .Text.Trim
                .Col = enmGrid.col_Invoice
                strInvoice = .Text.Trim
                .Col = enmGrid.col_CDR_Reference
                strReference = .Text.Trim

                strSQL = "INSERT INTO ASN_CUMMSADJST_DTL(UNIT_CODE,DOC_TYPE,DOC_NO,CUST_ITEM_CODE,CUST_ITEM_DESC,ITEM_CODE ,UOM,CDR_REFERENCE,CURRENT_CUMMS,ADJUSTED_CUMMS,AFTERADJST_CUMMS,NATURE,INVOICE) " & _
                        " VALUES ('" & gstrUNITID & "',9997 " & ",'" & strDocumentNumber & "','" & StrPartCode & "','" & strPartDesc & "','" & stritemcode & "','" & StrUOM & "','" & strReference & "'," & dblCums_Qty & _
                        "," & dblAdusted_cumms & "," & dblAfteradusted_cumms & ",'" & StrNature & "'," & strInvoice & " )"

                mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            Next
        End With


        mP_Connection.CommitTrans()

        InsertIntoCummsAdjustTable = True
        Exit Function
ErrHandler:
        mP_Connection.RollbackTrans()
        InsertIntoCummsAdjustTable = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub DeleteBlankRows()
        On Error GoTo ErrHandler
        Dim lngRowCount As Integer
        With Me.SpASNAdjst
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


    Private Sub CmdStckAdjustBttn_ButtonClick(ByVal Sender As System.Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles CmdASNAdjustBttn.ButtonClick
        On Error GoTo ErrHandler
        Dim REPDOC As ReportDocument
        Dim REPVIEWER As New eMProCrystalReportViewer
        Dim str As String
        Dim strQSNo As String
        Dim datDocument_date As Date
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD

                Call EnableControls(False, Me, True)
                txtLocationCode.Text = gstrUNITID
                lbllocationdesc.Text = GetQryOutput("SELECT Description FROM Location_Mst WHERE UNIT_CODE='" & gstrUNITID & "' AND LOCATION_CODE='" & gstrUNITID & "'")
                Call RefreshControls()

                With txtCustomerCode
                    .Enabled = True
                    .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    .Focus()
                End With
                cmdcustcodehelp.Enabled = True
                cmdInvoicetype.Enabled = True
                With txtinvoicetype
                    .Enabled = True
                    .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                End With

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                Call FRMMKTTRN0079_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
                txtLocationCode.Text = gstrUNITID

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                If (ValidatebeforeSave() = False) Then
                    If InsertIntoCummsAdjustTable() Then
                        CmdASNAdjustBttn.Revert()
                        MsgBox("ASN Cumms has been saved successfully with Document No. " & txtDocCode.Text, MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                        Call EnableControls(False, Me)
                        SpASNAdjst.Enabled = True
                        With Me.txtDocCode
                            .Enabled = True
                            .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            .Focus()
                        End With
                        cmdDocCode.Enabled = True
                        cmdInvoicetype.Enabled = True
                    End If
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
        End Select
        Exit Sub
ErrHandler:
        Me.Cursor = System.Windows.Forms.Cursors.Default
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SpASNAdjst_ButtonClicked(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles SpASNAdjst.ButtonClicked
        Dim stritemhelp As String = ""
        Dim strTempItemNo As String = ""
        Dim StrItemCode As String = ""
        Dim intLoopCounter As Long
        Dim strparthelp As String = ""
        Dim strpartdeschelp As String = ""
        Dim intUOMDecimalPlaces As Integer
        On Error GoTo ErrHandler
        Select Case e.col
            Case enmGrid.col_Item_Hlp
                Select Case CmdASNAdjustBttn.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        With SpASNAdjst
                            If SpASNAdjst.MaxRows > 1 Then
                                Dim INT_I As Integer
                                strTempItemNo = " And A.CUST_DRGNO Not IN ("
                                For INT_I = 1 To SpASNAdjst.MaxRows - 1
                                    .Row = INT_I
                                    .Col = 1
                                    strTempItemNo = strTempItemNo + "'" + .Text.Trim + "',"
                                Next
                                If VB.Right(strTempItemNo, 1) = "," Then
                                    strTempItemNo = Mid(strTempItemNo, 1, Len(strTempItemNo) - 1)
                                End If
                                strTempItemNo = strTempItemNo + ")"
                            End If
                            .Row = e.row
                            .Col = enmGrid.col_Part_Code
                            If Len(txtCustomerCode.Text.Trim) <= 0 Then
                                MsgBox("KINDLY DEFINE CUSTOMER CODE FIRST ", MsgBoxStyle.OkOnly, ResolveResString(100))
                                Exit Sub
                            End If
                            'stritemhelp = ShowList(0, Len(.Text), , "CUST_DRGNO", "CIM.DRG_DESC", "CUSTITEM_MST CIM", " AND UNIT_CODE ='" & gstrUNITID & "' AND active=1 and account_code='" & txtCustomerCode.Text.Trim & "'", , , , , "ITEM_CODE")

                            Dim sqlstring As String
                            Dim strHelp() As String

                            sqlstring = "  SELECT A.CUST_DRGNO,DRG_DESC,ITEM_CODE FROM CUSTITEM_MST  A where unit_code='" & gstrUNITID & "'" & _
                                        " AND ACTIVE=1 AND ACCOUNT_CODE='" & txtCustomerCode.Text.Trim & "' " & strTempItemNo


                            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
                            strHelp = ctlhelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, sqlstring, "LIST OF DRAWING/ITEM CODE", 1)

                            If UBound(strHelp) = -1 Then Exit Sub
                            If strHelp(0) = "0" Then
                                MsgBox("Part Code not defined .", MsgBoxStyle.Information, ResolveResString(100))

                            Else
                                strparthelp = strHelp(0)
                                strpartdeschelp = strHelp(1)
                                stritemhelp = strHelp(2)
                                .Row = e.row
                                .Col = enmGrid.col_Part_Code
                                .Text = strparthelp
                                .Col = enmGrid.col_Part_Code_Description
                                .Text = strpartdeschelp
                                .Col = enmGrid.col_Item_Code
                                .Text = stritemhelp

                                Call FillDataInSpread()


                            End If

                            
                        End With
                End Select

        End Select
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SpASNAdjst_EditChange(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles SpASNAdjst.EditChange
        On Error GoTo ErrHandler
        Dim intNoOfBatch As Integer
        Dim dblStkBeforeAdjustment As Double
        Dim dblcumms_AdstValue As Double
        Dim strnaturetype As String
        Select Case e.col
            Case enmGrid.col_Item_CummsDiff
                With SpASNAdjst
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
    Private Sub SpASNAdjst_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles SpASNAdjst.KeyPressEvent
        On Error GoTo ErrHandler
        Select Case CmdASNAdjustBttn.Mode
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
    Private Sub SpASNAdjst_KeyUpEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyUpEvent) Handles SpASNAdjst.KeyUpEvent
        On Error GoTo ErrHandler
        Select Case CmdASNAdjustBttn.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If ((e.keyCode = Keys.N) And (e.shift = ShiftConstants.CtrlMask)) Then
                    With SpASNAdjst
                        .Row = .MaxRows
                        .Col = enmGrid.col_Part_Code
                        If (.Text.Trim.Length > 0) Then
                            Call AddBlankRowInGrid()
                        End If
                    End With
                ElseIf ((e.keyCode = Keys.F1) And (e.shift = 0)) Then
                    With SpASNAdjst
                        Select Case .ActiveCol
                            Case enmGrid.col_Part_Code, enmGrid.col_Item_Hlp
                                Dim intActiveRow As Integer
                                intActiveRow = SpASNAdjst.ActiveRow
                                Call SpASNAdjst_ButtonClicked(sender, New AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent(2, intActiveRow, 0))
                        End Select
                    End With
                End If
        End Select
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SpASNAdjst_LeaveCell(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles SpASNAdjst.LeaveCell
        On Error GoTo ErrHandler
        Dim blnStatus As Boolean
        Dim strSql As String
        Dim strItemcode As String
        Dim strPartcode As String

        If e.newCol = -1 Then
            Exit Sub
        End If
        Select Case CmdASNAdjustBttn.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                With SpASNAdjst
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
    Private Sub SpASNAdjst_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles SpASNAdjst.Validating
        DeleteBlankRows()
    End Sub
    Private Sub SetGridHdrs()
        On Error GoTo ErrHandler
        With SpASNAdjst
            .MaxRows = 0
            .MaxCols = enmGrid.col_CDR_Reference
            .Enabled = True
            .Appearance = FPSpreadADO.AppearanceConstants.Appearance3DWithBorder
            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            .EditEnterAction = FPSpreadADO.EditEnterActionConstants.EditEnterActionNext
            '.ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsVertical
            .ProcessTab = True
            .UserResizeCol = FPSpreadADO.UserResizeConstants2.UserResizeOn
            .UserResizeRow = FPSpreadADO.UserResizeConstants2.UserResizeOff
            .ColsFrozen = 1
            .set_RowHeight(0, 750)
            .SetText(enmGrid.col_Part_Code, 0, "Part Code")
            .set_ColWidth(enmGrid.col_Part_Code, 1700)

            .SetText(enmGrid.col_Item_Hlp, 0, "Hlp")
            .set_ColWidth(enmGrid.col_Item_Hlp, 350)

            .SetText(enmGrid.col_Part_Code_Description, 0, "Part Description")
            .set_ColWidth(enmGrid.col_Part_Code_Description, 0)

            .SetText(enmGrid.col_Item_Code, 0, "Item Code")
            .set_ColWidth(enmGrid.col_Item_Code, 1200)

            .SetText(enmGrid.col_Item_UOM, 0, "UOM")
            .set_ColWidth(enmGrid.col_Item_UOM, 400)

            .SetText(enmGrid.col_Item_CummsBeforeAdjustment, 0, "Current Cumm. Qty")
            .set_ColWidth(enmGrid.col_Item_CummsBeforeAdjustment, 900)

            .SetText(enmGrid.col_Item_CummsDiff, 0, "Adjust Cumms ")
            .set_ColWidth(enmGrid.col_Item_CummsDiff, 800)

            .SetText(enmGrid.col_Item_Nature, 0, "Nature ")
            .set_ColWidth(enmGrid.col_Item_Nature, 1000)

            .SetText(enmGrid.col_Item_NewCumms, 0, "New Cumms After Adjustment")
            .set_ColWidth(enmGrid.col_Item_NewCumms, 1000)

            .SetText(enmGrid.col_Invoice, 0, "Invoice ")
            .set_ColWidth(enmGrid.col_Invoice, 950)

            .SetText(enmGrid.col_CDR_Reference, 0, "CDR Reference")
            .set_ColWidth(enmGrid.col_CDR_Reference, 2000)

        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub AddBlankRowInGrid()
        On Error GoTo ErrHandler
        With SpASNAdjst
            If (.MaxRows > 0) Then
                .Row = .MaxRows
                .Col = enmGrid.col_Part_Code
                If (.Text.Trim.Length = 0) Then
                    Exit Sub
                End If
            Else
                txtcustpartdesc.Text = String.Empty
                Call SetGridHdrs()
            End If
            If .MaxRows >= 10 Then
                .ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsVertical
                .ProcessTab = True
            End If
            SpASNAdjst.Enabled = True
            .Enabled = True
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .set_RowHeight(.Row, 300)

            .Col = enmGrid.col_Part_Code
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .Lock = True

            .Col = enmGrid.col_Item_Hlp
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
            .TypeButtonPicture = My.Resources.resEmpower.ico111.ToBitmap
            If CmdASNAdjustBttn.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                .ColHidden = True
            Else
                .ColHidden = False
            End If
            .Col = enmGrid.col_Part_Code_Description
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .ColHidden = True

            .Col = enmGrid.col_Item_Code
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .Lock = True

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
            If (CmdASNAdjustBttn.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD) Then
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
                .TypeComboBoxList = "ADDITION" & Chr(9) & "SUBTRACTION"
                .TypeComboBoxCurSel = 0
            Else
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText

            End If


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
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = 0
            .TypeFloatMin = 0
            .TypeFloatMax = 9999999999
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight

            .Col = enmGrid.col_CDR_Reference
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft


            If (CmdASNAdjustBttn.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW) Then
                .Row = 1
                .Row2 = .MaxRows
                .Col = 1
                .Col2 = .MaxCols
                .BlockMode = True
                .Lock = True
                .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                .BlockMode = False
            ElseIf (CmdASNAdjustBttn.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD) Then
                .Row = 1
                .Row = .MaxRows
                .Col = enmGrid.col_Item_NewCumms
                .Col2 = enmGrid.col_Item_NewCumms
                .BlockMode = True
                .Lock = True
                .BlockMode = False


                .Row = 1
                .Row2 = .MaxRows
                .Col = 1
                .Col2 = .MaxCols
                .BlockMode = True
                .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                .BlockMode = False

                .Col = enmGrid.col_Part_Code
                .Row = .MaxRows
                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                .Focus()
            End If

        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdcustcodehelp.Click
        Dim strCustHelp As String
        Dim blnshowbondedinsan As Boolean
        Dim strLocCondition As String

        On Error GoTo ErrHandler
        strLocCondition = "location_code='" & gstrUNITID & "'"

        Select Case CmdASNAdjustBttn.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD

                strCustHelp = ShowList(0, Len(txtLocationCode.Text), , "CUSTOMER_CODE", "CUST_NAME", "CUSTOMER_MST", " AND UNIT_CODE ='" & gstrUNITID & "' AND ALLOWASNTEXTGENERATION =1 ", "CUSTOMER HELP", , , , , )
                If strCustHelp = "-1" Then
                    txtCustomerCode.Focus()
                    Exit Sub

                ElseIf strCustHelp = "" Then
                    txtCustomerCode.Focus()
                Else
                    txtCustomerCode.Text = strCustHelp
                    lblCustomerdesc.Text = GetQryOutput("SELECT CUST_NAME FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & strCustHelp & "'")

                    'If SpASNAdjst.MaxRows = 0 Then
                    'Call AddBlankRowInGrid()
                    'End If

                End If
        End Select
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtCustomerCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.TextChanged
        With SpASNAdjst
            .MaxRows = 0
            If Me.txtCustomerCode.Text.Length = 0 Then Me.lblCustomerdesc.Text = String.Empty
        End With
    End Sub

    Private Sub txtCustomerCode_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
        On Error GoTo ErrHandler
        Dim strLocCondition As String
        Dim Cancel As Boolean = e.Cancel
        Dim strSql As String

        Select Case CmdASNAdjustBttn.Mode
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

    Private Sub SpASNAdjst_ClickEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles SpASNAdjst.ClickEvent

        On Error GoTo ErrHandler

        Dim blnStatus As Boolean
        Dim strSql As String
        Dim strItemcode As String
        Dim strPartcode As String

        Select Case CmdASNAdjustBttn.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                With SpASNAdjst
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

    Private Sub cmdInvoicetype_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInvoicetype.Click
        Dim StrINVHelp As String
        On Error GoTo ErrHandler

        If txtCustomerCode.Text.Length <= 0 Then
            MsgBox("Please Select First Customer ", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
            txtCustomerCode.Focus()
            Exit Sub
        End If

        StrINVHelp = ShowList(0, Len(txtinvoicetype.Text), txtinvoicetype.Text, "Invoice_type", "" & DateColumnNameInShowList("Description") & " as Description", "SALECONF", " ", "INVOICE TYPES")
        If StrINVHelp = "-1" Then

            Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            
            cmdInvoicetype.Enabled = True

        ElseIf StrINVHelp = "" Then
            Me.txtinvoicetype.Focus()
        Else
        
            If SpASNAdjst.MaxRows = 0 Then
                Call AddBlankRowInGrid()
            End If

            Me.txtinvoicetype.Text = StrINVHelp
            txtinvoicetype_Validating(txtinvoicetype, New System.ComponentModel.CancelEventArgs(False))
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtinvoicetype_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtinvoicetype.TextChanged
        With SpASNAdjst
            .MaxRows = 0
        End With
    End Sub


    Private Sub txtinvoicetype_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtinvoicetype.Validating
        On Error GoTo ErrHandler
        Dim strLocCondition As String
        Dim Cancel As Boolean = e.Cancel
        Dim strSql As String

        Select Case CmdASNAdjustBttn.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If (txtinvoicetype.Text.Trim.Length > 0) Then
                    strSql = "SELECT INVOICE_TYPE FROM SALECONF " & _
                    " WHERE UNIT_CODE='" & gstrUNITID & "' AND INVOICE_TYPE='" & txtinvoicetype.Text.Trim & "'"
                    txtinvoicetype.Text = GetQryOutput(strSql)
                    If Me.lblCustomerdesc.Text.Length <= 0 Then
                        MsgBox("Please Select First Customer ", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                        txtinvoicetype.Text = String.Empty
                        Exit Sub
                    Else
                        If (txtinvoicetype.Text.Trim.Length > 0) Then
                            If SpASNAdjst.MaxRows = 0 Then
                                Call AddBlankRowInGrid()
                            End If
                        Else
                            txtinvoicetype.Text = String.Empty
                            Call ConfirmWindow(10010)
                            txtinvoicetype.Clear()
                            Cancel = True

                        End If
                    End If

                End If
        End Select
        GoTo EventExitSub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        e.Cancel = Cancel
    End Sub
    Private Function Generate_docNo() As Integer
        On Error GoTo ErrHandler
        Dim objisexists As ClsResultSetDB
        Dim objdocno As New ClsResultSetDB
        Dim strDocNo As String
        Dim lngDocNo As Integer
        Dim strTempSeries As String
        Dim strFin_Start_Date As Date
        Dim strFin_End_Date As Date
        strDocNo = "select isnull(current_no,0)  as current_no,Fin_Start_Date,Fin_End_Date from documenttype_mst where doc_type=9997 and '" & VB6.Format(lblDisplayDate.Text, "dd mmm Yyyy") & "'  between Fin_Start_date and Fin_end_date AND UNIT_CODE = '" & gstrUNITID & "'"
        objdocno.GetResult(strDocNo, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not objdocno.EOFRecord Then
            lngDocNo = Val(objdocno.GetValue("current_no")) + 1
            strFin_Start_Date = objdocno.GetValue("Fin_Start_Date").ToString
            strFin_End_Date = objdocno.GetValue("Fin_End_Date").ToString()
        Else
            Generate_docNo = -1
            objdocno.ResultSetClose()
            objdocno = Nothing
            Exit Function
        End If
        objdocno.ResultSetClose()
        objdocno = Nothing

        mstrUpdateDocumentNoSQL = "UPDATE documenttype_mst SET current_no=" & Val(CStr(lngDocNo))
        mstrUpdateDocumentNoSQL = mstrUpdateDocumentNoSQL & " WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_type=9997 and '" & VB6.Format(Me.lblDisplayDate.Text, "dd mmm Yyyy") & "'  between fin_start_date and fin_end_date"
        mP_Connection.Execute(mstrUpdateDocumentNoSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        strTempSeries = lngDocNo
        Generate_docNo = strTempSeries
        Exit Function
ErrHandler:
        Generate_docNo = -1
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        objisexists = Nothing
        objdocno = Nothing
    End Function
End Class