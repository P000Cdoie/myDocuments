

Option Strict Off
Option Explicit On
Imports System.Data.SqlClient

Friend Class frmMKTMST0004
	Inherits System.Windows.Forms.Form
	
	'-----------------------------------------------------------------------------
	'Copyright(c)   - MIND
	'Name of module - frmMKTMST0004
	'Created by     -
	'Created Date   -
	'Description -    Sale Terms Master
	'Revised date - 16-01-2002 changed in changed CmdGRp function Commented code which
	' sets enable Property of Edit & Delete Button to False after rewerting the control
	'ref Checkout Form No is 4004
	'13-02-2002 changed in serial no help. chacked out form no 4046
	'------------------------------------------------------------------------------
    'Revised by     :   Vinod singh
    'Revision date  :   21/04/2011
    'Reason         :   Multi Unit changes
    '-------------------------------------------------------------------------------
	Dim strSQL As String
    Dim mRdoCls As ClsResultSetDB
	Dim blnDisp_rrecflag As Boolean
	Dim mintSerial_no As Short
	Dim blnCheckkey As Boolean
	Dim blnUnload As Boolean
	Dim strTermCode As String
	Dim strCheckCode As String
	Dim mIntSerialNo As Short
	Dim strErrMsg As String 'Stores the error message
	Dim intLineNumber As Short 'Stores the line number
	Dim ctlBlank As System.Windows.Forms.Control 'Stores reference to a control
	Dim mintFormIndex As Short
	
    Private Sub cmbType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbType.SelectedIndexChanged
        Select Case Me.CmdGrp1.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                'Call autoSerial_No()
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                txtSerial_No.Text = ""
                lblGlobalSerialNo.Text = ""
                lblGlobalTermtypeDesc.Text = ""
                Cmdhelp.Enabled = True
        End Select
        Me.CmdGrp1.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True

    End Sub
    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cmdhelp.Click
        On Error GoTo errHandler
        Dim strCheckRec As String
        On Error GoTo errHandler
        If CmdGrp1.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            Call procCheckTermsCode()
            If Len(txtSerial_No.Text) = 0 Then
                strCheckRec = ShowList(1, (txtSerial_No.MaxLength), "", "Serial_No", "Description", "saleterms_mst", " and saleterms_type='" & Trim(cmbType.SelectedValue.ToString) & "'")
            Else
                strCheckRec = ShowList(1, (txtSerial_No.MaxLength), txtSerial_No.Text, "Serial_No", "Description", "saleterms_mst", " and saleterms_type='" & Trim(cmbType.SelectedValue.ToString) & "'")
            End If
            If strCheckRec = "-1" Then
                Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                txtSerial_No.Text = "" : txtSerial_No.Focus()
            Else
                txtSerial_No.Text = strCheckRec
            End If
            txtSerial_No.Focus()
            Call txtSerial_No_Validating(txtSerial_No, New System.ComponentModel.CancelEventArgs(False))
            Exit Sub
        End If
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
       
    End Sub
    Public Sub getGlobalValue()

        On Error GoTo ErrHandler
        Dim dt As DataTable = New DataTable()
        Dim con As SqlConnection = New SqlConnection()
        con = SqlConnectionclass.GetConnection()
        Dim adp As SqlDataAdapter = New SqlDataAdapter("select GLOBAL_TERM_SLNO ,GLOBAL_TERM_TYPE,TERMTYPE_DESC  from SaleTerms_Mst inner join GLOBAL_SALE_TERMTYPE_MST on GLOBAL_TERM_TYPE=TERMTYPE_CODE  where SaleTerms_Type =@salterm and UNIT_CODE =@unitcode and Serial_No =@serialNo and GLOBAL_TERM_TYPE is not null and GLOBAL_TERM_SLNO is not null", con)
        adp.SelectCommand.Parameters.AddWithValue("@salterm", cmbType.SelectedValue.ToString())
        adp.SelectCommand.Parameters.AddWithValue("@unitcode", gstrUNITID)
        adp.SelectCommand.Parameters.AddWithValue("@serialNo", txtSerial_No.Text)
        adp.SelectCommand.CommandType = CommandType.Text
        adp.Fill(dt)
        If dt.Rows.Count > 0 Then
            lblGlobalSerialNo.Text = dt.Rows(0)(0).ToString()
            lblGlobalTermtypeDesc.Text = dt.Rows(0)(2).ToString()
            lblGlobalTermtypeDesc.Tag = dt.Rows(0)(1).ToString()
            cmdHelpEditCase.Enabled = False
            cmdHelpGlobalSerialNo.Enabled = False
        Else
            cmdHelpEditCase.Enabled = True
        End If
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub

    Private Sub frmMKTMST0004_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo errHandler
        'Checking the form name in the Windows list
        mdifrmMain.CheckFormName = mintFormIndex
        If Me.cmbType.Enabled = True Then Me.cmbType.Focus()
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTMST0004_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo errHandler
        'Make the node normal font
        frmModules.NodeFontBold(Tag) = False
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTMST0004_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlFormHeader1_Click(ctlFormHeader1, New System.EventArgs()) : Exit Sub
        End If
    End Sub
    Private Sub frmMKTMST0004_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            System.Windows.Forms.SendKeys.SendWait("{TAB}")
        ElseIf KeyAscii = System.Windows.Forms.Keys.Escape Then
            blnDisp_rrecflag = False
            If CmdGrp1.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION, 60095) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Then GoTo EventExitSub
                Call ChangeCmdgrp()
                GoTo EventExitSub
            End If
            Call ChangeCmdgrp()
        End If
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub frmMKTMST0004_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo errHandler
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag) 'For Form Tag Written by MUk on 19/03/2001
        Call FillLabelFromResFile(Me) 'To Fill label description from Resource file
        Call FitToClient(Me, Frame1, ctlFormHeader1, CmdGrp1) 'To fit the form in the MDI
        Call EnableControls(False, Me, True)
        ' Call AddTypesToComboBox()
        Call gettermtypeforEdit(String.Empty)
        blnCheckkey = False
        blnDisp_rrecflag = True
        CmdGrp1.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
        CmdGrp1.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
        CmdGrp1.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        Cmdhelp.Enabled = True : cmbType.Enabled = True : cmbType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        txtSerial_No.Enabled = False : txtSerial_No.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        cmdHelpEditCase.Enabled = False
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Function Savefunction() As Boolean
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Boolean Value
        'Function       :   Savefunction
        'Comments       :   Checking valid data at save points
        '----------------------------------------------------------------------------
        On Error GoTo errHandler
        Savefunction = True
        strErrMsg = ResolveResString(10059) & vbCrLf & vbCrLf
        If cmbType.Text = "" Then
            intLineNumber = intLineNumber + 1
            strErrMsg = strErrMsg & intLineNumber & ") " & ResolveResString(60044) & vbCrLf
            Savefunction = False
            If ctlBlank Is Nothing Then ctlBlank = cmbType
        End If
        If txtSerial_No.Text = "" Then
            intLineNumber = intLineNumber + 1
            strErrMsg = strErrMsg & intLineNumber & ") " & ResolveResString(60045) & vbCrLf
            Savefunction = False
            If ctlBlank Is Nothing Then ctlBlank = cmbType
        End If
        If Trim(txtdes.Text) = "" Then
            intLineNumber = intLineNumber + 1
            strErrMsg = strErrMsg & intLineNumber & ") " & ResolveResString(60012) & vbCrLf
            Savefunction = False
            If ctlBlank Is Nothing Then ctlBlank = txtdes
        End If
        Exit Function
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Public Function validateForUpdate() As Boolean
        On Error GoTo errHandler
        validateForUpdate = True
        strErrMsg = ResolveResString(10059) & vbCrLf & vbCrLf
        If lblGlobalSerialNo.Text = "" Then
            intLineNumber = intLineNumber + 1
            strErrMsg = strErrMsg & intLineNumber & ") " & ResolveResString(60044) & vbCrLf
            validateForUpdate = False
            If ctlBlank Is Nothing Then ctlBlank = lblGlobalSerialNo
        End If
        If lblGlobalTermtypeDesc.Text = "" Then
            intLineNumber = intLineNumber + 1
            strErrMsg = strErrMsg & intLineNumber & ") " & ResolveResString(60045) & vbCrLf
            validateForUpdate = False
            If ctlBlank Is Nothing Then ctlBlank = lblGlobalTermtypeDesc
        End If
      
        Exit Function
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Private Sub frmMKTMST0004_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        On Error GoTo errHandler
        'Declarations
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        blnUnload = False
        If CmdGrp1.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If UnloadMode >= 0 And UnloadMode <= 5 Then
                enmValue = ConfirmWindow(10055, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Or enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call CmdGrp1_ButtonClick(eventSender, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE))
                        If blnUnload = False Then
                            gblnCancelUnload = True
                        End If
                    Else
                        gblnCancelUnload = False
                        CmdGrp1.Focus()
                    End If
                Else
                    'Set the global variable
                    gblnCancelUnload = True
                    blnUnload = True
                    If CmdGrp1.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                        txtdes.Focus()
                    Else
                        CmdGrp1.Focus()
                    End If
                End If
            End If
        Else
            gblnCancelUnload = False
        End If
        If blnUnload = False Then
            gblnCancelUnload = False
        Else
            gblnCancelUnload = True
        End If
        'Checking the status
        If gblnCancelUnload = True Then eventArgs.Cancel = True
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTMST0004_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        frmModules.NodeFontBold(Tag) = False
        'Removing the form name from list
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        Me.Dispose()
    End Sub
    Private Sub txtdes_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtdes.Enter
        txtdes.SelectionStart = 0
        txtdes.SelectionLength = Len(txtdes.Text)
        If cmbType.SelectedIndex = 0 Then
            txtdes.MaxLength = 100
        Else
            txtdes.MaxLength = 60
        End If
    End Sub
    Private Sub txtDes_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtdes.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtSerial_No_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSerial_No.TextChanged
        txtSerial_No.Text = Number_Chk(txtSerial_No.Text) 'Checking Numaric Value
        If Len(txtSerial_No.Text) = 0 And Me.CmdGrp1.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            txtdes.Text = ""
            lblGlobalTermtypeDesc.Text = ""
            lblGlobalSerialNo.Text = ""
            Me.CmdGrp1.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
            Me.CmdGrp1.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
            Me.CmdGrp1.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        End If
    End Sub
    Private Sub txtSerial_No_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSerial_No.Enter
        txtSerial_No.SelectionStart = 0
        txtSerial_No.SelectionLength = Len(txtSerial_No.Text)
        Call procCheckTermsCode()
    End Sub
    Private Sub txtSerial_No_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSerial_No.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case 13
                If Len(txtSerial_No.Text) > 0 Then
                    Call txtSerial_No_Validating(txtSerial_No, New System.ComponentModel.CancelEventArgs(False))
                Else
                    Me.CmdGrp1.Focus()
                End If
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtSerial_No_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSerial_No.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = 112 Then Call cmdHelp_Click(Cmdhelp, New System.EventArgs())
    End Sub
    Private Sub txtSerial_No_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtSerial_No.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo errHandler
        Select Case Me.CmdGrp1.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                If Len(txtSerial_No.Text) > 0 Then
                    If funcDisplayRecord() Then
                        Me.CmdGrp1.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                        Me.CmdGrp1.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                        Me.CmdGrp1.Focus()
                        txtSerial_No.Enabled = True
                    Else
                        Me.CmdGrp1.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                        Me.CmdGrp1.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                        Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Cancel = True
                        txtSerial_No.Text = "" : txtSerial_No.Focus()
                    End If
                End If
        End Select
        GoTo EventExitSub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub cmbType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cmbType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrp1.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        txtdes.Focus()
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
   
    Private Sub cmbType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles cmbType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        If cmbType.Text <> "" Then
            Call procCheckTermsCode()
            '  Call autoSerial_No() 'This function generate auto serial number
            blnCheckkey = True
        End If
        eventArgs.Cancel = Cancel
    End Sub
    Public Sub RefreshForm()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Procedure      :   RefreshForm
        'Comments       :  Clear Fields
        '----------------------------------------------------------------------------
        txtSerial_No.Text = ""
        txtdes.Text = ""
        lblGlobalSerialNo.Text = ""
        lblGlobalTermtypeDesc.Text = ""
    End Sub
    Public Sub ChangeCmdgrp()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Procedure      :  ChangeCmdgrp
        'Comments       :  Change Caption cmdgrp control's
        '----------------------------------------------------------------------------
        On Error GoTo errHandler
        blnCheckkey = False
        Call EnableControls(True, Me)
        blnDisp_rrecflag = True
        CmdGrp1.Revert()
        CmdGrp1.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        cmbType.Enabled = True : cmbType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        txtSerial_No.Enabled = False : txtSerial_No.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        gblnCancelUnload = False : gblnFormAddEdit = False
        txtdes.Enabled = False : txtdes.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        cmbType.Focus()
        cmdHelpEditCase.Enabled = False
        cmdHelpGlobalSerialNo.Enabled = False
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub procCheckTermsCode()
        On Error GoTo errHandler
        'Select Case cmbType.SelectedIndex
        '    Case 0
        '        strTermCode = "PY"
        '    Case 1
        '        strTermCode = "PR"
        '    Case 2
        '        strTermCode = "PK"
        '    Case 3
        '        strTermCode = "FR"
        '    Case 4
        '        strTermCode = "TR"
        '    Case 5
        '        strTermCode = "OC"
        '    Case 6
        '        strTermCode = "MO"
        '    Case 7
        '        strTermCode = "DL"
        'End Select
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Function funcDisplayRecord() As Boolean
        On Error GoTo errHandler
        funcDisplayRecord = False
        mIntSerialNo = IIf(txtSerial_No.Text = "", 0, txtSerial_No.Text)
        strSQL = "select Description ,GLOBAL_TERM_TYPE ,GLOBAL_TERM_SLNO,TERMTYPE_DESC from saleterms_mst left outer join GLOBAL_SALE_TERMTYPE_MST on GLOBAL_TERM_TYPE=GLOBAL_SALE_TERMTYPE_MST.TERMTYPE_CODE  where unit_code='" & gstrUNITID & "' and  saleterms_type='" & Trim(cmbType.SelectedValue.ToString) & "' and Serial_No=" & Val(txtSerial_No.Text) & ""
        mRdoCls = New ClsResultSetDB
        If mRdoCls.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset) And mRdoCls.GetNoRows > 0 Then
            funcDisplayRecord = True
            txtdes.Text = mRdoCls.GetValue("Description")
            If Not IsDBNull(mRdoCls.GetValue("GLOBAL_TERM_TYPE")) Then
                lblGlobalTermtypeDesc.Text = mRdoCls.GetValue("TERMTYPE_DESC")
            Else
                lblGlobalTermtypeDesc.Text = ""
            End If
            If Not IsDBNull(mRdoCls.GetValue("GLOBAL_TERM_SLNO")) Then
                lblGlobalSerialNo.Text = mRdoCls.GetValue("GLOBAL_TERM_SLNO")
            Else
                lblGlobalSerialNo.Text = ""
            End If
        Else
            funcDisplayRecord = False
        End If
            mRdoCls.ResultSetClose()
            Exit Function
errHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
            Exit Function
    End Function
    Private Sub AddTypesToComboBox()
        On Error GoTo errHandler

        getComboSalesTypeValues()
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub

    Public Sub getComboSalesTypeValues()
        On Error GoTo errHandler
        Dim dtSalesType As DataTable
        dtSalesType = New DataTable()
        Dim con As SqlConnection = New SqlConnection()
        con = SqlConnectionclass.GetConnection()
        Dim sqlstring As String
        cmbType.DataSource = Nothing
        sqlstring = "  SELECT A.TERM_TYPE as TERMTYPE_CODE,B.TERMTYPE_DESC FROM( " & _
              " SELECT DISTINCT  TERM_TYPE FROM  GLOBAL_SALE_TERM_MST INNER JOIN " & _
              " GLOBAL_MASTER_MAPPING ON GLOBAL_MASTER_MAPPING.Global_SLNO = GLOBAL_SALE_TERM_MST.SLNO  " & _
              "WHERE(GLOBAL_SALE_TERM_MST.ISACTIVE = 1) " & _
              "and LTRIM(RTRIM(GLOBAL_MASTER_MAPPING.TableName)) = LTRIM(RTRIM('GLOBAL_SALE_TERM_MST'))  " & _
              "and  GLOBAL_MASTER_MAPPING.Unit_Code ='" & gstrUNITID & "' ) AS A," & _
              "GLOBAL_SALE_TERMTYPE_MST B WHERE(A.TERM_TYPE = B.TERMTYPE_CODE)"

        Dim sqladp As SqlDataAdapter = New SqlDataAdapter(sqlstring, con)
        sqladp.Fill(dtSalesType)
        If dtSalesType.Rows.Count > 0 Then
            cmbType.DataSource = dtSalesType
            cmbType.ValueMember = "TERMTYPE_CODE"
            cmbType.DisplayMember = "TERMTYPE_DESC".Trim

            cmbType.SelectedIndex = 0

        End If
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub gettermtypeforEdit(ByVal strTypecode As String)
        On Error GoTo errHandler
        Dim dtSalesType As DataTable
        dtSalesType = New DataTable()
        Dim con As SqlConnection = New SqlConnection()
        con = SqlConnectionclass.GetConnection()
        Dim sqlstring As String
        cmbType.DataSource = Nothing
        sqlstring = "select distinct termtype_code,termtype_desc from GLOBAL_SALE_TERMTYPE_MST" & _
                   " where  termtype_code in( SELECT distinct saleterms_type from " & _
                    "SaleTerms_Mst where UNIT_CODE ='" & gstrUNITID & "') "
        Dim sqladp As SqlDataAdapter = New SqlDataAdapter(sqlstring, con)
        sqladp.Fill(dtSalesType)
        If dtSalesType.Rows.Count > 0 Then
            cmbType.DataSource = dtSalesType
            cmbType.ValueMember = "TERMTYPE_CODE".Trim
            cmbType.DisplayMember = "TERMTYPE_DESC".Trim
            cmbType.SelectedIndex = 0
        End If
        If strTypecode <> String.Empty Then
            cmbType.SelectedValue = strTypecode
        End If
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub CmdGrp1_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles CmdGrp1.ButtonClick
        ' ---------------------------------------------------------------------
        '     Purpose   : This Method is used to Add/Edit/Delete/Save/Print
        '     Parameter : A read only parameter, it gives you the index of the
        '                 button clicked by the user.
        '    ---------------------------------------------------------------------
        On Error GoTo errHandler
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                Call EnableControls(True, Me, True)
                Call RefreshForm()
                Call AddTypesToComboBox()

                txtSerial_No.Enabled = False
                txtSerial_No.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                ' Cmdhelp.Enabled = False
                txtdes.Enabled = False
                blnDisp_rrecflag = False
                blnCheckkey = False
                ' cmbType.Focus()
                cmbType.Enabled = False
                cmdHelpEditCase.Enabled = True
                cmdHelpGlobalSerialNo.Enabled = False
                Cmdhelp.Enabled = False

                'Added By ekta uniyal on 29 Mar 2014 to support multi-unit functionality for Hilex
                CmdGrp1.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                'End Here

                Exit Sub
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                Call gettermtypeforEdit(cmbType.SelectedValue.ToString)

                'Call EnableControls(True, Me)

                getGlobalValue()
                '' GroupBoxGlobal.Visible = True
                cmbType.Enabled = False : cmbType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtSerial_No.Enabled = False : txtSerial_No.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtdes.Enabled = False
                txtdes.Focus()
                Cmdhelp.Enabled = False
                Exit Sub
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
                If Len(Trim(cmbType.Text)) <= 0 Then
                    Call EnableControls(True, Me)
                    txtdes.Enabled = False
                    cmbType.Focus()
                    blnCheckkey = False
                    blnDisp_rrecflag = True
                    Exit Sub
                End If
                strSQL = "DELETE FROM saleterms_mst WHERE unit_code='" & gstrUNITID & "' and saleterms_type = '" & Trim(strTermCode) & "' and Serial_No=" & mIntSerialNo & ""
                If ConfirmWindow(10054, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_CRITICAL, 60096) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    With mP_Connection
                        .Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End With
                    Call ChangeCmdgrp()
                    Call RefreshForm()
                    Call ConfirmWindow(10051, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    Me.cmbType.Focus()
                Else
                    CmdGrp1.Focus()
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                txtdes.Enabled = False
                If Not Savefunction() Then
                    Call MsgBox(strErrMsg, MsgBoxStyle.Information, "Error")
                    ctlBlank.Focus()
                    blnUnload = True
                    ctlBlank = Nothing
                    intLineNumber = 0
                    Exit Sub
                End If
                Select Case CmdGrp1.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        mIntSerialNo = Val(txtSerial_No.Text)
                        strSQL = "insert into saleterms_mst(saleterms_type,Serial_No,Description,ent_userid,ent_dt,upd_userid,upd_dt,unit_code,GLOBAL_TERM_TYPE,GLOBAL_TERM_SLNO)" & " values('" & cmbType.SelectedValue.ToString() & "'," & Convert.ToInt32(txtSerial_No.Text.Trim) & ",'" & Trim(txtdes.Text) & "','" & mP_User & "',getdate(),'" & mP_User & "',getdate(),'" & gstrUNITID & "','" & lblGlobalTermtypeDesc.Tag.trim & "','" & Convert.ToInt32(txtSerial_No.Text) & "')"
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        'If Not validateForUpdate() Then
                        '    Call MsgBox(strErrMsg, MsgBoxStyle.Information, "Error")
                        '    Exit Sub
                        'End If
                        mIntSerialNo = Val(txtSerial_No.Text)
                        If lblGlobalTermtypeDesc.Text <> "" And lblGlobalSerialNo.Text <> "" Then
                            strSQL = "update saleterms_mst set  Description='" & Trim(txtdes.Text) & "',ent_userid='" & mP_User & "',ent_dt=getdate(), global_term_type='" & lblGlobalTermtypeDesc.Tag.Trim & "',global_term_slno=" & Convert.ToInt32(lblGlobalSerialNo.Text) & ",upd_userid='" & mP_User & "',upd_dt=getdate()" & " where unit_code='" & gstrUNITID & "' and saleterms_type = '" & Trim(cmbType.SelectedValue.ToString) & "' and Serial_No=" & Convert.ToInt32(txtSerial_No.Text.Trim) & ""
                        ElseIf lblGlobalTermtypeDesc.Text.Trim = "" And lblGlobalSerialNo.Text.Trim = "" Then
                            strSQL = "update saleterms_mst set  Description='" & Trim(txtdes.Text) & "',ent_userid='" & mP_User & "',ent_dt=getdate(), upd_userid='" & mP_User & "',upd_dt=getdate()" & " where unit_code='" & gstrUNITID & "' and saleterms_type = '" & Trim(cmbType.SelectedValue.ToString) & "' and Serial_No=" & Convert.ToInt32(txtSerial_No.Text.Trim) & ""
                        End If
                        Cmdhelp.Enabled = True
                End Select
                With mP_Connection
                    .Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End With
                Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Call ChangeCmdgrp()
                CmdGrp1.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                CmdGrp1.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                cmdHelpEditCase.Enabled = False
                cmdHelpGlobalSerialNo.Enabled = False

                Cmdhelp.Enabled = False
                If CmdGrp1.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    txtSerial_No.Enabled = False
                    Cmdhelp.Enabled = False
                End If
                Exit Sub
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
                strSQL = "DELETE FROM saleterms_mst WHERE unit_code='" & gstrUNITID & "' and saleterms_type = '" & Trim(cmbType.SelectedValue.ToString) & "' and Serial_No=" & mIntSerialNo & ""
                If ConfirmWindow(10003) = MsgBoxResult.Yes Then
                    If DeleteRecordFromTable(strSQL) Then
                        blnDisp_rrecflag = True
                        Call RefreshForm()
                        Call EnableControls(False, Me)
                        Call Me.CmdGrp1.Revert()
                        blnCheckkey = False
                    End If
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION, 60095) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Then
                    If CmdGrp1.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                        cmbType.Focus()
                        Exit Sub
                    Else
                        txtdes.Focus()
                        If lblGlobalTermtypeDesc.Text <> "" Then
                            cmdHelpEditCase.Enabled = False
                            cmdHelpGlobalSerialNo.Enabled = False
                        Else
                            cmdHelpEditCase.Enabled = True
                        End If

                        Exit Sub
                    End If
                End If
                Call ChangeCmdgrp()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
                blnCheckkey = False
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
        End Select
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm") '("saleterms_master.htm")
    End Sub


    Private Sub cmdHelpEditCase_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelpEditCase.Click
        On Error GoTo ErrHandler
        Dim sqlstring As String
        Dim strHelp() As String
        sqlstring = "  SELECT A.TERM_TYPE,B.TERMTYPE_DESC FROM( " & _
             " SELECT DISTINCT  TERM_TYPE FROM  GLOBAL_SALE_TERM_MST INNER JOIN " & _
             " GLOBAL_MASTER_MAPPING ON GLOBAL_MASTER_MAPPING.Global_SLNO = GLOBAL_SALE_TERM_MST.SLNO  " & _
             "WHERE(GLOBAL_SALE_TERM_MST.ISACTIVE = 1) " & _
             "and LTRIM(RTRIM(GLOBAL_MASTER_MAPPING.TableName)) = LTRIM(RTRIM('GLOBAL_SALE_TERM_MST'))  " & _
             "and  GLOBAL_MASTER_MAPPING.Unit_Code ='" & gstrUNITID & "' ) AS A," & _
             "GLOBAL_SALE_TERMTYPE_MST B WHERE(A.TERM_TYPE = B.TERMTYPE_CODE)"

        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        strHelp = CtlHelpGlobalTerm.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, sqlstring, "List of Term(s)", 1)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(strHelp) = -1 Then Exit Sub
        If strHelp(0) = "0" Then
            MsgBox("No Term Type(s) are defined.", MsgBoxStyle.Information, ResolveResString(100))
            ' txtRecLoc.Focus()
        Else
            lblGlobalTermtypeDesc.Text = strHelp(1)
            lblGlobalTermtypeDesc.Tag = strHelp(0)
            cmdHelpGlobalSerialNo.Enabled = True
        End If
        'If lblGlobalTermtypeDesc.Text <> "" Then
        '    cmdHelpGlobalSerialNo.Enabled = True
        'Else
        '    cmdHelpGlobalSerialNo.Enabled = False
        'End If
        If CmdGrp1.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD = True Then
            cmbType.SelectedValue = strHelp(0)

            'Added By ekta uniyal on 29 Mar 2014 to support multi-unit functionality for Hilex
            CmdGrp1.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
            'End Here

        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub cmdHelpGlobalSerialNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelpGlobalSerialNo.Click
        On Error GoTo ErrHandler
        Dim sqlstring As String
        Dim strHelp() As String
        If CmdGrp1.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT = True Then
            sqlstring = "select  TERM_SLNO,TERM_DESCRIPTION  from [GLOBAL_SALE_TERMTYPE_MST]inner join " & _
                   "GLOBAL_SALE_TERM_MST on [GLOBAL_SALE_TERMTYPE_MST].[TERMTYPE_CODE]= " & _
       "GLOBAL_SALE_TERM_MST.term_type inner join  [GLOBAL_MASTER_MAPPING] on [GLOBAL_MASTER_MAPPING].Global_slno= " & _
       "GLOBAL_SALE_TERM_MST.SLNO  " & _
       "and ltrim(rtrim([GLOBAL_MASTER_MAPPING].TableName)) =  LTRIM (rtrim('GLOBAL_SALE_TERM_MST')) " & _
       "and  [GLOBAL_MASTER_MAPPING].Unit_Code ='" & gstrUNITID & "' where (ISACTIVE = 1) and  " & _
       "GLOBAL_SALE_TERMTYPE_MST.termtype_code='" & lblGlobalTermtypeDesc.Tag.Trim & "'  and not Exists(select GLOBAL_TERM_SLNO,GLOBAL_TERM_TYPE from " & _
       " (SELECT GLOBAL_TERM_SLNO,GLOBAL_TERM_TYPE  from SaleTerms_Mst Where Unit_Code ='" & gstrUNITID & "' and GLOBAL_TERM_SLNO is not null) as A " & _
       " where(a.GLOBAL_TERM_SLNO =GLOBAL_SALE_TERM_MST.TERM_SLNO)  and a.GLOBAL_TERM_TYPE ='" & lblGlobalTermtypeDesc.Tag.Trim & "') "


        Else
            sqlstring = "select  TERM_SLNO,TERM_DESCRIPTION  from [GLOBAL_SALE_TERMTYPE_MST]inner join " & _
                   "GLOBAL_SALE_TERM_MST on [GLOBAL_SALE_TERMTYPE_MST].[TERMTYPE_CODE]= " & _
       "GLOBAL_SALE_TERM_MST.term_type inner join  [GLOBAL_MASTER_MAPPING] on [GLOBAL_MASTER_MAPPING].Global_slno= " & _
       "GLOBAL_SALE_TERM_MST.SLNO  " & _
       "and ltrim(rtrim([GLOBAL_MASTER_MAPPING].TableName)) =  LTRIM (rtrim('GLOBAL_SALE_TERM_MST')) " & _
       "and  [GLOBAL_MASTER_MAPPING].Unit_Code ='" & gstrUNITID & "' where (ISACTIVE = 1) and  " & _
       "GLOBAL_SALE_TERMTYPE_MST.termtype_code='" & lblGlobalTermtypeDesc.Tag.Trim & "'  and not Exists(select Serial_No,SaleTerms_Type from " & _
       " ( SELECT Serial_No,SaleTerms_Type   from SaleTerms_Mst Where Unit_Code ='" & gstrUNITID & "' union " & _
       " SELECT GLOBAL_TERM_SLNO,GLOBAL_TERM_TYPE  from SaleTerms_Mst Where Unit_Code ='" & gstrUNITID & "') as A " & _
       " where(a.Serial_No =GLOBAL_SALE_TERM_MST.TERM_SLNO)  and a.SaleTerms_Type ='" & lblGlobalTermtypeDesc.Tag.Trim & "') "


        End If
       
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        strHelp = CtlHelGlobalSerial.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, sqlstring, "List of Global Serial No.(s)", 1)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(strHelp) = -1 Then Exit Sub
        If strHelp(0) = "0" Then
            MsgBox("No Global Serial No.(s) are defined.", MsgBoxStyle.Information, ResolveResString(100))
            ' txtRecLoc.Focus()
        Else
            lblGlobalSerialNo.Text = strHelp(0)
            lblGlobalSerialNo.Tag = strHelp(0)
        End If
        If CmdGrp1.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD = True Then
            txtSerial_No.Text = strHelp(0)
            txtdes.Text = strHelp(1)


        End If
        If lblGlobalTermtypeDesc.Text <> "" Then
            cmdHelpGlobalSerialNo.Enabled = True
        Else
            cmdHelpGlobalSerialNo.Enabled = False
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    
End Class