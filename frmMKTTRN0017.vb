Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Friend Class frmMKTTRN0017
	Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
	' File Name         :   FRMMKTTRN0017.frm
	' Function          :   Used to Print Bin card
	' Created By        :   Nisha Rai
	' Created On        :   05 June, 2003
	' Revision History  :   changed on 10 Sept 2003 By Nisha
	'                       1. Report file Parameter
	'                       2. BinQuantity on screen
	'                   :   Changed on 19/09/2003 By Nisha
	'                       1.Corrected the Bin Quantity Label
	'                   :   Changed on 02/01/2004 By Nisha
	'                       1.to increase maxlength of bin quantity to 6 chars
	'---------------------------------------------------------------------------------------------------------
	' Revised By         :   Davinder Singh
	' Revision Date      :   31/03/2006
	' Issue ID           :   17445
	' Revision History   :   To print or not to print the Invoice No. prefix according to the PrintPrefix flag in Sales_Parameter
	'---------------------------------------------------------------------------------------------------------
	' Revision By       :   Davinder Singh
	' Revision Date     :   01/05/2006
	' Issue ID          :   17707
	' Revision History  :   To pick the Bin Qty. from Sales_dtl instead of Custitem_mst
	'---------------------------------------------------------------------------------------------------------
    ' Revised By                 -   Roshan Singh
    ' Revision Date              -   03 JUN 2011
    ' Description                -   FOR MULTIUNIT FUNCTIONALITY
    '---------------------------------------------------------------------------------------------------------
    Dim mintFormIndex As Double
    Dim mBoxQuantity As Double
    Dim mSalesQuantity As Double

    Private Sub cmdHelpInvoice_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelpInvoice.Click
        Dim varHelp As Object
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        varHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT Doc_no,Invoice_date FROM Saleschallan_dtl WHERE  UNIT_CODE = '" & gstrUNITID & "' and Bill_flag =1 and Cancel_flag =0 and Location_code = '" & Trim(txtUnitCode.Text) & "' and Doc_no like '" & Trim(txtInvoice.Text) & "%'")
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(varHelp) <> -1 Then
            If varHelp(0) <> "0" Then
                txtInvoice.Text = Trim(varHelp(0))
                txtInvoice.Focus()
            Else
                MsgBox("No Record Available", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdInvoice_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As UCActXCtl.UCfraRepCmd.ButtonClickEventArgs) Handles Cmdinvoice.ButtonClick
        Dim NoOfLabels As Integer
        Dim mFrom As Short
        Dim mTo As Short
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim strLable As String
        Dim strDate As String
        Dim strLocation As String
        Dim strDINo As String
        Dim strsql As String
        Dim strReportFileName As String
        Dim strReportFilePath As String
        Dim intLastPageQuantity As Short
        Dim rsSalesChallnDtl As ClsResultSetDB
        Dim rsSalesParameter As ClsResultSetDB
        rsSalesChallnDtl = New ClsResultSetDB
        rsSalesParameter = New ClsResultSetDB
        Dim objRpt As ReportDocument
        Dim frmReportViewer As New eMProCrystalReportViewer
        objRpt = frmReportViewer.GetReportDocument()
        frmReportViewer.ShowPrintButton = True
        frmReportViewer.ShowTextSearchButton = True
        frmReportViewer.ShowZoomButton = True
        frmReportViewer.ReportHeader = Me.ctlFormHeader1.HeaderString()
        Dim dblQuantity As Double
        On Error GoTo ErrHandler
        If eventArgs.Button <> UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE Then
            If eventArgs.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_FILE Then
                MsgBox("Export is not available", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
            Call RefreshValues()
            If ValidatebeforeSave() = True Then
                rsSalesChallnDtl.GetResult("Select Invoice_date,SRVDINO,SrvLocation from SalesChallan_Dtl where  UNIT_CODE = '" & gstrUNITID & "' and doc_no = " & txtInvoice.Text)
                If rsSalesChallnDtl.GetNoRows > 0 Then
                    rsSalesChallnDtl.MoveFirst()
                    strDate = VB6.Format(rsSalesChallnDtl.GetValue("Invoice_date"), gstrDateFormat)
                    strDINo = rsSalesChallnDtl.GetValue("SRVDINO")
                    strLocation = rsSalesChallnDtl.GetValue("SRVLocation")
                End If
                rsSalesChallnDtl.ResultSetClose()
                rsSalesChallnDtl = Nothing
                strsql = "{Sales_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {Sales_Dtl.Unit_Code}='" & gstrUNITID & "' and {Sales_Dtl.Doc_No} =" & Trim(txtInvoice.Text)
                If mBoxQuantity <> Val(txtbinQuantity.Text) Then
                    mBoxQuantity = Val(txtbinQuantity.Text)
                End If
                NoOfLabels = Int(mSalesQuantity / mBoxQuantity)
                If (mSalesQuantity - (Int(mSalesQuantity / mBoxQuantity) * mBoxQuantity)) > 0 Then
                    NoOfLabels = NoOfLabels + 1
                End If
                If optSelected.Checked = True Then
                    mFrom = CShort(txtpagefrom.Text)
                Else
                    mFrom = 1
                End If
                rsSalesParameter.ResultSetClose()
                rsSalesParameter = New ClsResultSetDB
                rsSalesParameter.GetResult("Select BinCardFileName from Sales_parameter where UNIT_CODE = '" & gstrUNITID & "'")
                rsSalesParameter.MoveFirst()
                If Len(Trim(rsSalesParameter.GetValue("BinCardFileName"))) > 0 Then
                    strReportFileName = rsSalesParameter.GetValue("BinCardFileName")
                    If InStr(1, strReportFileName, ".") > 0 Then
                        strReportFileName = Mid(strReportFileName, 1, InStr(1, strReportFileName, ".") - 1)
                    End If
                Else
                    MsgBox("No Report File Name Defined in Sales_Parameter Table.", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If
                objRpt.Load(My.Application.Info.DirectoryPath & "\Reports\" & strReportFileName & ".rpt")
                objRpt.RecordSelectionFormula = strsql
                objRpt.DataDefinition.FormulaFields("InvoiceNo").Text = "'" & Print_InvoiceNumber(txtInvoice.Text) & "'"
                objRpt.DataDefinition.FormulaFields("InvoiceDate").Text = "'" & strDate & "'"
                objRpt.DataDefinition.FormulaFields("SRVLocation").Text = "'" & strLocation & "'"
                objRpt.DataDefinition.FormulaFields("SRVDINO").Text = "'" & strDINo & "'"
                objRpt.DataDefinition.FormulaFields("NoofBin").Text = "'" & NoOfLabels & "'"
                objRpt.DataDefinition.FormulaFields("BoxQuantity").Text = "'" & CStr(mBoxQuantity) & "'"
                rsSalesParameter.GetResult("Select DatePrintOnBinPrinting= Isnull(DatePrintOnBinPrinting,0) from Sales_parameter WHERE UNIT_CODE = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rsSalesParameter.RowCount > 0 Then
                    If rsSalesParameter.GetValue("DatePrintOnBinPrinting") Then
                        objRpt.DataDefinition.FormulaFields("DateofPrinting").Text = "'" & VB6.Format(Me.datePrinting.Value, gstrDateFormat) & "'"
                    End If
                End If
                rsSalesParameter.ResultSetClose()
            Else
                Exit Sub
            End If
        End If
        Select Case eventArgs.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_WINDOW
                strLable = mFrom & " OF " & NoOfLabels
                objRpt.DataDefinition.FormulaFields("PageCount").Text = "'" & strLable & "'"
                frmReportViewer.Show()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_PRINTER
                If optSelected.Checked = True Then
                    If ((Val(txtPageTo.Text) - Val(txtPageTo.Text)) + 1) >= NoOfLabels Then
                        intMaxLoop = NoOfLabels
                    Else
                        intMaxLoop = (((CDbl(txtPageTo.Text) - CDbl(txtpagefrom.Text))) + Val(txtpagefrom.Text))
                    End If
                Else
                    intMaxLoop = NoOfLabels
                End If
                For intLoopCounter = mFrom To intMaxLoop
                    strLable = intLoopCounter & " OF " & NoOfLabels
                    If intMaxLoop = NoOfLabels Then
                        If intLoopCounter = intMaxLoop Then
                            dblQuantity = mSalesQuantity - (intLoopCounter - 1) * mBoxQuantity
                            objRpt.DataDefinition.FormulaFields("BoxQuantity").Text = "'" & CStr(dblQuantity) & "'"
                        Else
                            objRpt.DataDefinition.FormulaFields("BoxQuantity").Text = "'" & CStr(mBoxQuantity) & "'"
                        End If
                    Else
                        objRpt.DataDefinition.FormulaFields("BoxQuantity").Text = "'" & CStr(mBoxQuantity) & "'"
                    End If
                    objRpt.DataDefinition.FormulaFields("PageCount").Text = "'" & strLable & "'"
                    frmReportViewer.SetReportDocument()
                    objRpt.PrintToPrinter(1, False, 0, 0)
                Next
        End Select
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdUnitCodeList_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUnitCodeList.Click
        Dim varHelp As Object
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        varHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT Unt_CodeID,unt_unitname FROM Gen_UnitMaster WHERE Unt_Status=1 and Unt_CodeID = '" & gstrUNITID & "'")
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(varHelp) <> -1 Then
            If varHelp(0) <> "0" Then
                txtUnitCode.Text = Trim(varHelp(0))
                txtUnitCode.Focus()
            Else
                MsgBox("No Record Available", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub optAll_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAll.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            If optAll.Checked = True Then
                txtpagefrom.Enabled = False : txtpagefrom.Enabled = False : txtpagefrom.Text = "" : txtpagefrom.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtPageTo.Enabled = False : txtPageTo.Enabled = False : txtPageTo.Text = "" : txtPageTo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Else
                txtPageTo.Enabled = True : txtPageTo.Enabled = True : txtPageTo.Text = "" : txtPageTo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtpagefrom.Enabled = True : txtpagefrom.Enabled = True : txtpagefrom.Text = "" : txtpagefrom.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            End If
            Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub optSelected_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSelected.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            If optSelected.Checked = True Then
                txtPageTo.Enabled = True : txtPageTo.Enabled = True : txtPageTo.Text = "" : txtPageTo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtpagefrom.Enabled = True : txtpagefrom.Enabled = True : txtpagefrom.Text = "" : txtpagefrom.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtpagefrom.Focus()
            Else
                txtpagefrom.Enabled = False : txtpagefrom.Enabled = False : txtpagefrom.Text = "" : txtpagefrom.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtPageTo.Enabled = False : txtPageTo.Enabled = False : txtPageTo.Text = "" : txtPageTo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            End If
            Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub txtbinQuantity_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtbinQuantity.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Errorhandler
        If KeyCode = System.Windows.Forms.Keys.Return And Shift = 0 Then datePrinting.Focus()
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtInvoice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtInvoice.KeyDown
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.Shift
        On Error GoTo Errorhandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then Call cmdHelpInvoice_Click(cmdHelpInvoice, New System.EventArgs())
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtInvoice_TextChanged(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles txtInvoice.TextChanged
        txtbinQuantity.Text = ""
        txtpagefrom.Text = ""
        txtPageTo.Text = ""
    End Sub
    Private Sub txtpagefrom_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtpagefrom.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{tab}")
        End If
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtPageTo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPageTo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{tab}")
        End If
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtUnitCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnitCode.TextChanged
        On Error GoTo ErrHandler
        txtInvoice.Text = ""
        txtbinQuantity.Text = ""
        txtpagefrom.Text = ""
        txtPageTo.Text = ""
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtUnitCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnitCode.Enter
        On Error GoTo ErrHandler
        'Selecting the text on focus
        With txtUnitCode
            .SelectionStart = 0 : .SelectionLength = Len(Trim(.Text))
        End With
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtUnitCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtUnitCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        'Show the help form when user pressed F1
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then Call cmdUnitCodeList_Click(cmdUnitCodeList, New System.EventArgs())
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtUnitCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtUnitCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{tab}")
            'Supressing ¬ ¤ ¦ » characters since these are being used as string delimiters
        ElseIf KeyAscii = 187 Or KeyAscii = 166 Or KeyAscii = 164 Or KeyAscii = 172 Or KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 96 Then
            KeyAscii = 0
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtUnitCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtUnitCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim strUnitDesc As String
        Dim mobjGLTrans As New prj_GLTransactions.cls_GLTransactions(gstrUNITID, GetServerDate)
        'Populate the details
        If Trim(txtUnitCode.Text) = "" Then GoTo EventExitSub
        strUnitDesc = mobjGLTrans.GetUnit(Trim(txtUnitCode.Text), ConnectionString:=gstrCONNECTIONSTRING)
        If CheckString(strUnitDesc) <> "Y" Then
            MsgBox(CheckString(strUnitDesc), MsgBoxStyle.Critical, ResolveResString(100))
            txtUnitCode.Text = ""
            Cancel = True
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub ctlFormHeader1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ctlFormHeader1.Click
        MsgBox("No Help Attached to This Form", MsgBoxStyle.Information, ResolveResString(100))
    End Sub
    Private Sub frmMKTTRN0017_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo Err_Handler
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0017_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo Err_Handler
        frmModules.NodeFontBold(Tag) = False
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0017_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlFormHeader1_ClickEvent(ctlFormHeader1, New System.EventArgs()) : Exit Sub
        ElseIf KeyCode = System.Windows.Forms.Keys.Escape And Shift = 0 Then
            Me.Close()
        End If
    End Sub
    Private Sub frmMKTTRN0017_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo Err_Handler
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FillLabelFromResFile(Me) 'To Fill label description from Resource file
        Call FitToClient(Me, fraInvoice, ctlFormHeader1, Cmdinvoice) 'To fit the form in the MDI
        datePrinting.Format = DateTimePickerFormat.Custom
        datePrinting.CustomFormat = gstrDateFormat
        optAll.Checked = True : optSelected.Checked = False
        gblnCancelUnload = False
        Me.datePrinting.Value = GetServerDate()
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0017_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo Err_Handler
        'REFRESH
        'Removing the form name from list
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        'Setting the corresponding node's tag
        frmModules.NodeFontBold(Tag) = False
        Me.Dispose()
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function ValidatebeforeSave() As Boolean
        Dim rsItemCount As ClsResultSetDB
        Dim rsBinQuantity As ClsResultSetDB
        Dim strErrMsg As String
        Dim strAccountCode As String
        Dim StrItemCode As String
        Dim strCustomerItemCode As String
        Dim ctlBlank As System.Windows.Forms.Control
        Dim lNo As Integer
        Dim blnInvalidData As Boolean
        Dim dblBinQuantity As Double
        Dim strsql As String
        On Error GoTo Err_Handler
        ValidatebeforeSave = False
        lNo = 1
        strErrMsg = ResolveResString(10059) & vbCrLf & vbCrLf
        If Len(Trim(txtUnitCode.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Location Code"
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = txtUnitCode
        End If
        If Len(Trim(txtInvoice.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & ResolveResString(60373)
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = txtInvoice
        Else
            strsql = "select * from saleschallan_dtl a,Sales_dtl b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND a.location_code = b.Location_code and a.Doc_no = b.doc_no and a.doc_no = " & Trim(txtInvoice.Text) & " and a.cancel_flag = 0 and a.bill_flag = 1"
            rsItemCount = New ClsResultSetDB
            rsItemCount.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsItemCount.GetNoRows > 1 Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & ". Selected invoice has more then one items associated with it "
                lNo = lNo + 1
                If ctlBlank Is Nothing Then ctlBlank = txtInvoice
            ElseIf rsItemCount.GetNoRows = 1 Then
                rsBinQuantity = New ClsResultSetDB
                rsItemCount.MoveFirst()
                strAccountCode = rsItemCount.GetValue("Account_code")
                StrItemCode = rsItemCount.GetValue("Item_code")
                strCustomerItemCode = rsItemCount.GetValue("Cust_Item_code")
                mSalesQuantity = rsItemCount.GetValue("Sales_Quantity")
                dblBinQuantity = Val(rsItemCount.GetValue("BinQuantity"))
                mBoxQuantity = dblBinQuantity
                If dblBinQuantity = 0 Then
                    blnInvalidData = True
                    strErrMsg = strErrMsg & vbCrLf & lNo & ". selected invoice has more then one items associated with it. "
                    lNo = lNo + 1
                    If ctlBlank Is Nothing Then ctlBlank = txtInvoice
                End If
            End If
            rsItemCount.ResultSetClose()
            rsItemCount = Nothing
        End If
        If Val(txtbinQuantity.Text) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & ". Bin Quantity Should be Greater then ZERO."
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = txtbinQuantity
        End If
        If optSelected.Checked = True Then
            If Len(Trim(txtpagefrom.Text)) = 0 Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & ". Page From "
                lNo = lNo + 1
                If ctlBlank Is Nothing Then ctlBlank = txtpagefrom
            ElseIf Val(Trim(txtpagefrom.Text)) = 0 Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & ". Enter Greater then ZERO."
                lNo = lNo + 1
                If ctlBlank Is Nothing Then ctlBlank = txtpagefrom
            End If
            If Len(Trim(txtPageTo.Text)) = 0 Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & ". Page To "
                lNo = lNo + 1
                If ctlBlank Is Nothing Then ctlBlank = txtPageTo
            ElseIf Val(Trim(txtPageTo.Text)) = 0 Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & ". Enter Greater then ZERO."
                lNo = lNo + 1
                If ctlBlank Is Nothing Then ctlBlank = txtPageTo
            End If
            If Val(txtPageTo.Text) < Val(txtpagefrom.Text) Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & ". Page To cannot be Greater then Page From"
                lNo = lNo + 1
                If ctlBlank Is Nothing Then ctlBlank = txtPageTo
            End If
        End If
        strErrMsg = VB.Left(strErrMsg, Len(strErrMsg) - 1)
        strErrMsg = strErrMsg & "."
        lNo = lNo + 1
        If blnInvalidData = True Then
            gblnCancelUnload = True
            Call MsgBox(strErrMsg, MsgBoxStyle.Information, ResolveResString(100))
            ctlBlank.Focus()
            Exit Function
        End If
        ValidatebeforeSave = True
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub RefreshValues()
        mBoxQuantity = 0
        mSalesQuantity = 0
    End Sub
    Private Sub datePrinting_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles datePrinting.KeyDown
        On Error GoTo Errorhandler
        If e.KeyCode = Keys.Return And e.Shift = 0 Then Cmdinvoice.Focus()
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtInvoice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoice.Leave
        Dim strsql As String
        Dim rsInvoice As ClsResultSetDB
        Dim strItem_code As String
        Dim strCustDrgno As String
        Dim strCustomer As String
        Dim boxQuantity As Double
        On Error GoTo Err_Handler
        If Len(txtInvoice.Text) = 0 Then Exit Sub
        strsql = "Select * from SalesChallan_dtl a,sales_dtl b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND  bill_flag = 1 and cancel_flag =0 and a.doc_no like '"
        strsql = strsql & Trim(txtInvoice.Text) & "%' and a.Location_code = '" & Trim(txtUnitCode.Text) & "' and "
        strsql = strsql & " a.Location_code = b.Location_code and a.Doc_no = b.Doc_no"
        rsInvoice = New ClsResultSetDB
        rsInvoice.GetResult(strsql)
        If rsInvoice.GetNoRows > 0 Then
            If rsInvoice.GetNoRows > 1 Then
                MsgBox("selected invoice has more then one items associated with it", MsgBoxStyle.Information, ResolveResString(100))
                txtInvoice.Text = ""
                txtInvoice.Focus()
            Else
                strCustomer = rsInvoice.GetValue("account_code")
                strItem_code = rsInvoice.GetValue("Item_code")
                strCustDrgno = rsInvoice.GetValue("Cust_Item_code")
                mSalesQuantity = rsInvoice.GetValue("Sales_Quantity")
                boxQuantity = Val(rsInvoice.GetValue("BinQuantity"))
                rsInvoice.ResultSetClose()
                rsInvoice = Nothing
                If boxQuantity > 0 Then
                    txtbinQuantity.Text = CStr(boxQuantity)
                    mBoxQuantity = boxQuantity
                Else
                    MsgBox("No Bin Quantity is Defined for this item " & strItem_code & " in Customer Item Master. ", MsgBoxStyle.Information, ResolveResString(100))
                    txtInvoice.Text = ""
                    txtInvoice.Focus()
                End If
                If optSelected.Checked = True Then
                    Me.txtpagefrom.Focus()
                Else
                    Me.txtbinQuantity.Focus()
                End If
            End If
        Else
            txtInvoice.Text = ""
            txtInvoice.Focus()
            Exit Sub
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtInvoice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtInvoice.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        On Error GoTo Err_Handler
        If KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{tab}")
        End If
        If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
            KeyAscii = KeyAscii
        Else
            KeyAscii = 0
        End If
        If KeyAscii = 39 Then
            e.Handled = True
        End If
        If KeyAscii = 0 Then
            e.Handled = True
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
End Class