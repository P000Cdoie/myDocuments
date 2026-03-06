Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Friend Class frmMKTTRN0039
	Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Designs Ltd. All rights reserverd.
	' File Name                :       FRMMKTTRN0039.frm
	' Purpose                  :       Used to Print SRV57F4 Bin card
	' Created By               :       Brij Bohara
	' Created On               :       15 Dec, 2004
	' Revision History         :       -
	'-----------------------------------------------------------------------
	' Revision Date            :       08/05/2006
	' Revision By              :       Davinder Singh
	' Issue ID                 :       17736
	' Revision History         :       1) Replace the table 'PrintedSRV_Dtl' with tables
	'                                     'mkt_57F4Challan_Hdr' and 'mkt_57F4Challan_Dtl'
	'                                  2) Pick the BinQuantity from the 'mkt_57F4Challan_Dtl'
    '                                     table instead of picking it from Custitem_mst
    'MODIFIED BY AJAY SHUKLA ON 10/MAY/2011 FOR MULTIUNIT CHANGE
	'-----------------------------------------------------------------------
	
	Dim mintFormIndex As Double 'To store the form Index
	Dim mBoxQuantity As Double 'To store the Box quantity
	Dim mSalesQuantity As Double 'To store the Sales Quantity
	
	Private Sub cmdHelpInvoice_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelpInvoice.Click
		'----------------------------------------------------
		'Author              -      Brij Bohara
		'Create Date         -      15/12/2004
		'Arguments           -      None
		'Return Value        -      None
		'Purpose             -      To Display Help Form
		'----------------------------------------------------
		Dim varHelp As Object 'To trap the help result
		
		On Error GoTo ErrHandler
		Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen,  , System.Windows.Forms.Cursors.WaitCursor) 'To change the mouse pointer
		
		
        varHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select A.doc_no,B.Invoice_no,B.Invoice_date from mkt_57F4Challan_Hdr A,Grn_Hdr B where A.Grin_No=B.doc_no and a.unit_code = b.unit_code AND A.Doc_no like '" & Trim(txtInvoice.Text) & "%' and a.unit_code='" & gstrUNITID & "'")

        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default) 'To change the mouse pointer
        If UBound(varHelp) <> -1 Then
            If varHelp(0) <> "0" Then
                txtInvoice.Text = Trim(varHelp(0))
                txtInvoice_Leave(txtInvoice, Nothing)
            Else
                MsgBox("No Record Available", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub Cmdinvoice_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCfraRepCmd.ButtonClickEventArgs) Handles Cmdinvoice.ButtonClick
        '----------------------------------------------------
        'Author              -      Brij Bohara
        'Create Date         -      15/12/2004
        'Arguments           -      None
        'Return Value        -      None
        'Purpose             -      To call Button Click Events
        '----------------------------------------------------
        Dim NoOfLabels As Integer 'To store number of labels
        Dim mFrom As Short 'To store From
        Dim mTo As Short 'To store To
        Dim intLoopCounter As Short 'To count the iterations
        Dim intMaxLoop As Short 'To store maximum number of iteration of loop
        Dim strLable As String 'To store Label Sttring
        Dim strDate As String 'To store date
        Dim strLocation As String 'To store location
        Dim strInvNo As String
        Dim strDINo As String 'To store DI numbner which was saved during SRV
        Dim strsql As String 'To store query string
        Dim strReportFileName As String 'To store Report file name
        Dim strReportFilePath As String 'To store report file path
        Dim rsSRVDtl As ClsResultSetDB 'To retrive the values from sales challan_dtl table
        Dim rsSalesParameter As ClsResultSetDB 'To retrive the vales from sales_parameter table

        Dim RDoc As ReportDocument
        Dim CRViewer As New eMProCrystalReportViewer
        RDoc = CRViewer.GetReportDocument()

        rsSRVDtl = New ClsResultSetDB
        rsSalesParameter = New ClsResultSetDB
        On Error GoTo ErrHandler
        If e.Button <> UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE Then
            If e.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_FILE Then
                MsgBox("Export is not available", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor) 'To change the mouse pointer
            Call RefreshValues()
            If ValidatebeforeSave() = True Then
                rsSRVDtl.GetResult("Select B.Invoice_No,B.Invoice_Date,isnull(D.KanbanNo,'') AS KanbanNo,isnull(D.UNLOC,'') AS UNLOC,B.Vendor_Code from Mkt_57f4challan_Hdr A inner join Grn_Hdr B on A.Grin_No = B.Doc_No and A.unit_code = B.unit_code left outer Join Mkt_57F4ChallanKanBan_Dtl C on A.Doc_No = C.Doc_No and A.unit_code = C.unit_code left outer join Mkt_EnagareDtl D on C.Kanban_No = D.KanbanNo  and c.unit_code = d.unit_code where A.doc_no = '" & Trim(txtInvoice.Text) & "' and a.unit_code='" & gstrUNITID & "'")
                If rsSRVDtl.GetNoRows > 0 Then
                    rsSRVDtl.MoveFirst()
                    strInvNo = rsSRVDtl.GetValue("Invoice_No")
                    strDate = VB6.Format(rsSRVDtl.GetValue("Invoice_date"), gstrDateFormat)
                    strDINo = rsSRVDtl.GetValue("KanbanNo")
                    strLocation = rsSRVDtl.GetValue("UNLOC")
                End If

                strsql = "{mkt_57F4Challan_hdr.Doc_No} =" & Trim(txtInvoice.Text) & _
                " And {CustItem_Mst.Account_Code}='" & rsSRVDtl.GetValue("Vendor_Code") & "'" & _
                " And {mkt_57F4Challan_hdr.UNIT_CODE}='" & gstrUNITID & "'"

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

                With RDoc

                    strReportFilePath = My.Application.Info.DirectoryPath & "\Reports\"
                    rsSalesParameter.GetResult("Select SRVBinCardFileName from Sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'")
                    rsSalesParameter.MoveFirst()
                    If Len(Trim(rsSalesParameter.GetValue("SRVBinCardFileName"))) > 0 Then
                        strReportFileName = rsSalesParameter.GetValue("SRVBinCardFileName")
                        If InStr(1, strReportFileName, ".") > 0 Then
                            strReportFileName = Mid(strReportFileName, 1, InStr(1, strReportFileName, ".") - 1)
                        End If
                    Else
                        MsgBox("No Report File Name Defined in Sales_Parameter Table.", MsgBoxStyle.Information, "empower")
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default) 'To change the mouse pointer
                        Exit Sub
                    End If
                    .Load(strReportFilePath & strReportFileName & ".rpt")
                    .RecordSelectionFormula = strsql
                    .DataDefinition.FormulaFields("InvoiceNo").Text = "'" + Int(CDbl(txtInvoice.Text)).ToString + "'"
                    .DataDefinition.FormulaFields("InvoiceDate").Text = "'" & strDate & "'"
                    .DataDefinition.FormulaFields("SRVLocation").Text = "'" & strLocation & "'"
                    .DataDefinition.FormulaFields("SRVDINO").Text = "'" & strDINo & "'"
                    .DataDefinition.FormulaFields("NoofBin").Text = "'" & NoOfLabels & "'"
                    .DataDefinition.FormulaFields("BoxQuantity").Text = "'" & mBoxQuantity & "'"

                End With
                CRViewer.ReportHeader = Me.ctlFormHeader1.HeaderString()

            Else
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default) 'To change the mouse pointer
                Exit Sub
            End If
        End If


        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_WINDOW
                strLable = mFrom & " OF " & NoOfLabels
                RDoc.DataDefinition.FormulaFields("PageCount").Text = "'" & strLable & "'"
                CRViewer.Show()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_PRINTER
                CRViewer.SetReportDocument()
                If optSelected.Checked = True Then
                    If ((Val(txtPageTo.Text) - Val(txtPageTo.Text)) + 1) >= NoOfLabels Then
                        intMaxLoop = NoOfLabels
                    Else
                        intMaxLoop = (((CDbl(txtPageTo.Text) - CDbl(txtpagefrom.Text))) + Val(txtpagefrom.Text))
                    End If
                Else
                    intMaxLoop = NoOfLabels
                End If
                For intLoopCounter = mFrom To intMaxLoop '   MsgBox intLoopCounter & " OF " & NoOfLabels
                    strLable = intLoopCounter & " OF " & NoOfLabels
                    RDoc.DataDefinition.FormulaFields("PageCount").Text = "'" & strLable & "'"
                    RDoc.PrintToPrinter(1, False, 0, 0)
                Next
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
                Exit Sub
        End Select
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default) 'To change the mouse pointer
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default) 'To change the mouse pointer
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdUnitCodeList_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUnitCodeList.Click
        '----------------------------------------------------
        'Author              -      Brij Bohara
        'Create Date         -      15/12/2004
        'Arguments           -      None
        'Return Value        -      None
        'Purpose             -      To Show Help Form
        '----------------------------------------------------
        Dim varHelp As Object
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        varHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT Unt_CodeID,unt_unitname FROM Gen_UnitMaster WHERE Unt_Status=1 and Unt_CodeID='" & gstrUNITID & "'")
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(varHelp) <> -1 Then
            If varHelp(0) <> "0" Then
                txtUnitCode.Text = Trim(varHelp(0))
                txtInvoice.Focus() 'Set focus to next control
            Else
                MsgBox("No Record Available", MsgBoxStyle.Information, "empower")
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
	
	Private Sub frmMKTTRN0039_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		'----------------------------------------------------
		'Author              -      Davinder Singh
		'Create Date         -      08/05/2006
		'Arguments           -      Keycode of the key pressed,shift
		'Return Value        -      None
		'Purpose             -
		'----------------------------------------------------
		On Error GoTo ErrHandler
		If KeyCode = System.Windows.Forms.Keys.Escape And Shift = 0 Then Me.Close()
ErrHandler: 
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
	End Sub
	
	Private Sub optAll_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAll.CheckedChanged
        If eventSender.Checked Then
            '----------------------------------------------------
            'Author              -      Brij Bohara
            'Create Date         -      15/12/2004
            'Arguments           -      None
            'Return Value        -      None
            'Purpose             -      To Enable/Disable Page from & Page To Text Boxes
            '----------------------------------------------------
            On Error GoTo ErrHandler
            If optAll.Checked = True Then
                txtpagefrom.Enabled = False : txtpagefrom.Enabled = False : txtpagefrom.Text = "" : txtpagefrom.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtPageTo.Enabled = False : txtPageTo.Enabled = False : txtPageTo.Text = "" : txtPageTo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Else
                txtpagefrom.Enabled = True : txtpagefrom.Enabled = True : txtpagefrom.Text = "" : txtpagefrom.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtPageTo.Enabled = True : txtPageTo.Enabled = True : txtPageTo.Text = "" : txtPageTo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            End If
            Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub

    Private Sub optSelected_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSelected.CheckedChanged
        If eventSender.Checked Then
            '----------------------------------------------------
            'Author              -      Brij Bohara
            'Create Date         -      15/12/2004
            'Arguments           -      None
            'Return Value        -      None
            'Purpose             -      To Enable/Disable Page from & Page To Text Boxes
            '----------------------------------------------------
            On Error GoTo ErrHandler
            If optSelected.Checked = True Then
                txtpagefrom.Enabled = True : txtpagefrom.Enabled = True : txtpagefrom.Text = "" : txtpagefrom.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtPageTo.Enabled = True : txtPageTo.Enabled = True : txtPageTo.Text = "" : txtPageTo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            Else
                txtpagefrom.Enabled = False : txtpagefrom.Enabled = False : txtpagefrom.Text = "" : txtpagefrom.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtPageTo.Enabled = False : txtPageTo.Enabled = False : txtPageTo.Text = "" : txtPageTo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            End If
            Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
	
	
	Private Sub txtbinQuantity_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtbinQuantity.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'----------------------------------------------------
		'Author              -      Davinder Singh
		'Create Date         -      08/05/2004
		'Arguments           -      None
		'Return Value        -      None
		'Purpose             -      To set the focus on the Main Cmd buttons
		'----------------------------------------------------
		On Error GoTo Err_Handler
		If KeyAscii = 13 And optAll.Checked Then cmdInvoice.Focus()
		GoTo EventExitSub 'This is to avoid the execution of the error handler
Err_Handler: 
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	


    Private Sub txtUnitCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUnitCode.GotFocus
        '----------------------------------------------------
        'Author              -      Brij Bohara
        'Create Date         -      15/12/2004
        'Arguments           -      None
        'Return Value        -      None
        'Purpose             -      To Selected Set Focus
        '----------------------------------------------------
        On Error GoTo ErrHandler
        'Selecting the text on focus
        With txtUnitCode
            .SelectionStart = 0 : .SelectionLength = Len(Trim(.Text))
        End With
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub
	
	Private Sub txtUnitCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnitCode.TextChanged
        '----------------------------------------------------
        'Author              -      Brij Bohara
        'Create Date         -      15/12/2004
        'Arguments           -      None
        'Return Value        -      None
        'Purpose             -      To cleare Date
        '----------------------------------------------------
        On Error GoTo ErrHandler
        txtInvoice.Text = ""
        txtbinQuantity.Text = ""
        txtpagefrom.Text = ""
        txtPageTo.Text = ""
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
		Private Sub txtUnitCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtUnitCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '----------------------------------------------------
        'Author              -      Brij Bohara
        'Create Date         -      15/12/2004
        'Arguments           -      None
        'Return Value        -      None
        'Purpose             -      To Call help on F1 Press
        '----------------------------------------------------
        On Error GoTo ErrHandler
        'If Ctrl/Alt/Shift is also pressed
        If Shift <> 0 Then Exit Sub
        'Show the help form when user pressed F1
        If KeyCode = System.Windows.Forms.Keys.F1 Then Call cmdUnitCodeList_Click(cmdUnitCodeList, New System.EventArgs())
        If KeyCode = Keys.Enter Then System.Windows.Forms.SendKeys.Send("{TAB}")
        Exit Sub 'This is to avoid the execution of the error handler

ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub
	Private Sub txtUnitCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtUnitCode.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'----------------------------------------------------
		'Author              -      Brij Bohara
		'Create Date         -      15/12/2004
		'Arguments           -      None
		'Return Value        -      None
		'Purpose             -      To Set Focus on Next Control
		'----------------------------------------------------
		On Error GoTo ErrHandler
		KeyAscii = Asc(UCase(Chr(KeyAscii)))
		If KeyAscii = 13 Then
			System.Windows.Forms.SendKeys.Send("{TAB}")
			'Supressing ¬ ¤ ¦ » characters since these are being used as string delimiters
		ElseIf KeyAscii = 187 Or KeyAscii = 166 Or KeyAscii = 164 Or KeyAscii = 172 Or KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 96 Then 
			KeyAscii = 0
		End If
		GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler: 
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtUnitCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtUnitCode.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		'----------------------------------------------------
		'Author              -      Brij Bohara
		'Create Date         -      15/12/2004
		'Arguments           -      None
		'Return Value        -      None
		'Purpose             -      To Validate selected unit code
		'----------------------------------------------------
		On Error GoTo ErrHandler
		Dim strUnitDesc As String 'To store the unit Name
        Dim mobjGLTrans As New prj_GLTransactions.cls_GLTransactions(gstrUNITID, GetServerDate) 'To retrieve the unit name
		'Populate the details
		If Trim(txtUnitCode.Text) = "" Then GoTo EventExitSub
		strUnitDesc = mobjGLTrans.GetUnit(Trim(txtUnitCode.Text), ConnectionString:=gstrCONNECTIONSTRING)
		If CheckString(strUnitDesc) <> "Y" Then
			MsgBox(CheckString(strUnitDesc), MsgBoxStyle.Critical, "eMpower")
			txtUnitCode.Text = ""
			Cancel = True
		End If
		GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler: 
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
		
EventExitSub: 
		eventArgs.Cancel = Cancel
	End Sub
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        '----------------------------------------------------
        'Author              -      Brij Bohara
        'Create Date         -      15/12/2004
        'Arguments           -      None
        'Return Value        -      None
        'Purpose             -      To Call Empower help
        '----------------------------------------------------
        On Error GoTo ErrHandler
        MsgBox("No Help Attached to This Form", MsgBoxStyle.Information, "empower")
        Exit Sub 'To avoid the error handler execution
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub
  
	Private Sub frmMKTTRN0039_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '----------------------------------------------------
        'Author              -      Brij Bohara
        'Create Date         -      15/12/2004
        'Arguments           -      None
        'Return Value        -      None
        'Purpose             -      To intialise required
        '----------------------------------------------------
        On Error GoTo Err_Handler
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub 'To avoid the error handler execution

Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub frmMKTTRN0039_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '----------------------------------------------------
        'Author              -      Brij Bohara
        'Create Date         -      15/12/2004
        'Arguments           -      None
        'Return Value        -      None
        'Purpose             -      To released Values
        '----------------------------------------------------
        On Error GoTo Err_Handler
        frmModules.NodeFontBold(Tag) = False
        Exit Sub 'To avoid the error handler execution

Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
	Private Sub frmMKTTRN0039_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		'----------------------------------------------------
		'Author              -      Brij Bohara
		'Create Date         -      15/12/2004
		'Arguments           -      None
		'Return Value        -      None
		'Purpose             -      To display empower help on click of F4 in empower
		'----------------------------------------------------
		On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then Call ctlFormHeader1_Click(ctlFormHeader1, New System.EventArgs())
		Exit Sub 'To avoid the error handler execution
		
ErrHandler: 
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
		
	End Sub
	Private Sub frmMKTTRN0039_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'----------------------------------------------------
		'Author              -      Brij Bohara
		'Create Date         -      15/12/2004
		'Arguments           -      None
		'Return Value        -      None
		'Purpose             -      To initialise required data
		'----------------------------------------------------
		On Error GoTo Err_Handler
		mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FitToClient(Me, fraInvoice, ctlFormHeader1, Cmdinvoice) 'To fit the form in the MDI
		optAll.Checked = True : optSelected.Checked = False
        gblnCancelUnload = False
        Exit Sub 'To avoid the error handler execution
Err_Handler: 
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
	End Sub
	Private Sub frmMKTTRN0039_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'----------------------------------------------------
		'Author              -      Brij Bohara
		'Create Date         -      15/12/2004
		'Arguments           -      None
		'Return Value        -      None
		'Purpose             -      To Release memory
		'----------------------------------------------------
		On Error GoTo Err_Handler
		mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex 'Removing the form name from list
		frmModules.NodeFontBold(Tag) = False 'Setting the corresponding node's tag
        Me.Dispose() 'Releasing the form reference
		Exit Sub 'To avoid the error handler execution
Err_Handler: 
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

	Public Function ValidatebeforeSave() As Boolean
		'----------------------------------------------------
		'Author              -      Brij Bohara
		'Create Date         -      15/12/2004
		'Arguments           -      None
		'Return Value        -      Boolean (True/False)
		'Purpose             -      To Check Valid Feilds Value before saving
		'----------------------------------------------------
		Dim rsItemCount As ClsResultSetDB 'To count records
		Dim strErrMsg As String 'To trap error message
		Dim ctlBlank As System.Windows.Forms.Control 'To trap the control name
		Dim lNo As Integer 'To count line number
		Dim blnInvalidData As Boolean 'To store validation result
		Dim dblBinQuantity As Double 'To store Bin quantity
		Dim strsql As String 'To store SQL statement
		
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
            strsql = "select isnull(BinQuantity,1) as BinQuantity,Invoice_Qty from mkt_57F4Challan_dtl where Doc_no =  '" & Trim(txtInvoice.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'"
			rsItemCount = New ClsResultSetDB
			rsItemCount.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
			If rsItemCount.GetNoRows > 1 Then
				blnInvalidData = True
				strErrMsg = strErrMsg & vbCrLf & lNo & ". select invoice has more then one items. "
				lNo = lNo + 1
				If ctlBlank Is Nothing Then ctlBlank = txtInvoice
			ElseIf rsItemCount.GetNoRows = 1 Then 
                dblBinQuantity = Val(rsItemCount.GetValue("BinQuantity"))
				mBoxQuantity = dblBinQuantity
                mSalesQuantity = Val(rsItemCount.GetValue("Invoice_Qty"))
			End If
		End If
        If Val(txtbinQuantity.Text) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & ". Bin Quantity Should be Greater than ZERO."
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
				strErrMsg = strErrMsg & vbCrLf & lNo & ". Enter Greater than ZERO."
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
				strErrMsg = strErrMsg & vbCrLf & lNo & ". Enter Greater than ZERO."
				lNo = lNo + 1
				If ctlBlank Is Nothing Then ctlBlank = txtPageTo
			End If
			If Val(txtPageTo.Text) < Val(txtpagefrom.Text) Then
				blnInvalidData = True
				strErrMsg = strErrMsg & vbCrLf & lNo & ". Page To cannot be Greater than Page From"
				lNo = lNo + 1
				If ctlBlank Is Nothing Then ctlBlank = txtPageTo
			End If
		End If
		strErrMsg = VB.Left(strErrMsg, Len(strErrMsg) - 1)
		strErrMsg = strErrMsg & "."
		lNo = lNo + 1
		If blnInvalidData = True Then
			gblnCancelUnload = True
            Call MsgBox(strErrMsg, MsgBoxStyle.Information, "Error")
			ctlBlank.Focus()
			Exit Function
		End If
		ValidatebeforeSave = True
		Exit Function 'To avoid error handler
		
Err_Handler: 
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
	End Function
	Public Sub RefreshValues()
		'----------------------------------------------------
		'Author              -      Brij Bohara
		'Create Date         -      15/12/2004
		'Arguments           -      None
		'Return Value        -      None
		'Purpose             -      To intialise Values
		'----------------------------------------------------
		On Error GoTo ErrHandler
		mBoxQuantity = 0
		mSalesQuantity = 0
		Exit Sub
		
ErrHandler: 
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub


    Private Sub txtInvoice_TextChanged(ByVal Sender As Object, ByVal e As System.EventArgs) Handles txtInvoice.TextChanged
        '----------------------------------------------------
        'Author              -      Brij Bohara
        'Create Date         -      15/12/2004
        'Arguments           -      None
        'Return Value        -      None
        'Purpose             -      To Clear Related Data
        '----------------------------------------------------
        On Error GoTo Err_Handler
        txtbinQuantity.Text = ""
        txtpagefrom.Text = ""
        txtPageTo.Text = ""
        Exit Sub      'This is to avoid the execution of the error handler
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub

    Private Sub txtInvoice_KeyDown1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtInvoice.KeyDown
        '----------------------------------------------------
        'Author              -      Brij Bohara
        'Create Date         -      15/12/2004
        'Arguments           -      None
        'Return Value        -      None
        'Purpose             -      To show help and navigation
        '----------------------------------------------------
        If e.Shift <> 0 Then Exit Sub
        If e.KeyCode = Keys.F1 Then Call cmdHelpInvoice_Click(cmdHelpInvoice, New System.EventArgs())

        If e.KeyCode = Keys.Enter Then System.Windows.Forms.SendKeys.Send("{TAB}")
        Exit Sub      'This is to avoid the execution of the error handler

Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtInvoice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtInvoice.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        '----------------------------------------------------
        'Author              -      Brij Bohara
        'Create Date         -      15/12/2004
        'Arguments           -      None
        'Return Value        -      None
        'Purpose             -      To set focus on next control
        '----------------------------------------------------
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
        Exit Sub             'To avoid the error handler execution
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub

    Private Sub txtInvoice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoice.Leave
        '----------------------------------------------------
        'Author              -      Brij Bohara
        'Create Date         -      15/12/2004
        'Arguments           -      None
        'Return Value        -      None
        'Purpose             -      To Validate invoice No
        '----------------------------------------------------
        Dim strSQL As String               'To store the query
        Dim rsInvoice As ClsResultSetDB    'To retrieve the Invoice Number

        On Error GoTo Err_Handler
        If Len(txtInvoice.Text) = 0 Then Exit Sub

        If IsNumeric(txtInvoice.Text) = False Then
            MsgBox("Enter a Valid Invoice No.", vbInformation, "eMpro")
            txtInvoice.Text = ""
            txtInvoice.Focus()
            Exit Sub
        End If

        strSQL = "SELECT isnull(BinQuantity,1) as BinQuantity, Invoice_Qty From mkt_57F4Challan_dtl  where doc_no ='" & Trim(txtInvoice.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'"
        rsInvoice = New ClsResultSetDB
        rsInvoice.GetResult(strSQL)
        If rsInvoice.GetNoRows > 0 Then
            If rsInvoice.GetNoRows > 1 Then
                MsgBox("selected SRV has more than one items.", vbInformation, "empower")
                txtInvoice.Text = ""
                txtInvoice.Focus()
            Else
                txtbinQuantity.Text = Val(rsInvoice.GetValue("BinQuantity"))
                mBoxQuantity = Val(txtbinQuantity.Text)
                mSalesQuantity = Val(rsInvoice.GetValue("Invoice_Qty"))
            End If
        Else
            MsgBox("Entered Invoice No. Is Not Valid", vbInformation + vbOKOnly, ResolveResString(100))
            txtInvoice.Text = ""
            txtInvoice.Focus()
            Exit Sub
        End If

        If optSelected.Checked = True Then
            txtpagefrom.Focus()
        Else
            txtbinQuantity.Focus()
        End If

        Exit Sub    'This is to avoid the execution of the error handler
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub
End Class