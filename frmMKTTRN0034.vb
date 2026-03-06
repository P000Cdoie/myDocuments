Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.IO

Friend Class frmMKTTRN0034
	Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
	' File Name         :   FRMMKTTRN0034.frm
	' Function          :   Used for Range Printing of Invoice
	' Created By        :   Arshad Ali
	' Created On        :   31 May, 2004
	'===================================================================================
	'Revision History
	'-------------------------------------------------------
	'06/07/2004
	'Query Replaced By Arshad ali to include Captive invoice
	'-------------------------------------------------------
	'12/07/2004
	'Condition Added by Arshad to include captive invoice
	'14/10/2004 Add Captive Invoice Range Printing By Rajani Kant
	'-------------------------------------------------------
	'21/11/2004
	'Check box Added by Brij Bohara for making Bar Code Printing Optional
	'-------------------------------------------------------
	'Revised By                 -  Davinder Singh
	'Revision Date              -  20/07/2005
	'Revision History           -  Invoice printing was not in sequence and this problem is solved by filling the
	'                              List View by placing the 'order by' in the query
	'---------------------------------------------------------------------------------------------------------
	'Revised By                 -  Ashutosh Verma
	'Revision Date              -  22/12/2005
	'Revision History           -  Issue Id:16625, for printing Bar code on Invoice at SUNVAC according to check box for Bar code printing.
	'---------------------------------------------------------------------------------------------------------
	'Revised By        : Ashutosh on 07-01-2006
    'Revision History  : Issue Id:16780, To show Captive invoice on New Invoice Format (SUNVAC).
    'Modified by Sameer Srivastava on 2011-May-20
    '   Modified to support MultiUnit functionality
	'=============================================================
    'Modified By Roshan Singh on 19 Dec 2011 for multiUnit change management    

	Dim mintFormIndex As Double
	Dim mCtlHdrInvoiceNo As System.Windows.Forms.ColumnHeader
	Dim mCtlHdrInvoiceDate As System.Windows.Forms.ColumnHeader
	Dim mCtlHdrBinQty As System.Windows.Forms.ColumnHeader
	Dim mlvwInvoice As System.Windows.Forms.ListViewItem
	Dim mStrFirstFileName As String
    Dim objInvoicePrint As New prj_InvoicePrinting.clsInvoicePrinting(gstrDateFormat)

    Private Sub chkRemoval_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkRemoval.CheckStateChanged
        If chkRemoval.CheckState Then
            dtpRemoval.Enabled = True
            dtpRemovalTime.Enabled = True
        Else
            dtpRemoval.Enabled = False
            dtpRemovalTime.Enabled = False
        End If
    End Sub
	
	Private Sub cmdclose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdclose.Click
		FraInvoicePreview.Visible = False
		lvwInvoice.Focus()
	End Sub
	
	Private Sub cmdHelpInvoice_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelpInvoice.Click
		'----------------------------------------------------
		'Author              - Arshad Ali
		'Create Date         - 25/05/2003
		'Arguments           - None
		'Return Value        - None
		'Function            - To Show Help Form
		'----------------------------------------------------
		Dim varHelp As Object
		On Error GoTo ErrHandler
		Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen,  , System.Windows.Forms.Cursors.WaitCursor)
        varHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT Doc_no," & DateColumnNameInShowList("invoice_date") & " as invoice_date,cust_name FROM Saleschallan_dtl WHERE UNIT_CODE = '" & gstrUNITID & "' AND invoice_type='" & lblInvoiceType.Text & "' and sub_category='" & lblInvoiceSubType.Text & "' and Bill_flag =1 and Cancel_flag =0 and Doc_no like '" & Trim(txtInvoice.Text) & "%' order by doc_no")
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(varHelp) <> -1 Then
            If varHelp(0) <> "0" Then
                txtInvoice.Text = Trim(varHelp(0))
                txtInvoice.Focus()
            Else
                MsgBox("No Record Available", MsgBoxStyle.Information, "empower")
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mp_connection)
    End Sub

    Private Sub cmdHelpInvoice2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelpInvoice2.Click
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Show Help Form
        '----------------------------------------------------
        Dim varHelp As Object
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        varHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT Doc_no," & DateColumnNameInShowList("invoice_date") & " as invoice_date,cust_name FROM Saleschallan_dtl WHERE UNIT_CODE = '" & gstrUNITID & "' AND invoice_type='" & lblInvoiceType.Text & "' and sub_category='" & lblInvoiceSubType.Text & "' and Bill_flag =1 and Cancel_flag =0 and doc_no > " & Val(txtInvoice.Text) & " and Doc_no like '" & Trim(txtInvoiceTo.Text) & "%' order by doc_no")
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(varHelp) <> -1 Then
            If varHelp(0) <> "0" Then
                txtInvoiceTo.Text = Trim(varHelp(0))
                txtInvoiceTo.Focus()
            Else
                MsgBox("No Record Available", MsgBoxStyle.Information, "empower")
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mp_connection)
    End Sub

    Private Sub Cmdinvoice_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCfraRepCmd.ButtonClickEventArgs) Handles Cmdinvoice.ButtonClick
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To call Button Click Events
        '----------------------------------------------------
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim intCount As Integer
        On Error GoTo ErrHandler
        If e.Button <> UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE Then
            If e.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_FILE Then
                MsgBox("Export is not available", MsgBoxStyle.Information, "empower")
                Exit Sub
            End If
        End If
        Dim intTotalselected As Short
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_WINDOW
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                intTotalselected = 0
                For intCount = 0 To lvwInvoice.Items.Count - 1
                    If lvwInvoice.Items.Item(intCount).Checked = True Then
                        intTotalselected = intTotalselected + 1
                    End If
                Next
                If intTotalselected > 1 Then
                    MsgBox("Please select only one Invoice at once to preview.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "empower")
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    Exit Sub
                End If
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                If ValidatebeforeSave() Then
                    If CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE = '" & gstrUNITID & "'")) Then
                        mStrFirstFileName = ""
                        'Generatin all text files
                        Call PrintingInvoice()
                        'Previewing the first one on the Screen
                        If mStrFirstFileName <> "" Then
                            rtbInvoicePreview.LoadFile(mStrFirstFileName, RichTextBoxStreamType.PlainText)
                            rtbInvoicePreview.BackColor = System.Drawing.Color.White
                            cmdPrint.Image = My.Resources.ico231.ToBitmap
                            cmdClose.Image = My.Resources.ico217.ToBitmap
                            FraInvoicePreview.Visible = True
                            FraInvoicePreview.Enabled = True
                            FraInvoicePreview.BringToFront()

                            FraInvoicePreview.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 1050)
                            FraInvoicePreview.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) - 400)
                            FraInvoicePreview.Left = VB6.TwipsToPixelsX(100)
                            FraInvoicePreview.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(ctlFormHeader1.Height) - 50)

                            rtbInvoicePreview.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(FraInvoicePreview.Height) - 1000)
                            rtbInvoicePreview.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(FraInvoicePreview.Width) - 200)
                            rtbInvoicePreview.Left = VB6.TwipsToPixelsX(100)
                            rtbInvoicePreview.Top = VB6.TwipsToPixelsY(900)
                            rtbInvoicePreview.RightMargin = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(rtbInvoicePreview.Width) + 5000)

                            shpInvoice.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(FraInvoicePreview.Width) - VB6.PixelsToTwipsX(shpInvoice.Width)) / 2)
                            cmdPrint.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(shpInvoice.Left) + 100)
                            cmdClose.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(cmdPrint.Left) + VB6.PixelsToTwipsX(cmdPrint.Width) + 100)

                            cmdPrint.Enabled = True : cmdClose.Enabled = True
                            FraInvoicePreview.Enabled = True : rtbInvoicePreview.Enabled = True : rtbInvoicePreview.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            ReplaceJunkCharacters()
                            rtbInvoicePreview.Focus()
                        End If
                    End If
                End If
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_PRINTER
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                If ValidatebeforeSave() = True Then
                    If CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE = '" & gstrUNITID & "'")) Then
                        mStrFirstFileName = ""
                        'Generatin all text files and send to printer them
                        Call PrintingInvoiceToPrinter()
                    End If
                End If
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Select
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mp_connection)
    End Sub

    Private Sub cmdPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrint.Click
        On Error GoTo ErrHandler
        Call Cmdinvoice_ButtonClick(Cmdinvoice, New UCActXCtl.UCfraRepCmd.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_PRINTER))
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub cmdUnitCodeList_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUnitCodeList.Click
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Show Help Form
        '----------------------------------------------------
        Dim varHelp As Object
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        varHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT s.Location_Code,l.Description FROM Location_mst l,SaleConf s WHERE s.Location_Code = l.Location_Code AND s.unit_code = l.unit_code and s.unit_code = '" & gstrUNITID & "'")

        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(varHelp) <> -1 Then
            If varHelp(0) <> "0" Then
                txtUnitCode.Text = Trim(varHelp(0))
                txtUnitCode.Focus()
            Else
                MsgBox("No Record Available", MsgBoxStyle.Information, "empower")
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mp_connection)
    End Sub


	Private Sub frmMKTTRN0034_Paint(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
		txtInvoice.Enabled = True
        If txtInvoice.Enabled Then
        End If
	End Sub
	
	Private Sub lvwInvoice_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lvwInvoice.Click
		
	End Sub
	
    Private Sub optAll_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptAll.CheckedChanged
        If eventSender.Checked Then
            '----------------------------------------------------
            'Author              - Arshad Ali
            'Create Date         - 25/05/2003
            'Arguments           - None
            'Return Value        - None
            '----------------------------------------------------
            On Error GoTo ErrHandler
            Dim intCount As Integer
            If OptAll.Checked = True Then
                For intCount = 0 To lvwInvoice.Items.Count - 1
                    If lvwInvoice.Items.Item(intCount).Checked = False Then
                        lvwInvoice.Items.Item(intCount).Checked = True
                    End If
                Next
            Else
                For intCount = 0 To lvwInvoice.Items.Count - 1
                    If lvwInvoice.Items.Item(intCount).Checked = True Then
                        lvwInvoice.Items.Item(intCount).Checked = False
                    End If
                Next
            End If
            Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub optSelected_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptSelected.CheckedChanged
        If eventSender.Checked Then
            '----------------------------------------------------
            'Author              - Arshad Ali
            'Create Date         - 25/05/2003
            'Arguments           - None
            'Return Value        - None
            '----------------------------------------------------
            On Error GoTo ErrHandler
            Dim intCount As Integer
            If OptAll.Checked = True Then
                For intCount = 0 To lvwInvoice.Items.Count - 1
                    If lvwInvoice.Items.Item(intCount).Checked = False Then
                        lvwInvoice.Items.Item(intCount).Checked = True
                    End If
                Next
            Else
                For intCount = 0 To lvwInvoice.Items.Count - 1
                    If lvwInvoice.Items.Item(intCount).Checked = True Then
                        lvwInvoice.Items.Item(intCount).Checked = False
                    End If
                Next
            End If
            Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub

	Private Sub txtInvoiceTo_Change(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles txtInvoiceTo.Change
		'----------------------------------------------------
		'Author              - Arshad Ali
		'Create Date         - 25/05/2003
		'Arguments           - None
		'Return Value        - None
		'Function            -
		'----------------------------------------------------
		On Error GoTo ErrHandler
		lvwInvoice.Items.Clear()
ErrHandler: 
		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mp_connection)
	End Sub
	
	Private Sub txtInvoiceTo_KeyDown(ByVal Sender As System.Object, ByVal e As CtlGeneral.KeyDownEventArgs) Handles txtInvoiceTo.KeyDown
		Dim KeyCode As Short = e.KeyCode
		Dim Shift As Short = e.Shift
		If KeyCode = System.Windows.Forms.Keys.F1 Then
			cmdHelpInvoice2_Click(cmdHelpInvoice2, New System.EventArgs())
		End If
		
	End Sub
	
	Private Sub txtInvoiceTo_KeyPress(ByVal Sender As System.Object, ByVal e As CtlGeneral.KeyPressEventArgs) Handles txtInvoiceTo.KeyPress
		Dim KeyAscii As Short = e.KeyAscii
		'----------------------------------------------------
		'Author              - Arshad Ali
		'Create Date         - 25/05/2003
		'Arguments           - None
		'Return Value        - None
		'Function            - To set focus on next control
		'----------------------------------------------------
		On Error GoTo Err_Handler
		If KeyAscii = 13 Then
            OptSelected.Checked = False
            Call FillInvoicesToList()
			System.Windows.Forms.SendKeys.Send(vbTab)
		End If
		If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
			KeyAscii = KeyAscii
		Else
			KeyAscii = 0
		End If
		Exit Sub
Err_Handler: 
		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mp_connection)
	End Sub
	
	Private Sub txtInvoiceTo_LostFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtInvoiceTo.LostFocus
		Call FillInvoicesToList()
	End Sub
	
    Private Sub txtUnitCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnitCode.TextChanged
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To cleare Date
        '----------------------------------------------------
        On Error GoTo ErrHandler
        txtInvoice.Text = ""
        txtInvoiceTo.Text = ""
        lvwInvoice.Items.Clear()
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
	Private Sub txtUnitCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnitCode.Enter
		'----------------------------------------------------
		'Author              - Arshad Ali
		'Create Date         - 25/05/2003
		'Arguments           - None
		'Return Value        - None
		'Function            - To Selected Set Focus
		'----------------------------------------------------
		On Error GoTo ErrHandler
		'Selecting the text on focus
		With txtUnitCode
			.SelectionStart = 0 : .SelectionLength = Len(Trim(.Text))
		End With
		Exit Sub 'This is to avoid the execution of the error handler
ErrHandler: 
		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mp_connection)
		Exit Sub
	End Sub
	Private Sub txtUnitCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtUnitCode.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		'----------------------------------------------------
		'Author              - Arshad Ali
		'Create Date         - 25/05/2003
		'Arguments           - None
		'Return Value        - None
		'Function            - To Call help on F1 Press
		'----------------------------------------------------
		On Error GoTo ErrHandler
		'If Ctrl/Alt/Shift is also pressed
		If Shift <> 0 Then Exit Sub
		'Show the help form when user pressed F1
		If KeyCode = System.Windows.Forms.Keys.F1 Then Call cmdUnitCodeList_Click(cmdUnitCodeList, New System.EventArgs())

		Exit Sub 'This is to avoid the execution of the error handler
ErrHandler: 
		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mp_connection)
		Exit Sub
	End Sub
	Private Sub txtUnitCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtUnitCode.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'----------------------------------------------------
		'Author              - Arshad Ali
		'Create Date         - 25/05/2003
		'Arguments           - None
		'Return Value        - None
		'Function            - To Set Focus on Next Control
		'----------------------------------------------------
		If KeyAscii = 13 Then
			System.Windows.Forms.SendKeys.Send("{TAB}")
			'Supressing ¬ ¤ ¦ » characters since these are being used as string delimiters
		ElseIf KeyAscii = 187 Or KeyAscii = 166 Or KeyAscii = 164 Or KeyAscii = 172 Or KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 96 Then 
			KeyAscii = 0
		End If
		KeyAscii = Asc(UCase(Chr(KeyAscii)))
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub txtUnitCode_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnitCode.Leave
		If dtcInvoiceType.Enabled = True Then dtcInvoiceType.Focus()
	End Sub
	
	Private Sub txtUnitCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtUnitCode.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		'----------------------------------------------------
		'Author              - Arshad Ali
		'Create Date         - 25/05/2003
		'Arguments           - None
		'Return Value        - None
		'Function            - To Validate selected unit code
		'----------------------------------------------------
		On Error GoTo ErrHandler
		Dim strUnitDesc As String
        Dim mobjGLTrans As New prj_GLTransactions.cls_GLTransactions(gstrUNITID, GetServerDate)
		Dim strSQL As String
		Dim rsSaleConf As New ADODB.Recordset
		
		'Populate the details
		
		If Trim(txtUnitCode.Text) = "" Then GoTo EventExitSub
		
		With dtcInvoiceType
            strSQL = "select Distinct invoice_type, description from Saleconf where UNIT_CODE = '" & gstrUNITID & "' AND Invoice_Type in('INV','SMP','TRF','REJ','JOB','EXP','SRC','CPV') and datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0"
            If rsSaleConf.State = ADODB.ObjectStateEnum.adStateOpen Then rsSaleConf.Close()
			rsSaleConf.Open(strSQL, mp_connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            Dim cnt As Integer

            For cnt = 0 To rsSaleConf.RecordCount - 1
                .Items.Insert(cnt, rsSaleConf.Fields("Description").Value)
                rsSaleConf.MoveNext()
            Next cnt

			If rsSaleConf.RecordCount > 0 Then
				rsSaleConf.MoveFirst()
                .Text = rsSaleConf.Fields("Description").Value
			End If
		End With
		dtcInvoiceType.Enabled = True
		dtcInvoiceType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
		dtcInvoiceSubType.Enabled = True
		dtcInvoiceSubType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
		txtInvoice.Enabled = True
		cmdHelpInvoice.Enabled = True
		txtInvoice.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
		txtInvoiceTo.Enabled = True
		cmdHelpInvoice2.Enabled = True
		txtInvoiceTo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        chkRemoval.Enabled = True
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler: 
		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mp_connection)
		
EventExitSub: 
		eventArgs.Cancel = Cancel
	End Sub

    Private Sub frmMKTTRN0034_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To intialise required
        '----------------------------------------------------
        On Error GoTo Err_Handler
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        txtInvoice.Enabled = True
        If txtInvoice.Enabled Then txtInvoice.Focus()
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0034_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To relesed Values
        '----------------------------------------------------
        On Error GoTo Err_Handler
        frmModules.NodeFontBold(Tag) = False
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
	Private Sub frmMKTTRN0034_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		'----------------------------------------------------
		'Author              - Arshad Ali
		'Create Date         - 25/05/2003
		'Arguments           - None
		'Return Value        - None
		'Function            - To display empower help on click of F4 in empower
		'----------------------------------------------------
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then Call ctlFormHeader1_Click(ctlFormHeader1, New System.EventArgs()) : Exit Sub
	End Sub
	Private Sub frmMKTTRN0034_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'----------------------------------------------------
		'Author              - Arshad Ali
		'Create Date         - 25/05/2003
		'Arguments           - None
		'Return Value        - None
		'Function            - To initialise required data
		'----------------------------------------------------
        On Error GoTo Err_Handler
        dtpRemoval.Format = DateTimePickerFormat.Custom
        dtpRemoval.CustomFormat = gstrDateFormat

		Dim intLoopCounter As Short
		mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
		Call FillLabelFromResFile(Me) 'To Fill label description from Resource file
		Call FitToClient(Me, fraInvoice, ctlFormHeader1, Cmdinvoice) 'To fit the form in the MDI
		optAll.Checked = False : optSelected.Checked = False
		AddColumnsInListView()
		gblnCancelUnload = False
		
		dtcInvoiceType.Enabled = False
		dtcInvoiceType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
		dtcInvoiceSubType.Enabled = False
		dtcInvoiceSubType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
		txtInvoice.Enabled = False
		txtInvoiceTo.Enabled = False
		txtInvoice.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
		txtInvoice.Text = ""
		txtInvoiceTo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
		txtInvoiceTo.Text = ""
		cmdHelpInvoice.Enabled = False
		cmdHelpInvoice2.Enabled = False
		dtpRemoval.Enabled = False : dtpRemovalTime.Enabled = False
		dtpRemoval.value = GetServerDate
        dtpRemovalTime.Value = dtpRemoval.Value
        txtUnitCode.Text = Find_Value("SELECT distinct l.Location_Code FROM Location_mst l,SaleConf s WHERE s.Location_Code = l.Location_Code AND s.UNIT_CODE = l.UNIT_CODE AND s.UNIT_CODE = '" & gstrUNITID & "' order by l.Location_Code desc")
		txtUnitCode_Validating(txtUnitCode, New System.ComponentModel.CancelEventArgs(False))
        lblInvoiceType.Text = "INV"
        dtcInvoiceType.Text = "NORMAL INVOICE"
        Call dtcInvoiceType_Click(eventSender, eventArgs)
        Call dtcInvoiceType_SelectedIndexChanged(eventSender, eventArgs)

        lblInvoiceSubType.Text = "F"
        dtcInvoiceSubType.Text = "FINISHED GOODS"
        If Not Directory.Exists(gstrLocalCDrive + "EmproInv") Then
            Directory.CreateDirectory(gstrLocalCDrive + "EmproInv")
        End If
        Exit Sub

Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0034_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Release memory
        '----------------------------------------------------
        On Error GoTo Err_Handler
        'REFRESH
        'Removing the form name from list
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        'Setting the corresponding node's tag
        frmModules.NodeFontBold(Tag) = False
        'Closing the recordset
        'Releasing the form reference
        Me.Dispose()
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function ValidatebeforeSave() As Boolean
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - Boolean
        'Function            - To Check Valid Feilds Value
        '----------------------------------------------------
        Dim strErrMsg As String
        Dim ctlBlank As System.Windows.Forms.Control
        Dim lNo As Integer
        Dim blnInvalidData As Boolean
        Dim intCount As Integer
        Dim blnSelected As Boolean
        On Error GoTo Err_Handler
        ValidatebeforeSave = False
        lNo = 1
        strErrMsg = ResolveResString(10059) & vbCrLf
        If Len(Trim(txtUnitCode.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & " Location Code"
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = txtUnitCode
        End If
        If Len(Trim(dtcInvoiceType.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & " Invoice Type"
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = dtcInvoiceType
        End If
        If Len(Trim(dtcInvoiceSubType.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & " Invoice Sub Type"
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = dtcInvoiceSubType
        End If
        If Len(Trim(txtInvoice.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & " From Invoice No."
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = txtInvoice
        End If
        If Len(Trim(txtInvoiceTo.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & " To Invoice No."
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = txtInvoiceTo
        End If

        For intCount = 0 To lvwInvoice.Items.Count - 1
            If lvwInvoice.Items.Item(intCount).Checked Then
                blnSelected = True
                If Len(Trim(lvwInvoice.Items.Item(intCount).Text)) = 0 Then
                    blnInvalidData = True
                    strErrMsg = strErrMsg & vbCrLf & lNo & ". " & ResolveResString(60373)
                    lNo = lNo + 1
                    If ctlBlank Is Nothing Then ctlBlank = lvwInvoice
                End If
            End If
        Next
        If Not blnSelected Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & ". No Invoice Selected."
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = lvwInvoice
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

    Private Sub AddColumnsInListView()
        '***********************************
        'To add Columns Headers in the ListView in the form load
        '***********************************
        On Error GoTo ErrHandler
        With Me.lvwInvoice
            mCtlHdrInvoiceNo = .Columns.Add("")
            mCtlHdrInvoiceNo.Text = "Invoice No"
            mCtlHdrInvoiceNo.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwInvoice.Width) / 2)
            mCtlHdrInvoiceDate = .Columns.Add("")
            mCtlHdrInvoiceDate.Text = "Invoice Date"
            mCtlHdrInvoiceDate.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwInvoice.Width) / 2 - 300)
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Public Sub FillInvoicesToList()
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Fill Invoices into the Listview
        '----------------------------------------------------
        On Error GoTo ErrHandler
        Dim strSQL As String
        Dim rsInvoice As New ClsResultSetDB
        If Trim(txtInvoice.Text) <> "" And Trim(txtInvoiceTo.Text) <> "" Then
            strSQL = "SELECT Doc_no,Invoice_date FROM Saleschallan_dtl WHERE UNIT_CODE = '" & gstrUNITID & "' AND invoice_type='" & lblInvoiceType.Text & "' and sub_category='" & lblInvoiceSubType.Text & "' and Bill_flag =1 and Cancel_flag =0 and doc_no >= " & Val(txtInvoice.Text) & " and doc_no <= " & Val(txtInvoiceTo.Text) & " order by doc_no"
            rsInvoice.GetResult(strSQL)
            If rsInvoice.GetNoRows > 0 Then
                lvwInvoice.Items.Clear()
                rsInvoice.MoveFirst()
                While Not rsInvoice.EOFRecord
                    mlvwInvoice = Me.lvwInvoice.Items.Add(rsInvoice.GetValue("doc_no"))
                    If mlvwInvoice.SubItems.Count > 1 Then
                        mlvwInvoice.SubItems(1).Text = setDateFormat(rsInvoice.GetValue("invoice_date"))
                    Else
                        mlvwInvoice.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, setDateFormat(rsInvoice.GetValue("invoice_date"))))
                    End If
                    rsInvoice.MoveNext()
                End While
            End If
            rsInvoice.ResultSetClose()
            rsInvoice = Nothing
        End If
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
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
		Rs.Open(strField, mp_connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
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
		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mp_connection)
	End Function
	
	'########################################################################################################
	'******************************************PRINTING******************************************************
	'########################################################################################################
	Sub PrintingInvoice()
		On Error GoTo ErrHandler
		Dim strInvoiceFileName As String
		Dim strBarcodeFileName As String
		Dim intCount As Integer
		
		objInvoicePrint.ConnectionString = gstrCONNECTIONSTRING
		objInvoicePrint.Connection()
		objInvoicePrint.CompanyName = gstrCOMPANY
		objInvoicePrint.Address1 = gstr_RGN_ADDRESS1
		objInvoicePrint.Address2 = gstr_RGN_ADDRESS2
		On Error Resume Next
        Kill(gstrLocalCDrive & "Tmpfile\I*.txt")
        Kill(gstrLocalCDrive & "Tmpfile\B*.txt")
		On Error GoTo ErrHandler
        For intCount = 0 To lvwInvoice.Items.Count - 1
            If lvwInvoice.Items.Item(intCount).Checked = True Then
                strInvoiceFileName = gstrLocalCDrive & "Tmpfile\I" & VB.Right(lvwInvoice.Items.Item(intCount).Text, 7) & ".txt"
                strBarcodeFileName = gstrLocalCDrive & "Tmpfile\B" & VB.Right(lvwInvoice.Items.Item(intCount).Text, 7) & ".txt"
                If mStrFirstFileName = "" Then mStrFirstFileName = strInvoiceFileName

                objInvoicePrint.FileName = strInvoiceFileName
                objInvoicePrint.BCFileName = strBarcodeFileName

                If chkRemoval.CheckState = System.Windows.Forms.CheckState.Checked Then
                    objInvoicePrint.Print_Invoice(gstrUNITID, True, (txtUnitCode.Text), (lvwInvoice.Items.Item(intCount).Text), dtpRemoval.Text & " " & VB6.Format(dtpRemovalTime.Value.Hour, "00") & ":" & VB6.Format(dtpRemovalTime.Value.Minute, "00"))
                Else
                    objInvoicePrint.Print_Invoice(gstrUNITID, True, (txtUnitCode.Text), (lvwInvoice.Items.Item(intCount).Text))
                End If
            End If
        Next
		Exit Sub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mp_connection)
		'Function called, if error occurred
	End Sub
	
	Private Sub ReplaceJunkCharacters()
		'----------------------------------------------------------------------------
		'Author         :   Arshad Ali
		'Argument       :   Non
		'Return Value   :   Non
		'Function       :   Removes all special characters used for formating from text file
		'Comments       :   Nil
		'----------------------------------------------------------------------------
		On Error GoTo Errorhandler
		rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(15), "") 'Remove Uncompress Character
		rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(18), "") 'Remove Decompress Character
		rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "G", "") 'Remove Bold Character
		rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "H", "") 'Remove DeBold Character
		rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(12), "") 'Remove DeUnderline Character
		rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "-1", "") 'Remove Underline Character
		rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "-0", "") 'Remove DeUnderline Character
		rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "W1", "") 'Remove DoubleWidth Character
		rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "W0", "") 'Remove DeDoubleWidth Character
		rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "M", "") 'Remove Middle Character
		rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "P", "") 'Remove DeMiddle Character
		rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "E", "") 'Remove Elite Character
		rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "F", "") 'Remove DeElite Character
		Exit Sub
Errorhandler: 
		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mp_connection)
	End Sub
	
	Sub PrintingInvoiceToPrinter()
		'----------------------------------------------------------------------------
		'Author         :   Arshad Ali
		'Argument       :   Non
		'Return Value   :   Non
		'Function       :   Printing of invoices to printer
		'Comments       :   Nil
		'----------------------------------------------------------------------------
		'Revised By                 -  Ashutosh Verma
		'Revision Date              -  22/12/2005
		'Revision History           -  Issue Id:16625, for printing Bar code on Invoice at SUNVAC according to check box for Bar code printing.
		'---------------------------------------------------------------------------------------------------------
		'Revised By        : Ashutosh on 07-01-2006
		'Revision History  : Issue Id:16780, To show Captive invoice on New Invoice Format (SUNVAC).
		'=============================================================
		
		On Error GoTo ErrHandler
		Dim strInvoiceFileName As String
		Dim strBarcodeFileName As String
		Dim intCount As Integer
		Dim varTemp As Object
		Dim dblWaitingTime As Double
		
		objInvoicePrint.ConnectionString = gstrCONNECTIONSTRING
		objInvoicePrint.Connection()
		objInvoicePrint.CompanyName = gstrCOMPANY
		objInvoicePrint.Address1 = gstr_RGN_ADDRESS1
		objInvoicePrint.Address2 = gstr_RGN_ADDRESS2
		On Error Resume Next
        Kill(gstrLocalCDrive & "Tmpfile\I*.txt")
        Kill(gstrLocalCDrive & "Tmpfile\B*.txt")
        'Kill(gstrLocalCDrive & "TypeToPrn.bat")
        Kill(gstrLocalCDrive & "EmproInv\BarCodePageFeed.txt")

		On Error GoTo ErrHandler
        dblWaitingTime = Val(Find_Value("select waitingTime from sales_parameter WHERE UNIT_CODE = '" & gstrUNITID & "'"))
		If dblWaitingTime = 0 Then
			dblWaitingTime = 5000
		End If
TypeFileNotFoundCreateRetry:
        For intCount = 0 To lvwInvoice.Items.Count - 1
            If lvwInvoice.Items.Item(intCount).Checked = True Then
                lblPrint.Text = "Printing Invoice No. " & lvwInvoice.Items.Item(intCount).Text
                strInvoiceFileName = gstrLocalCDrive & "Tmpfile\I" & VB.Right(lvwInvoice.Items.Item(intCount).Text, 7) & ".txt"
                strBarcodeFileName = gstrLocalCDrive & "Tmpfile\B" & VB.Right(lvwInvoice.Items.Item(intCount).Text, 7) & ".txt"
                If mStrFirstFileName = "" Then mStrFirstFileName = strInvoiceFileName

                objInvoicePrint.FileName = strInvoiceFileName '"C:\InvoicePrint.txt"
                objInvoicePrint.BCFileName = strBarcodeFileName '"C:\BarCode.txt"

                If chkRemoval.CheckState = System.Windows.Forms.CheckState.Checked Then
                    objInvoicePrint.Print_Invoice(gstrUNITID, True, (txtUnitCode.Text), (lvwInvoice.Items.Item(intCount).Text), dtpRemoval.Text & " " & VB6.Format(dtpRemovalTime.Value.Hour, "00") & ":" & VB6.Format(dtpRemovalTime.Value.Minute, "00"))
                Else
                    objInvoicePrint.Print_Invoice(gstrUNITID, True, (txtUnitCode.Text), (lvwInvoice.Items.Item(intCount).Text))
                End If
                lblPrint.Text = "Printing Invoice No. " & lvwInvoice.Items.Item(intCount).Text & "..."
                'varTemp = System.Diagnostics.Process.Start("C:\TypeToPrn.bat " & strInvoiceFileName)
                varTemp = Shell("cmd.exe /c " & gstrLocalCDrive & "EmproInv\TypeToPrn.bat " & strInvoiceFileName, AppWinStyle.Hide)
                Sleep(dblWaitingTime)

                If chkPrintBarCode.CheckState = System.Windows.Forms.CheckState.Checked Then
                    lblPrint.Text = "Printing Bar Code of Invoice No. " & lvwInvoice.Items.Item(intCount).Text & "......"
                    Call printBarCode(strBarcodeFileName)
                    Sleep(dblWaitingTime)
                Else
                    If UCase(Trim(txtUnitCode.Text)) <> "SUN" Then
                        Call PrintBarCodeSpace("BarCodeSkip.txt")
                        Sleep(dblWaitingTime)
                    Else
                        varTemp = Shell("cmd.exe /c " & gstrLocalCDrive & "EmproInv\TypeToPrn.bat " & gstrLocalCDrive & "EmproInv\BarCodePageFeed.txt", AppWinStyle.Hide)
                    End If
                End If

                If chkPrintBarCode.CheckState = System.Windows.Forms.CheckState.Checked Then
                    If UCase(Trim(txtUnitCode.Text)) <> "SUN" Then
                        lblPrint.Text = "Adjusting Page for Invoice No. " & lvwInvoice.Items.Item(intCount).Text & "........."
                        If ((intCount Mod 5) = 0) And (intCount <> 30 Or intCount <> 60 Or intCount <> 90) Then
                            varTemp = Shell("cmd.exe /c " & gstrLocalCDrive & "EmproInv\TypeToPrn.bat " & gstrLocalCDrive & "EmproInv\PageFeedSet.txt", AppWinStyle.Hide)
                        Else
                            varTemp = Shell("cmd.exe /c " & gstrLocalCDrive & "EmproInv\TypeToPrn.bat " & gstrLocalCDrive & "EmproInv\PageFeed.txt", AppWinStyle.Hide)
                        End If
                    Else
                        varTemp = Shell("cmd.exe /c " & gstrLocalCDrive & "EmproInv\TypeToPrn.bat " & gstrLocalCDrive & "EmproInv\BarCodePageFeed.txt", AppWinStyle.Hide)
                    End If
                Else
                End If
            End If
            lblPrint.Text = ""
        Next
        lblPrint.Text = ""
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        If Err.Number = 53 Then
            FileOpen(1, gstrLocalCDrive & "EmproInv\TypeToPrn.bat", OpenMode.Append)
            PrintLine(1, "Type %1> prn") '& Printer.Port
            FileClose(1)
            GoTo TypeFileNotFoundCreateRetry
        End If
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        'Function called, if error occurred
    End Sub

    Sub printBarCode(ByVal pstrFileName As String)
        'Author         :   Arshad Ali
        'Argument       :
        'Return Value   :
        'Function       :
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        Dim varTemp As Object
        Dim strString As String
        strString = gstrLocalCDrive & "EmproInv\pdf-dot.bat " & pstrFileName & " 4 2 2 1"
        varTemp = Shell("cmd.exe /c " & strString)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Sub PrintBarCodeSpace(ByVal pstrFileName As String)
        'Author         :   Brij Bohara
        'Argument       :   Name of File in which Space is stored
        'Return Value   :   None
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        Dim varTemp As Object
        varTemp = Shell("cmd.exe /c " & gstrLocalCDrive & "EmproInv\TypeToPrn.bat " & gstrLocalCDrive & "EmproInv\BarCodeSkip.txt", AppWinStyle.Hide)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtInvoice_Change(ByVal Sender As Object, ByVal e As System.EventArgs) Handles txtInvoice.Change
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Clear Related Data
        '----------------------------------------------------
        txtInvoiceTo.Text = ""
        lvwInvoice.Items.Clear()

    End Sub

    Private Sub txtInvoice_KeyDown(ByVal Sender As Object, ByVal e As CtlGeneral.KeyDownEventArgs) Handles txtInvoice.KeyDown
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.Shift
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            cmdHelpInvoice_Click(cmdHelpInvoice, New System.EventArgs())
        End If
        If KeyCode = Keys.Enter Then System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub txtInvoice_KeyPress(ByVal Sender As Object, ByVal e As CtlGeneral.KeyPressEventArgs) Handles txtInvoice.KeyPress
        Dim KeyAscii As Short = e.KeyAscii
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To set focus on next control
        '----------------------------------------------------
        On Error GoTo Err_Handler
        If KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send(vbTab)
        End If
        If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
            KeyAscii = KeyAscii
        Else
            KeyAscii = 0
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub

    Private Sub txtInvoice_KeyUp(ByVal Sender As Object, ByVal e As CtlGeneral.KeyUpEventArgs) Handles txtInvoice.KeyUp
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.Shift
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Set focus on Next Control
        '----------------------------------------------------
        On Error GoTo Err_Handler
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub

    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call Empower help
        '----------------------------------------------------
        MsgBox("No Help Attached to This Form", MsgBoxStyle.Information, "empower")

    End Sub

    Private Sub dtcInvoiceSubType_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtcInvoiceSubType.KeyDown
        If e.KeyCode = Keys.Enter Then System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub dtcInvoiceSubType_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dtcInvoiceSubType.KeyPress
        If Asc(e.KeyChar) = 13 Then System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub dtcInvoiceSubType_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtcInvoiceSubType.TextChanged
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            -
        '----------------------------------------------------
        On Error GoTo ErrHandler
        txtInvoice.Text = ""
        txtInvoiceTo.Text = ""
        lvwInvoice.Items.Clear()
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub

    Private Sub dtcInvoiceType_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtcInvoiceType.Click
        
    End Sub

    Private Sub dtcInvoiceType_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtcInvoiceType.KeyDown
        If e.KeyCode = Keys.Enter Then System.Windows.Forms.SendKeys.Send("{TAB}")
    End Sub

    Private Sub dtcInvoiceType_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dtcInvoiceType.KeyPress

    End Sub

    Private Sub dtcInvoiceType_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtcInvoiceType.TextChanged
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            -
        '----------------------------------------------------
        On Error GoTo ErrHandler
        dtcInvoiceSubType.Text = ""
        txtInvoice.Text = ""
        txtInvoiceTo.Text = ""
        lvwInvoice.Items.Clear()
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub

    Private Sub dtcInvoiceType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtcInvoiceType.SelectedIndexChanged
        On Error GoTo ErrHandler
        Dim strSQL As String
        Dim rsSaleConf As New ADODB.Recordset
        Dim rsSaleConf1 As New ClsResultSetDB
        With dtcInvoiceSubType
            strSQL = "SELECT distinct(invoice_type) as invoice_Type  FROM Saleconf WHERE UNIT_CODE = '" & gstrUNITID & "' AND description = '" & dtcInvoiceType.Text & "'"
            rsSaleConf1.GetResult(strSQL)
            If rsSaleConf1.GetNoRows > 0 Then
                lblInvoiceType.Text = rsSaleConf1.GetValue("invoice_Type")
            End If

            strSQL = "select Distinct sub_type, sub_type_description from Saleconf where UNIT_CODE = '" & gstrUNITID & "' AND Invoice_Type in('" & Me.lblInvoiceType.Text & "') and datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0 order by sub_type_description"

            If rsSaleConf.State = ADODB.ObjectStateEnum.adStateOpen Then rsSaleConf.Close()
            rsSaleConf.Open(strSQL, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

            Dim cnt As Integer
            .Items.Clear()
            For cnt = 0 To rsSaleConf.RecordCount - 1
                .Items.Insert(cnt, rsSaleConf.Fields("sub_type_description").Value)
                rsSaleConf.MoveNext()
            Next cnt

            If rsSaleConf.RecordCount > 0 Then
                rsSaleConf.MoveFirst()
                lblInvoiceSubType.Text = rsSaleConf.Fields("sub_type").Value
            End If
        End With
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub

    Private Sub dtcInvoiceSubType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtcInvoiceSubType.SelectedIndexChanged
        Dim strSQL As String
        Dim rsSaleConf As New ADODB.Recordset
        Dim rsSaleConf1 As New ClsResultSetDB

        strSQL = "select Distinct sub_type, sub_type_description from Saleconf where UNIT_CODE = '" & gstrUNITID & "' AND sub_type_description in('" & Me.dtcInvoiceSubType.Text & "') and datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0 order by sub_type_description"

        If rsSaleConf.State = ADODB.ObjectStateEnum.adStateOpen Then rsSaleConf.Close()
        rsSaleConf.Open(strSQL, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        With dtcInvoiceSubType
            If rsSaleConf.RecordCount > 0 Then
                rsSaleConf.MoveFirst()
                lblInvoiceSubType.Text = rsSaleConf.Fields("sub_type").Value
            End If
        End With
    End Sub
End Class