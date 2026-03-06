Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.IO
Friend Class frmMKTTRN0036
	Inherits System.Windows.Forms.Form
	'===================================================================================
	'(c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
	'File Name          :   FRMMKTTRN0036.frm
	'Function           :   Used for Range Printing of SRV57F4
	'Created By         :   Arshad Ali
	'Created On         :   10 June, 2004
	'Revision  By       :   Brij Bohara
	'Revision On        :   24  Nov, 2004
	'History            :   To add total Ecess
	'Revision On        :   13 Dec, 2004
	'History            :   To save printed SRV numbers
	'History            :   To add total Ecess
	'Revision On        :   24 Dec, 2004
	'History            :   To save KanBanNo
	'===================================================================================
	' Changed by Arshad Ali On 08-August-2005
	' Description : Insert statement changed to insert sales tax and
    '               excise tax details into PrintedSRV_dtl table
    'MODIFIED BY AJAY SHUKLA ON 10/MAY/2011 FOR MULTIUNIT CHANGE
	'===================================================================================
	
	Dim mintFormIndex As Double
	Dim mStrFirstFileName As String
    Dim mGrnQty As String
	Dim strYear As String
	Dim strMonth As String
	Dim strDay As String
	Dim blnPreview As Boolean
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
    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        FraInvoicePreview.Visible = False
        txtGrinFrom.Focus()
    End Sub
    Private Sub cmdGRNFrom_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGRNFrom.Click
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 11/06/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Show Help Form
        '----------------------------------------------------
        Dim varHelp As Object
        Dim strSQL As String
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        strSQL = "select  distinct c.doc_no, c.GRN_Date "
        strSQL = strSQL & " from cust_ord_hdr a, cust_ord_dtl b , grn_hdr c, grn_dtl  d  "
        strSQL = strSQL & " where "
        strSQL = strSQL & " a.account_code = b.account_code"
        strSQL = strSQL & " and a.cust_ref=b.cust_ref "
        strSQL = strSQL & " and a.amendment_no=b.amendment_no  "
        strSQL = strSQL & " and a.UNIT_CODE=b.UNIT_CODE  "
        strSQL = strSQL & " and c.doc_no=d.doc_no "
        strSQL = strSQL & " and c.doc_type=d.doc_type "
        strSQL = strSQL & " and c.from_location=d.from_location "
        strSQL = strSQL & " and c.UNIT_CODE=d.UNIT_CODE  "
        strSQL = strSQL & " and a.account_code=c.vendor_code "
        strSQL = strSQL & " and A.UNIT_CODE=C.UNIT_CODE "
        strSQL = strSQL & " and b.item_code=d.item_code "
        strSQL = strSQL & " and b.UNIT_CODE=d.UNIT_CODE"
        strSQL = strSQL & " and c.doc_category='U' "
        strSQL = strSQL & " and c.QA_authorized_code is not null  "
        strSQL = strSQL & " and c.doc_no like '" & Trim(txtGrinFrom.Text) & "%' and a.active_flag='A'  and a.authorized_flag=1 "
        strSQL = strSQL & " and a.UNIT_CODE='" & gstrUNITID & "' "
        strSQL = strSQL & " and c.doc_no not in  ( Select Distinct  Doc_no From PrintedSRV_Dtl WHERE UNIT_CODE='" & gstrUNITID & "' ) "
        varHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "GRIN Help")
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(varHelp) <> -1 Then
            If varHelp(0) <> "0" Then
                txtGrinFrom.Text = Trim(varHelp(0))
                txtGrinFrom.Focus()
            Else
                MsgBox("No Record Available", MsgBoxStyle.Information, "empower")
            End If
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdGRNTo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGRNTo.Click
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 11/06/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Show Help Form
        '----------------------------------------------------
        Dim varHelp As Object
        Dim strSQL As String
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        strSQL = "select  c.doc_no, c.GRN_Date "
        strSQL = strSQL & " from cust_ord_hdr a, cust_ord_dtl b , grn_hdr c, grn_dtl  d  "
        strSQL = strSQL & " where a.account_code = b.account_code "
        strSQL = strSQL & " and a.cust_ref=b.cust_ref "
        strSQL = strSQL & " and a.amendment_no=b.amendment_no  "
        strSQL = strSQL & " and a.UNIT_CODE =b.UNIT_CODE"
        strSQL = strSQL & " and c.doc_no=d.doc_no "
        strSQL = strSQL & " and c.doc_type=d.doc_type "
        strSQL = strSQL & " and c.from_location=d.from_location    "
        strSQL = strSQL & " and c.UNIT_CODE=d.UNIT_CODE"
        strSQL = strSQL & " and a.account_code=c.vendor_code"
        strSQL = strSQL & " and a.UNIT_CODE=c.UNIT_CODE "
        strSQL = strSQL & " and b.item_code=d.item_code "
        strSQL = strSQL & " and b.UNIT_CODE=d.UNIT_CODE"
        strSQL = strSQL & " and c.doc_category='U' "
        strSQL = strSQL & " and c.QA_authorized_code is not null "
        strSQL = strSQL & " and c.doc_no >   " & Val(txtGrinFrom.Text)
        strSQL = strSQL & " and c.doc_no like '" & Trim(txtGrinTo.Text) & "%' "
        strSQL = strSQL & " and a.active_flag='A'  "
        strSQL = strSQL & " and a.authorized_flag=1 "
        strSQL = strSQL & " and a.UNIT_CODE='" & gstrUNITID & "' "
        strSQL = strSQL & " and  c.doc_no not in ( Select Distinct  Doc_no From PrintedSRV_Dtl  WHERE UNIT_CODE='" & gstrUNITID & "'  ) "
        varHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "GRIN Help")
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(varHelp) <> -1 Then

            If varHelp(0) <> "0" Then

                txtGrinTo.Text = Trim(varHelp(0))
                txtGrinTo.Focus()
            Else
                MsgBox("No Record Available", MsgBoxStyle.Information, "empower")
            End If
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
        Dim intCount As Short
        Dim fs As Scripting.FileSystemObject
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
                With fpSpread
                    For intCount = 1 To .MaxRows
                        .Col = 1
                        .Row = intCount
                        If CBool(.Value) = True Then
                            intTotalselected = intTotalselected + 1
                        End If
                    Next
                End With
                If intTotalselected > 1 Then
                    MsgBox("Please select only one Invoice at once to preview.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "empower")
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    Exit Sub
                End If
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                If ValidatebeforeSave() Then
                    If CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'")) Then
                        mStrFirstFileName = ""
                        fs = CreateObject("Scripting.FileSystemObject")
                        On Error Resume Next
                        fs.CreateFolder(gstrLocalCDrive & "Tmpfile")
                        On Error GoTo ErrHandler
                        'Generating all text files
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
                    If CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'")) Then
                        mStrFirstFileName = ""
                        fs = CreateObject("Scripting.FileSystemObject")
                        On Error Resume Next
                        fs.CreateFolder(gstrLocalCDrive & "Tmpfile")
                        On Error GoTo ErrHandler
                        'Generatin all text files and send to printer them
                        Call PrintingInvoiceToPrinter()
                    End If
                End If

                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        End Select
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrint.Click
        '----------------------------------------------------
        'Author              - Brij Bohara
        'Create Date         - 24/11/2004
        'Arguments           - None
        'Return Value        - None
        'Function            - To print the previewed
        '----------------------------------------------------
        On Error GoTo ErrHandler
        Call Cmdinvoice_ButtonClick(Cmdinvoice, New UCActXCtl.UCfraRepCmd.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_PRINTER))
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub
    Private Sub frmMKTTRN0036_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
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
        txtGrinFrom.Enabled = True
        If txtGrinFrom.Enabled Then txtGrinFrom.Focus()
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0036_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
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
    Private Sub frmMKTTRN0036_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To initialise required data
        '----------------------------------------------------
        On Error GoTo Err_Handler
        Dim intLoopCounter As Short
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FitToClient(Me, fraInvoice, ctlFormHeader1, Cmdinvoice) 'To fit the form in the MDI
        OptAll.Checked = False : OptSelected.Checked = False
        gblnCancelUnload = False
        dtpRemoval.Enabled = False : dtpRemovalTime.Enabled = False
        dtpRemoval.Format = DateTimePickerFormat.Custom
        dtpRemoval.CustomFormat = gstrDateFormat
        dtpRemoval.Value = GetServerDate()
        dtpRemovalTime.Format = DateTimePickerFormat.Custom
        dtpRemovalTime.CustomFormat = "HH:mm"
        Call AllignGrid()
        If Directory.Exists(gstrLocalCDrive + "EmproInv") = False Then
            Directory.CreateDirectory(gstrLocalCDrive + "EmproInv")
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0036_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
        Dim intCount As Short
        Dim blnSelected As Boolean
        On Error GoTo Err_Handler
        ValidatebeforeSave = False
        lNo = 1
        strErrMsg = ResolveResString(10059) & vbCrLf
        If Len(Trim(txtGrinFrom.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & " From GRIN No."
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = txtGrinFrom
        End If
        If Len(Trim(txtGrinTo.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & " To GRIN No."
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = txtGrinTo
        End If
        With fpSpread
            For intCount = 1 To .MaxRows
                .Col = 1
                .Row = intCount
                If CBool(.Value) = True Then
                    blnSelected = True
                End If
            Next
        End With
        If Not blnSelected Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & ". No GRIN Selected."
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = fpSpread
        End If
        strErrMsg = VB.Left(strErrMsg, Len(strErrMsg) - 1)
        strErrMsg = strErrMsg & "."
        lNo = lNo + 1
        If blnInvalidData = True Then
            gblnCancelUnload = True
            Call MsgBox(strErrMsg, MsgBoxStyle.Information, "Error")
            txtGrinFrom.Focus()
            Exit Function
        End If
        ValidatebeforeSave = True
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
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
        Dim rsGRN As New ClsResultSetDB
        Dim intMaxLoop As Short
        Dim intLoopCounter As Short
        If Trim(txtGrinFrom.Text) <> "" And Trim(txtGrinTo.Text) <> "" Then
            strSQL = "select distinct   c.Pur_order_no, c.Doc_no, convert(varchar,c.GRN_Date,103) as GRN_Date, c.invoice_no, "
            strSQL = strSQL & " convert(varchar,c.invoice_date,103) as invoice_date, a.cust_ref "
            strSQL = strSQL & " from cust_ord_hdr a, cust_ord_dtl b , grn_hdr c, grn_dtl  d  "
            strSQL = strSQL & " where "
            strSQL = strSQL & " a.account_code = b.account_code"
            strSQL = strSQL & " and a.cust_ref=b.cust_ref "
            strSQL = strSQL & " and a.amendment_no=b.amendment_no  "
            strSQL = strSQL & " and a.UNIT_CODE=b.UNIT_CODE  "
            strSQL = strSQL & " and c.doc_no=d.doc_no "
            strSQL = strSQL & " and c.doc_type=d.doc_type "
            strSQL = strSQL & " and c.from_location=d.from_location "
            strSQL = strSQL & " and c.UNIT_CODE=d.UNIT_CODE  "
            strSQL = strSQL & " and a.account_code=c.vendor_code "
            strSQL = strSQL & " and a.UNIT_CODE=c.UNIT_CODE  "
            strSQL = strSQL & " and b.item_code=d.item_code "
            strSQL = strSQL & " and b.UNIT_CODE=d.UNIT_CODE "
            strSQL = strSQL & " and c.doc_category='U' "
            strSQL = strSQL & " and c.QA_authorized_code is not null "
            strSQL = strSQL & " and c.doc_no >=   " & Val(txtGrinFrom.Text)
            strSQL = strSQL & " and c.doc_no <=  " & Val(txtGrinTo.Text)
            strSQL = strSQL & " and a.active_flag='A'  "
            strSQL = strSQL & " and a.authorized_flag=1 "
            strSQL = strSQL & " and a.UNIT_CODE='" & gstrUNITID & "'"
            strSQL = strSQL & " and c.doc_no NOT IN ( SELECT distinct Doc_No from PrintedSRV_Dtl  WHERE UNIT_CODE='" & gstrUNITID & "'  )"
            rsGRN.GetResult(strSQL)
            With fpSpread
                OptAll.Checked = False : OptSelected.Checked = False
                .MaxRows = 0
                intMaxLoop = rsGRN.GetNoRows
                rsGRN.MoveFirst()
                For intLoopCounter = 1 To intMaxLoop
                    .MaxRows = .MaxRows + 1
                    .Row = intLoopCounter
                    .set_RowHeight(intLoopCounter, 315)
                    .Col = 1 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .TypeCheckCenter = True
                    .Col = 2

                    .Text = Trim(rsGRN.GetValue("Pur_order_no")) : .Lock = True
                    .Col = 3

                    .Text = Trim(rsGRN.GetValue("Doc_no")) : .Lock = True
                    .Col = 4

                    .Text = VB6.Format(rsGRN.GetValue("GRN_Date"), gstrDateFormat) : .Lock = True
                    .Col = 5

                    .Text = Trim(rsGRN.GetValue("invoice_no")) : .Lock = True
                    .Col = 6

                    .Text = VB6.Format(Trim(rsGRN.GetValue("invoice_date")), gstrDateFormat) : .Lock = True
                    .Col = 7

                    .Text = Trim(rsGRN.GetValue("cust_ref")) : .Lock = True
                    .Col = 8 : .Text = "" : .Lock = False
                    .Col = 9 : .Text = "" : .Lock = False
                    .Col = 10 : .Text = "" : .Lock = True
                    .Col = 11 : .Text = "" : .Lock = False
                    .Col = 12 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonPicture = My.Resources.ico111.ToBitmap : .Lock = False
                    rsGRN.MoveNext() 'move to next record
                Next
            End With
            rsGRN.ResultSetClose()

            rsGRN = Nothing
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
        Dim rs As New ADODB.Recordset
        rs = New ADODB.Recordset
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.Open(strField, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
        If rs.RecordCount > 0 Then

            If IsDBNull(rs.Fields(0).Value) = False Then
                Find_Value = rs.Fields(0).Value
            Else
                Find_Value = ""
            End If
        Else
            Find_Value = ""
        End If
        rs.Close()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    '########################################################################################################
    '******************************************PRINTING******************************************************
    '########################################################################################################
    Sub PrintingInvoice()
        '----------------------------------------------------------------------------
        'Author         :   Arshad Ali
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Prepares the text files
        'Comments       :   Nil
        '----------------------------------------------------------------------------

        On Error GoTo ErrHandler
        Dim strInvoiceFileName As String
        Dim strBarcodeFileName As String
        Dim intCount As Short
        Dim strDILocation As String
        Dim strDINo As String
        Dim strCustRef As String
        Dim strSQLInsert As String ' To insert VBalues in Printed SRV
        Dim blnInsert As Boolean ' To save in Table or Not
        Dim strKanBan As String 'To save Kanban No.
        Dim strSQLSeg As String
        Dim rsVal As New ClsResultSetDB
        blnInsert = False

        objInvoicePrint.ConnectionString = gstrCONNECTIONSTRING
        objInvoicePrint.Connection()
        objInvoicePrint.CompanyName = gstrCOMPANY
        objInvoicePrint.Address1 = gstr_RGN_ADDRESS1
        objInvoicePrint.Address2 = gstr_RGN_ADDRESS2
        On Error Resume Next
        Kill(gstrLocalCDrive & "Tmpfile\S*.txt")
        Kill(gstrLocalCDrive & "Tmpfile\B*.txt")
        On Error GoTo ErrHandler
        With fpSpread
            For intCount = 1 To .MaxRows
                .Col = 1
                .Row = intCount
                If CBool(.Value) = True Then
                    .Row = intCount
                    .Col = 3
                    strInvoiceFileName = gstrLocalCDrive & "Tmpfile\S" & VB.Right(.Text, 7) & ".txt"
                    strBarcodeFileName = gstrLocalCDrive & "Tmpfile\B" & VB.Right(.Text, 7) & ".txt"
                    If mStrFirstFileName = "" Then mStrFirstFileName = strInvoiceFileName
                    objInvoicePrint.FileName = strInvoiceFileName
                    objInvoicePrint.BCFileName = strBarcodeFileName

                    .Row = intCount
                    .Col = 7
                    strCustRef = Trim(.Text)
                    .Col = 8
                    strDILocation = Trim(.Text)
                    .Col = 9
                    strDINo = Trim(.Text)
                    .Col = 3

                    .Col = 11
                    .Row = intCount
                    strKanBan = Trim(.Text)
                    .Col = 3

                    If Len(Find_Value("Select KanBan_No from printedsrv_dtl where Doc_no='" & Trim(.Text) & "' and unit_code='" & gstrUNITID & "'")) = 0 Then

                        ''Added for KAnBanNo
                        If MsgBox("Do you want to save these GRIN numbers in file ", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Empower") = MsgBoxResult.Yes Then
                            blnInsert = True
                        Else
                            If MsgBox("If you will not save these GRIN numbers then these GRIN may be reprinted" & vbCrLf & " Are you sure you do not  want to save these GRIN numbers", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Empower") = MsgBoxResult.Yes Then
                                blnInsert = False
                            Else
                                blnInsert = True
                            End If
                        End If
                        'Addition ends here

                        If UpdateSchedule() Then
                            mStrFirstFileName = ""
                            Exit Sub
                        End If

                    Else
                        MsgBox("GRIN Number is already saved.", MsgBoxStyle.Information, "empower")
                        blnInsert = False
                    End If

                    .Col = 11
                    .Row = intCount
                    strKanBan = Trim(.Text)
                    .Col = 3

                    strSQLInsert = "Select c.kanbanno,c.quantity - (sum(isnull(b.sales_quantity,0)) + sum(isnull(p.sales_quantity,0))) as Balance "
                    strSQLInsert = strSQLInsert & " from mkt_enagareDtl c "
                    strSQLInsert = strSQLInsert & " left outer join sales_dtl b "
                    strSQLInsert = strSQLInsert & " on b.srvdino = c.kanbanNo "
                    strSQLInsert = strSQLInsert & " and b.UNIT_CODE=c.UNIT_CODE"
                    strSQLInsert = strSQLInsert & " left outer join salesChallan_dtl a "
                    strSQLInsert = strSQLInsert & " on a.location_code = b.location_code "
                    strSQLInsert = strSQLInsert & " and a.doc_no=b.doc_no "
                    strSQLInsert = strSQLInsert & " and a.UNIT_CODE=b.UNIT_CODE"
                    strSQLInsert = strSQLInsert & " and a.bill_flag= 1  "
                    strSQLInsert = strSQLInsert & " Left Outer join PrintedSRV_Dtl as p "
                    strSQLInsert = strSQLInsert & " on c.kanbanno=p.kanban_no "
                    strSQLInsert = strSQLInsert & " and c.UNIT_CODE =p.UNIT_CODE"
                    strSQLInsert = strSQLInsert & " where c.kanbanno = '" & strKanBan & "' "
                    strSQLInsert = strSQLInsert & " and c.UNIT_CODE='" & gstrUNITID & "'"
                    strSQLInsert = strSQLInsert & " group by c.kanbanno, c.quantity "


                    If blnInsert = True And Val(Find_Value(strSQLInsert)) < CDbl(mGrnQty) Then
                        MsgBox("GRIN Quantity is greater than Balance Quantity of Kanban. Cannot Save GRIN Number", MsgBoxStyle.Information, "empower")
                        blnInsert = False
                        blnPreview = False
                    End If


                    If blnInsert = True Then

                        blnPreview = True
                        strSQLInsert = "Insert into PrintedSRV_Dtl"
                        strSQLInsert = strSQLInsert & " (Account_code, Doc_no, ddt, item_code, pdt, active_flag, cust_vendor_code,"
                        strSQLInsert = strSQLInsert & " cust_item_code, cust_item_desc, invoice_no, invoice_date, cust_ref, sales_quantity, "
                        strSQLInsert = strSQLInsert & " rate, box_qty, DILocation, DINo, Kanban_No,"
                        strSQLInsert = strSQLInsert & " SalesTax_Type, SalesTax_Per, SalesTax_Amount, "
                        strSQLInsert = strSQLInsert & " Excise_Type, Excise_Per, Excise_Tax,UNIT_CODE) "
                        strSQLSeg = "select distinct isnull(c.account_code,'') as account_code, a.doc_no, ddt=convert(char(10),a.grn_date,103) ,h.item_code ,  pdt=convert(char(10),c.order_date,103),c.active_flag,  e.cust_vendor_code, cust_item_code=f.cust_drgno, cust_item_desc=f.drg_desc,a.invoice_no,"
                        strSQLSeg = strSQLSeg & " convert(varchar,a.invoice_date,103) as invoice_date,c.cust_ref,sales_quantity=b.accepted_quantity,rate=b.item_rate,"
                        strSQLSeg = strSQLSeg & " isnull(h.box_Qty,0) box_Qty ,'" & strDILocation & "' as DILocation ,'" & strDINo & "' as DiNo,'" & strKanBan & "' as KanBan_No,"
                        strSQLSeg = strSQLSeg & " g1.TxRt_Rate_no as SalesTax_Type, isnull(g1.TxRt_percentage,0) as sales_tax,"
                        strSQLSeg = strSQLSeg & " round(((b.item_rate * b.accepted_quantity)+ round(((b.item_rate * b.accepted_quantity) / 100) * isnull(g2.TxRt_percentage,0),2))/ 100 * isnull(g1.TxRt_percentage,0),2)   as sales_Tax_Amount,"
                        strSQLSeg = strSQLSeg & " g2.TxRt_Rate_no as Excise_Type, isnull(g2.TxRt_percentage,0) as excise_per,"
                        strSQLSeg = strSQLSeg & " round(((b.item_rate * b.accepted_quantity) / 100) * isnull(g2.TxRt_percentage,0),2) as excise_amount,A.UNIT_CODE"
                        strSQLSeg = strSQLSeg & " from grn_hdr a"
                        strSQLSeg = strSQLSeg & " inner join grn_dtl b on a.doc_no=b.doc_no AND a.UNIT_CODE=b.UNIT_CODE "
                        strSQLSeg = strSQLSeg & " inner join cust_ord_hdr   c on a.vendor_code=c.account_code  AND a.UNIT_CODE=C.UNIT_CODE and c.active_flag = 'A'"
                        strSQLSeg = strSQLSeg & " inner join cust_ord_dtl   d on c.account_code = d.account_code and c.cust_ref = d.cust_ref and c.amendment_no = d.amendment_no AND C.UNIT_CODE=D.UNIT_CODE and d.active_flag='A'"
                        strSQLSeg = strSQLSeg & " inner join customer_mst   e on a.vendor_code=e.customer_code  AND a.UNIT_CODE=e.UNIT_CODE "
                        strSQLSeg = strSQLSeg & " inner join custitem_mst   f on a.vendor_code=f.account_code and b.item_code=f.item_code AND a.UNIT_CODE=f.UNIT_CODE "
                        strSQLSeg = strSQLSeg & " inner join item_mst       h on b.item_code=h.item_code   AND b.UNIT_CODE=h.UNIT_CODE "
                        strSQLSeg = strSQLSeg & " left outer join Gen_TaxRate g1 on c.SalesTax_Type = g1.TxRt_Rate_no AND C.UNIT_CODE=G1.UNIT_CODE "
                        strSQLSeg = strSQLSeg & " left outer join Gen_TaxRate g2 on d.excise_duty = g2.TxRt_Rate_no AND D.UNIT_CODE=G2.UNIT_CODE "

                        strSQLSeg = strSQLSeg & " Where a.Doc_No = " & Trim(.Text) & ""
                        strSQLSeg = strSQLSeg & " AND a.UNIT_CODE= '" & gstrUNITID & "'"
                        strSQLSeg = strSQLSeg & " and c.cust_ref='" & strCustRef & "'"

                        mP_Connection.Execute(strSQLInsert & strSQLSeg, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                        rsVal.GetResult(strSQLSeg)
                        Call UpdateMktSchedule(rsVal.GetValue("Account_code"), strYear, strMonth, strDay, rsVal.GetValue("Cust_item_code"), rsVal.GetValue("Item_code"))
                        rsVal = Nothing
                    End If
                    '----------
                    If chkRemoval.CheckState = System.Windows.Forms.CheckState.Checked Then
                        objInvoicePrint.Print_SRV57F4(gstrUNITID, False, Trim(.Text), strCustRef, dtpRemoval.Text & " " & VB6.Format(dtpRemovalTime.Value.Hour, "00") & ":" & VB6.Format(dtpRemovalTime.Value.Minute, "00"), strDILocation, strDINo)
                    Else
                        objInvoicePrint.Print_SRV57F4(gstrUNITID, False, Trim(.Text), strCustRef, "", strDILocation, strDINo)
                    End If
                End If
            Next
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Sub PrintingInvoiceToPrinter()
        '----------------------------------------------------------------------------
        'Author         :   Arshad Ali
        'Argument       :   Non
        'Return Value   :   Non
        'Function       :   Printing of invoices to printer
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strInvoiceFileName As String
        Dim strBarcodeFileName As String
        Dim intCount As Short
        Dim varTemp As Object
        Dim strDILocation As String
        Dim strDINo As String
        Dim strCustRef As String
        Dim dblWaitingTime As Double
        Dim strSQLInsert As String ' To insert VBalues in Printed SRV
        Dim blnInsert As Boolean ' To save in Table or Not
        Dim strKanBan As String 'To save Kanban No.
        Dim strSQLSeg As String
        Dim rsGrinDate As New ADODB.Recordset

        Dim rsVal As New ClsResultSetDB
        blnInsert = False

        objInvoicePrint.ConnectionString = gstrCONNECTIONSTRING
        objInvoicePrint.Connection()
        objInvoicePrint.CompanyName = gstrCOMPANY
        objInvoicePrint.Address1 = gstr_RGN_ADDRESS1
        objInvoicePrint.Address2 = gstr_RGN_ADDRESS2
        On Error Resume Next
        Kill(gstrLocalCDrive & "Tmpfile\S*.txt")
        Kill(gstrLocalCDrive & "Tmpfile\B*.txt")
        Kill(gstrLocalCDrive & "EmproInv\TypeToPrn.bat")
        On Error GoTo ErrHandler
        dblWaitingTime = Val(Find_Value("select waitingTime from sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'"))
        If dblWaitingTime = 0 Then
            dblWaitingTime = 5000
        End If
        If blnPreview = True Then
            Call PrintAfterPriew()
        End If

        If MsgBox("Do you want to save these GRIN numbers in file ", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Empower") = MsgBoxResult.Yes Then
            blnInsert = True
        Else
            If MsgBox("If you will not save these GRIN numbers then these GRIN may be reprinted" & vbCrLf & " Are you sure you do not  want to save these GRIN numbers", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Empower") = MsgBoxResult.Yes Then
                blnInsert = False
            Else
                blnInsert = True
            End If
        End If

TypeFileNotFoundCreateRetry:
        With fpSpread
            If UpdateSchedule() Then
                Exit Sub
            End If
            For intCount = 1 To .MaxRows
                .Row = intCount
                .Col = 1
                If CBool(.Value) Then
                    .Col = 7
                    strCustRef = Trim(.Text)
                    .Col = 8
                    strDILocation = Trim(.Text)
                    .Col = 9
                    strDINo = Trim(.Text)
                    .Col = 11
                    strKanBan = Trim(.Text)
                    .Col = 3
                    If blnInsert = True Then
                        strSQLInsert = "Insert into PrintedSRV_Dtl"
                        strSQLInsert = strSQLInsert & " (Account_code, Doc_no, ddt, item_code, pdt, active_flag, cust_vendor_code,"
                        strSQLInsert = strSQLInsert & " cust_item_code, cust_item_desc, invoice_no, invoice_date, cust_ref, sales_quantity, "
                        strSQLInsert = strSQLInsert & " rate, box_qty, DILocation, DINo, Kanban_No,"
                        strSQLInsert = strSQLInsert & " SalesTax_Type, SalesTax_Per, SalesTax_Amount, "
                        strSQLInsert = strSQLInsert & " Excise_Type, Excise_Per, Excise_Tax, UNIT_CODE) "

                        strSQLSeg = "select distinct isnull(c.account_code,'') as account_code, a.doc_no, ddt=convert(char(10),a.grn_date,103) ,h.item_code ,  pdt=convert(char(10),c.order_date,103),c.active_flag,  e.cust_vendor_code, cust_item_code=f.cust_drgno, cust_item_desc=f.drg_desc,a.invoice_no,"
                        strSQLSeg = strSQLSeg & " convert(varchar,a.invoice_date,103) as invoice_date,c.cust_ref,sales_quantity=b.accepted_quantity,rate=b.item_rate,"
                        strSQLSeg = strSQLSeg & " isnull(h.box_Qty,0) box_Qty ,'" & strDILocation & "' as DILocation ,'" & strDINo & "' as DiNo,'" & strKanBan & "' as KanBan_No,"
                        strSQLSeg = strSQLSeg & " g1.TxRt_Rate_no as SalesTax_Type, isnull(g1.TxRt_percentage,0) as sales_tax,"
                        strSQLSeg = strSQLSeg & " round(((b.item_rate * b.accepted_quantity)+ round(((b.item_rate * b.accepted_quantity) / 100) * isnull(g2.TxRt_percentage,0),2))/ 100 * isnull(g1.TxRt_percentage,0),2)   as sales_Tax_Amount,"
                        strSQLSeg = strSQLSeg & " g2.TxRt_Rate_no as Excise_Type, isnull(g2.TxRt_percentage,0) as excise_per,"
                        strSQLSeg = strSQLSeg & " round(((b.item_rate * b.accepted_quantity) / 100) * isnull(g2.TxRt_percentage,0),2) as excise_amount, a.UNIT_CODE"

                        strSQLSeg = strSQLSeg & " from grn_hdr a"
                        strSQLSeg = strSQLSeg & " inner join grn_dtl b on a.doc_no=b.doc_no and a.unit_code=b.unit_code"
                        strSQLSeg = strSQLSeg & " inner join cust_ord_hdr   c on a.vendor_code=c.account_code and a.unit_code=c.unit_code and c.active_flag = 'A'"
                        strSQLSeg = strSQLSeg & " inner join cust_ord_dtl   d on c.account_code = d.account_code and c.cust_ref = d.cust_ref and c.amendment_no = d.amendment_no and c.unit_code=d.unit_code and d.active_flag='A'"
                        strSQLSeg = strSQLSeg & " inner join customer_mst   e on a.vendor_code=e.customer_code and  a.unit_code=e.unit_code "
                        strSQLSeg = strSQLSeg & " inner join custitem_mst   f on a.vendor_code=f.account_code and b.item_code=f.item_code and a.unit_code=f.unit_code "
                        strSQLSeg = strSQLSeg & " inner join item_mst       h on b.item_code=h.item_code and b.unit_code=h.unit_code"

                        strSQLSeg = strSQLSeg & " left outer join Gen_TaxRate g1 on c.SalesTax_Type = g1.TxRt_Rate_no and c.unit_code=g1.unit_code "
                        strSQLSeg = strSQLSeg & " left outer join Gen_TaxRate g2 on d.excise_duty = g2.TxRt_Rate_no and d.unit_code=g2.unit_code "
                        strSQLSeg = strSQLSeg & " Where a.Doc_No = " & Trim(.Text) & ""
                        strSQLSeg = strSQLSeg & " and a.unit_code= '" & gstrUNITID & "'"
                        strSQLSeg = strSQLSeg & " and c.cust_ref='" & strCustRef & "'"

                        mP_Connection.Execute(strSQLInsert & strSQLSeg, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        rsVal.GetResult(strSQLSeg)
                        Call UpdateMktSchedule(rsVal.GetValue("Account_code"), strYear, strMonth, strDay, rsVal.GetValue("Cust_item_code"), rsVal.GetValue("Item_code"))

                        rsVal = Nothing
                    End If
                    '--Storing GRIN Info ends here

                    lblPrint.Text = "Printing Invoice No. " & .Text
                    strInvoiceFileName = gstrLocalCDrive & "Tmpfile\S" & VB.Right(.Text, 7) & ".txt"
                    strBarcodeFileName = gstrLocalCDrive & "Tmpfile\B" & VB.Right(.Text, 7) & ".txt"
                    If mStrFirstFileName = "" Then mStrFirstFileName = strInvoiceFileName
                    objInvoicePrint.FileName = strInvoiceFileName
                    objInvoicePrint.BCFileName = strBarcodeFileName
                    If chkRemoval.CheckState = System.Windows.Forms.CheckState.Checked Then
                        objInvoicePrint.Print_SRV57F4(gstrUNITID, False, Trim(.Text), strCustRef, dtpRemoval.Text & " " & VB6.Format(dtpRemovalTime.Value.Hour, "00") & ":" & VB6.Format(dtpRemovalTime.Value.Minute, "00"), strDILocation, strDINo)
                    Else
                        objInvoicePrint.Print_SRV57F4(gstrUNITID, False, Trim(.Text), strCustRef, "", strDILocation, strDINo)
                    End If
                    lblPrint.Text = "Printing Invoice No. " & .Text & "..."

                    varTemp = Shell("cmd.exe /c " & gstrLocalCDrive & "EmproInv\TypeToPrn.bat " & strInvoiceFileName, AppWinStyle.Hide)
                    Sleep(dblWaitingTime)

                    lblPrint.Text = "Printing Invoice No. " & .Text & "......"
                    Call printBarCode(strBarcodeFileName)
                    Sleep(dblWaitingTime)

                    lblPrint.Text = "Printing Invoice No. " & .Text & "........."
                    If (intCount Mod 5) = 0 And (intCount <> 30 Or intCount <> 60 Or intCount <> 90) Then
                        varTemp = Shell("cmd.exe /c " & gstrLocalCDrive & "EmproInv\TypeToPrn.bat " & gstrLocalCDrive & "EmproInv\PageFeedSet.txt", AppWinStyle.Hide)
                    Else
                        varTemp = Shell("cmd.exe /c " & gstrLocalCDrive & "EmproInv\TypeToPrn.bat " & gstrLocalCDrive & "EmproInv\PageFeed.txt", AppWinStyle.Hide)
                    End If
                End If
                lblPrint.Text = ""
            Next
        End With
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
        On Error GoTo ErrHandler
        strString = gstrLocalCDrive & "EmproInv\pdf-dot.bat " & pstrFileName & " 4 2 2 1"

        varTemp = Shell("cmd.exe /c " & strString)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub fpSpread_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles fpSpread.ClickEvent
        On Error GoTo errrHandler
        If eventArgs.col = 12 Then
            Call showNagareHelp()
        End If
        Exit Sub
errrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub fpSpread_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles fpSpread.KeyDownEvent
        On Error GoTo errrHandler
        If eventArgs.keyCode = System.Windows.Forms.Keys.F1 And eventArgs.shift = 0 Then
            Call showNagareHelp()
        End If
        Exit Sub
errrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub fpSpread_LeaveCell(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles fpSpread.LeaveCell
        Dim strSQL As String
        Dim rsKanBan As New ClsResultSetDB
        On Error GoTo errrHandler
        With fpSpread
            If eventArgs.col = 11 Then

                .Row = .ActiveRow
                .Col = 11
                If Len(Trim(.Text)) Then
                    strSQL = "select KanBanNo from mkt_enagareDtl where KanBanNo='" & Trim(.Text) & "' and unit_code='" & gstrUNITID & "'"
                    rsKanBan.GetResult(strSQL)
                    If rsKanBan.RowCount <= 0 Then
                        MsgBox("This KanBanNo is not valid", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "empower")
                        .Col = 11
                        .Text = ""
                        Exit Sub
                    End If
                End If
            End If
        End With
        Exit Sub
errrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub optAll_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptAll.CheckedChanged
        If eventSender.Checked Then
            Dim intCount As Short
            On Error GoTo ErrHandler
            With fpSpread
                For intCount = 1 To .MaxRows
                    .Row = intCount
                    .Col = 1
                    .Value = CStr(1)
                Next
            End With
            Exit Sub
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

        End If
    End Sub
    Private Sub optSelected_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptSelected.CheckedChanged
        If eventSender.Checked Then

            Dim intCount As Short
            On Error GoTo ErrHandler
            With fpSpread
                For intCount = 1 To .MaxRows
                    .Row = intCount
                    .Col = 1
                    .Value = CStr(0)
                Next
            End With
            Exit Sub
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

        End If
    End Sub
    Private Sub txtGrinFrom_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtGrinFrom.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call cmdGRNFrom_Click(cmdGRNFrom, New System.EventArgs())
        End If
        If KeyCode = 13 Then
            System.Windows.Forms.SendKeys.Send(vbTab)
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub
    Private Sub txtGrinTo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtGrinTo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call cmdGRNTo_Click(cmdGRNTo, New System.EventArgs())
        End If
        If KeyCode = 13 Then
            System.Windows.Forms.SendKeys.Send(vbTab)
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub
    Sub AllignGrid()
        '-----------------------------------------------------------------
        'Author         :   Arshad Ali
        'Arguments      :   None
        'Return Value   :   None
        'Function       :   To Allign/Format spread sheet
        '-----------------------------------------------------------------
        On Error GoTo ErrHandler

        With fpSpread
            .MaxCols = 12
            .MaxRows = 0
            .Row = 0
            .UnitType = FPSpreadADO.UnitTypeConstants.UnitTypeTwips
            .set_RowHeight(0, 315)
            .Col = 1 : .Text = "Check It" : .set_ColWidth(1, 700) : .TypeCheckCenter = True
            .Col = 2 : .Text = "PO No" : .set_ColWidth(2, 1000) : .Lock = True : .ColHidden = True
            .Col = 3 : .Text = "Grin No" : .set_ColWidth(3, 1000) : .Lock = True
            .Col = 4 : .Text = "Grin Date" : .set_ColWidth(4, 1000) : .Lock = True
            .Col = 5 : .Text = "Inv. No" : .set_ColWidth(5, 1000) : .Lock = True
            .Col = 6 : .Text = "Inv. Date" : .set_ColWidth(6, 1000) : .Lock = True
            .Col = 7 : .Text = "SO No." : .set_ColWidth(7, 1200) : .Lock = True
            .Col = 8 : .Text = "Location" : .set_ColWidth(8, 2000)
            .Col = 9 : .Text = "DI No" : .set_ColWidth(9, 2000)
            .Col = 10 : .Text = "" : .set_ColWidth(10, 2000) : .ColHidden = True
            .Col = 11 : .Text = "KanBan No" : .set_ColWidth(11, 2000) : .Lock = False
            .Col = 12 : .Text = " " : .set_ColWidth(12, 300)
            .Row = 1
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub
    Private Sub txtGrinTo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtGrinTo.Leave
        On Error GoTo ErrHandler
        FillInvoicesToList()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function UpdateSchedule() As Boolean
        '-----------------------------------------------------------------
        'Author         :   Brij Bohara
        'Arguments      :   None
        'Return Value   :   True or False depending upon the updation status successful/failure
        'Purpose        :   Updates the schedule Quantity
        '-----------------------------------------------------------------
        Dim strDINo As String 'To store DI No
        Dim strCustRef As String 'To store CustRef
        Dim strDILocation As String 'To store Location
        Dim strSQL As String 'To store Query String
        Dim rsSch As New ClsResultSetDB 'To retrieve the values from schedule
        Dim rsgrin As New ClsResultSetDB 'To retrive the cust item code
        Dim strKanBan As String
        Dim intCount As Short
        Dim StrItemCode As String
        Dim strItem As String
        Dim strCust As String
        On Error GoTo ErrHandler

        UpdateSchedule = True
        With fpSpread
            For intCount = 1 To .MaxRows
                .Row = intCount
                .Col = 1
                If CBool(.Value) Then
                    .Col = 7
                    strCustRef = Trim(.Text)
                    .Col = 8
                    strDILocation = Trim(.Text)
                    .Col = 9
                    strDINo = Trim(.Text)
                    .Col = 11
                    strKanBan = Trim(.Text)
                    .Col = 3
                    rsgrin.GetResult("Select * From mkt_eNagareDtl WHERE kanbanNo ='" & strKanBan & "' and unit_code='" & gstrUNITID & "'")

                    If rsgrin.RowCount > 0 Then
                        rsgrin.MoveFirst()
                        StrItemCode = rsgrin.GetValue("item_code")
                        strCustRef = rsgrin.GetValue("account_code")
                        strCust = rsgrin.GetValue("cust_DrgNo")
                        strYear = CStr(DatePart(Microsoft.VisualBasic.DateInterval.Year, rsgrin.GetValue("Sch_Date")))
                        strMonth = CStr(DatePart(Microsoft.VisualBasic.DateInterval.Month, rsgrin.GetValue("Sch_Date")))
                        strDay = CStr(DatePart(Microsoft.VisualBasic.DateInterval.Day, rsgrin.GetValue("Sch_Date")))
                    End If

                    If Len(Trim(strKanBan)) = 0 Then 'Not selected the KanBanNo
                        MsgBox("Select KanBanNo first for schedule", MsgBoxStyle.Information, "empower")
                        Exit Function
                    Else
                        strSQL = "SELECT * FROM DailyMktSchedule WHERE Account_Code='" & strCustRef & "' AND  datepart(yyyy,Trans_Date)='" & strYear & "' AND datepart(m,Trans_Date)='" & strMonth & "' and datepart(d,Trans_Date)='" & strDay & "' and unit_code='" & gstrUNITID & "'"
                        strSQL = strSQL & " and Cust_DrgNo ='" & strCust & "' and Item_code = '" & StrItemCode & "' and Status =1 "
                        rsSch.GetResult(strSQL)
                        'Checking that Sale for that KanBan is Already complete
                        If rsSch.GetNoRows <= 0 Then
                            MsgBox("No schedule found for this KanBan combonation ", MsgBoxStyle.Information, "empower")

                            rsSch = Nothing

                            rsgrin = Nothing
                            Exit Function
                        Else
                            rsSch.MoveFirst()
                        End If

                        'Compare the quantity in Printed SRV
                        strItem = "select ISNULL(sum(isnull(sales_Quantity,0)),0) as SalesQty"
                        strItem = strItem & " from PrintedSRV_dtl"
                        strItem = strItem & " Where kanban_no ='" & strKanBan & "' and Account_code='" & Trim(strCustRef) & "' and cust_Item_code='" & Trim(strCust) & "' and unit_code='" & gstrUNITID & "'"

                        rsgrin.GetResult(strItem)
                        If rsgrin.RowCount > 0 Then rsgrin.MoveFirst() 'Compare the schedule quantity and Despatch Quantity
                        If Val(rsSch.GetValue("Schedule_Quantity")) - Val(rsSch.GetValue("Despatch_Qty")) <= 0 Or (Val(rsSch.GetValue("Schedule_Quantity")) - Val(rsgrin.GetValue("SalesQty")) = 0) Then
                            MsgBox("Quantity for this Kanban is already despatched " & vbCrLf & " Select another kanban ", MsgBoxStyle.Information, "empower")

                            rsSch = Nothing

                            rsgrin = Nothing
                            Exit Function
                        End If
                        If (Val(mGrnQty) > (Val(rsSch.GetValue("Schedule_Quantity")) - Val(rsSch.GetValue("Despatch_Qty")))) Then
                            MsgBox("Quantity for this Schedule is less than SRV quantity " & vbCrLf & " Select another kanban ", MsgBoxStyle.Information, "empower")
                            rsSch = Nothing
                            rsgrin = Nothing
                            Exit Function
                        End If
                    End If
                End If
            Next

        End With

        rsSch = Nothing

        rsgrin = Nothing
        UpdateSchedule = False
        Exit Function 'This is to avoid the execution of error handler

ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Function
    Private Function ValidateForPrint() As Boolean
        '-----------------------------------------------------------------
        'Author         :   Brij Bohara
        'Arguments      :   None
        'Return Value   :   True/False
        'Purpose        :   Updates the schedule Quantity
        '-----------------------------------------------------------------
        On Error GoTo ErrHandler
        'Check for quantity in DailyMktSchedule for that grin and itemcode
        ValidateForPrint = True
        Exit Function 'This is to avoid the execution of error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Function
    Private Sub showNagareHelp()
        On Error GoTo ErrHandler
        Dim strSQL As String
        Dim lngGRNNo As Integer
        Dim strNagare() As String
        Dim rsgrin As New ClsResultSetDB
        Dim strGRIN As String
        Dim strCustRef As String
        Dim strDILocation As String
        Dim strDINo As String
        Dim strKanBan As String
        '-------
        Dim strAccountCode As String
        Dim strCustDrgno As String
        Dim StrItemCode As String
        Dim strQuantity As String
        '--

        With fpSpread
            .Row = .ActiveRow
            .Col = 7
            strCustRef = Trim(.Text)
            .Col = 8
            strDILocation = Trim(.Text)
            .Col = 9
            strDINo = Trim(.Text)
            .Col = 11
            strKanBan = Trim(.Text)
            .Col = 3

            'Check quantity for that grin
            strGRIN = "select distinct isnull(c.account_code,'') as account_code, a.doc_no, ddt=convert(char(10),a.grn_date,103) ,h.item_code ,"
            strGRIN = strGRIN & " pdt=convert(char(10),c.order_date,103),c.active_flag,  e.cust_vendor_code,"
            strGRIN = strGRIN & " cust_item_code=f.cust_drgno, cust_item_desc=f.drg_desc,"
            strGRIN = strGRIN & " convert(varchar,a.invoice_date,103) as invoice_date,c.cust_ref,"
            strGRIN = strGRIN & " sales_quantity=isnull(b.accepted_quantity,0),rate=b.item_rate,"
            strGRIN = strGRIN & " isnull(h.box_Qty,0) box_Qty ,'" & strDILocation & "','" & strDINo & "','" & strKanBan & "' "
            strGRIN = strGRIN & " from grn_hdr a"
            strGRIN = strGRIN & " inner join grn_dtl b on a.doc_no=b.doc_no and a.unit_code=b.unit_code"
            strGRIN = strGRIN & " inner join cust_ord_hdr c on a.vendor_code=c.account_code and a.unit_code=c.unit_code and c.active_flag = 'A'"
            strGRIN = strGRIN & " inner join cust_ord_dtl d on c.account_code = d.account_code and c.cust_ref = d.cust_ref  and c.unit_code=d.unit_code"
            strGRIN = strGRIN & " and c.amendment_no = d.amendment_no and d.active_flag='A'"
            strGRIN = strGRIN & " inner join customer_mst e on a.vendor_code=e.customer_code and a.unit_code=e.unit_code"
            strGRIN = strGRIN & " inner join custitem_mst f on a.vendor_code=f.account_code "
            strGRIN = strGRIN & " and b.item_code=f.item_code and a.unit_code=f.unit_code"
            strGRIN = strGRIN & " inner join item_mst  h on b.item_code=h.item_code and b.unit_code=h.unit_code"
            strGRIN = strGRIN & " Where a.Doc_No = " & Trim(.Text) & ""
            strGRIN = strGRIN & " and c.cust_ref='" & strCustRef & "'"
            strGRIN = strGRIN & " and a.unit_code='" & gstrUNITID & "'"

            rsgrin.GetResult(strGRIN)


            strAccountCode = rsgrin.GetValue("account_code")

            strCustDrgno = rsgrin.GetValue("cust_item_code")

            StrItemCode = rsgrin.GetValue("item_code")

            If rsgrin.GetValue("sales_quantity") = "Unknown" Then
                strQuantity = "0"
            Else
                strQuantity = rsgrin.GetValue("sales_quantity")
            End If
            mGrnQty = strQuantity 'To store the Grin quantity

            strSQL = " Select cust_drgNo, m.KanbanNo, UNLOC,USLOC, Sch_Time,Sch_Date,Quantity, Quantity-consumeqty as Balance from MKT_Enagaredtl m  "
            strSQL = strSQL & " Inner join ( Select c.kanbanno, sum(isnull(b.sales_quantity,0)) + sum(isnull(p.sales_quantity,0)) as consumeqty, C.UNIT_CODE   "
            strSQL = strSQL & " from mkt_enagareDtl c   "
            strSQL = strSQL & " left outer join sales_dtl b "
            strSQL = strSQL & " 	on b.srvdino = c.kanbanNo AND B.UNIT_CODE=c.UNIT_CODE "
            strSQL = strSQL & " left outer join salesChallan_dtl a "
            strSQL = strSQL & " 	on a.location_code = b.location_code  AND A.UNIT_CODE=B.UNIT_CODE and a.doc_no=b.doc_no and a.bill_flag= 1 "
            strSQL = strSQL & " Left Outer join PrintedSRV_Dtl as p "
            strSQL = strSQL & " 	on c.kanbanno=p.kanban_no  AND c.UNIT_CODE=p.UNIT_CODE	"
            strSQL = strSQL & " where c.unit_code='" & gstrUNITID & "' group by c.kanbanno, c.quantity, c.UNIT_CODE ) as B   "
            strSQL = strSQL & " On m.kanbanno=B.kanbanno  AND m.UNIT_CODE=B.UNIT_CODE  "
            strSQL = strSQL & " WHERE m.Quantity > consumeqty"
            strSQL = strSQL & " AND m.UNIT_CODE='" & gstrUNITID & "'"
            .Col = 11
            If Len(Trim(.Text)) > 0 Then
                strSQL = strSQL & " and m.KanBanNo like '%" & Trim(.Text) & "%'"
            End If

            strSQL = strSQL & " order by sch_date desc, Sch_time asc"

            strNagare = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "eNagare Details")

            If UBound(strNagare) = -1 Then Exit Sub
            If strNagare(0) = "0" Then
                MsgBox("No Record Available to Display", MsgBoxStyle.Information, "empower")
            Else
                .Col = 11

                .Text = IIf(IsDBNull(strNagare(1)), "", strNagare(1))
            End If
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub UpdateMktSchedule(ByRef pstrAccountCode As String, ByRef pstrYear As String, ByRef pstrMon As String, ByRef pstrDay As String, ByRef pstrCust_Drg As String, ByRef pstrItem_code As String)
        On Error GoTo ErrHandler
        Dim strupdate As String

        strupdate = " Update DailyMktSchedule set Despatch_qty =isnull(Despatch_Qty,0)  +" & Val(mGrnQty) & ""
        strupdate = strupdate & " Where Account_Code= '" & pstrAccountCode & "' and  datepart(yyyy,Trans_Date)='" & pstrYear & "'"
        strupdate = strupdate & " and datepart(m,Trans_Date)='" & pstrMon & "' and datepart(d,Trans_Date)='" & pstrDay & "'"
        strupdate = strupdate & " and Cust_DrgNo ='" & pstrCust_Drg & "' and Item_code = '" & pstrItem_code & "' and Status =1 and unit_code='" & gstrUNITID & "'"

        mP_Connection.Execute(strupdate, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub PrintAfterPriew()
        '----------------------------------------------------------------------------
        'Author         :   Brij Bohara
        'Argument       :   None
        'Return Value   :   None
        'Function       :   Printing of invoices to printer which are viewed by priew
        'Comments       :   Nill
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strInvoiceFileName As String
        Dim strBarcodeFileName As String
        Dim intCount As Short
        Dim varTemp As Object
        Dim strDILocation As String
        Dim strDINo As String
        Dim strCustRef As String
        Dim dblWaitingTime As Double

        objInvoicePrint.ConnectionString = gstrCONNECTIONSTRING
        objInvoicePrint.Connection()
        objInvoicePrint.CompanyName = gstrCOMPANY
        objInvoicePrint.Address1 = gstr_RGN_ADDRESS1
        objInvoicePrint.Address2 = gstr_RGN_ADDRESS2
        On Error Resume Next
        Kill(gstrLocalCDrive & "Tmpfile\S*.txt")
        Kill(gstrLocalCDrive & "Tmpfile\B*.txt")
        Kill(gstrLocalCDrive & "EmproInv\TypeToPrn.bat")
        On Error GoTo ErrHandler
        dblWaitingTime = Val(Find_Value("select waitingTime from sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'"))
        If dblWaitingTime = 0 Then
            dblWaitingTime = 5000
        End If
TypeFileNotFoundCreateRetry:

        With fpSpread
            For intCount = 1 To .MaxRows
                .Col = 1
                If CBool(.Value) Then
                    .Col = 7
                    strCustRef = Trim(.Text)
                    .Col = 8
                    strDILocation = Trim(.Text)
                    .Col = 9
                    strDINo = Trim(.Text)
                    .Col = 3
                    lblPrint.Text = "Printing Invoice No. " & .Text
                    strInvoiceFileName = gstrLocalCDrive & "Tmpfile\S" & VB.Right(.Text, 7) & ".txt"
                    strBarcodeFileName = gstrLocalCDrive & "Tmpfile\B" & VB.Right(.Text, 7) & ".txt"
                    If mStrFirstFileName = "" Then mStrFirstFileName = strInvoiceFileName
                    objInvoicePrint.FileName = strInvoiceFileName
                    objInvoicePrint.BCFileName = strBarcodeFileName
                    If chkRemoval.CheckState = System.Windows.Forms.CheckState.Checked Then
                        objInvoicePrint.Print_SRV57F4(gstrUNITID, False, Trim(.Text), strCustRef, dtpRemoval.Text & " " & VB6.Format(dtpRemovalTime.Value.Hour, "00") & ":" & VB6.Format(dtpRemovalTime.Value.Minute, "00"), strDILocation, strDINo)
                    Else
                        objInvoicePrint.Print_SRV57F4(gstrUNITID, False, Trim(.Text), strCustRef, "", strDILocation, strDINo)
                    End If
                    lblPrint.Text = "Printing Invoice No. " & .Text & "..."

                    varTemp = Shell("cmd.exe /c " & gstrLocalCDrive & "EmproInv\TypeToPrn.bat " & strInvoiceFileName, AppWinStyle.Hide)
                    Sleep(dblWaitingTime)

                    lblPrint.Text = "Printing Invoice No. " & .Text & "......"
                    Call printBarCode(strBarcodeFileName)
                    Sleep(dblWaitingTime)
                    lblPrint.Text = "Printing Invoice No. " & .Text & "........."
                    If (intCount Mod 5) = 0 And (intCount <> 30 Or intCount <> 60 Or intCount <> 90) Then

                        varTemp = Shell("cmd.exe /c " & gstrLocalCDrive & "EmproInv\TypeToPrn.bat " & gstrLocalCDrive & "EmproInv\PageFeedSet.txt", AppWinStyle.Hide)
                    Else

                        varTemp = Shell("cmd.exe /c " & gstrLocalCDrive & "EmproInv\TypeToPrn.bat " & gstrLocalCDrive & "EmproInv\PageFeed.txt", AppWinStyle.Hide)
                    End If
                End If
                lblPrint.Text = ""
            Next
        End With
        lblPrint.Text = ""
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        If Err.Number = 53 Then
            FileOpen(1, gstrLocalCDrive & "EmproInv\TypeToPrn.bat", OpenMode.Append)
            PrintLine(1, "Type %1> prn") '& Printer.Port
            FileClose(1)
            GoTo TypeFileNotFoundCreateRetry
        End If
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred

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

End Class