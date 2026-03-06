Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMKTTRN0048
	Inherits System.Windows.Forms.Form
	'-----------------------------------------------------------------------
	' Copyright (c)     :MIND Ltd.
	' Form Name         :frmMKTTRN0048
	' Function Name     :Supplymentry Invoice Cancellation
	' Description       :This form is used to Cancel the Locked Supplymentary Challan
	' Created By        :Davinder singh
	' Created On        :12-April-2006
	'-----------------------------------------------------------------------
    'Modified by    :   Virendra Gupta
    'Modified ON    :   20/05/2011
    'Modified to support MultiUnit functionality
    '-----------------------------------------------------------------------
	Private Enum enumPreInvoiceDetails
		Select_Invoice = 1
		Invoice_No = 2
		Invoice_Date = 3
		LastSupplementary = 4
		SupplementaryDate = 5
		Quantity = 6
		Rate = 7
		New_Rate = 8
		Rate_diff = 9
		TotalPacking = 10
		NewPacking = 11
		NewTotalPacking = 12
		Basic = 13
		NewBasic = 14
		BasicDiff = 15
		TotalCustSuppMaterial = 16
		NewCustSuppMaterial = 17
		NewTotalCustSuppMaterial = 18
		CustSuppMaterial_diff = 19
		ToolCost = 20
		newToolCost = 21
		NewTotalToolCost = 22
		ToolCost_diff = 23
		AccessableValue = 24
		NewAccessableValue = 25
		AccessableValue_Diff = 26
		TotalExciseValue = 27
		NewExciseValue = 28
		NewCVDValue = 29
		NewSADValue = 30
		NewTotalExciseValue = 31
		TotalExciseValueDiff = 32
		TotalEcessValue = 33
		NewEcessValue = 34
		TotalEcssDiff = 35
		SalesTaxValue = 36
		NewSalesTaxValue = 37
		SalesTaxValueDiff = 38
		SSTVlaue = 39
		NewSSTValue = 40
		SSTVlaueDiff = 41
		TotalCurrInvValue = 42
		TotalInvoiceValue = 43
		Flag = 44
	End Enum
	
	Private Enum enumInvoiceSummery
		Rate = 1
		BasicValue = 2
		CustSuppMat = 3
		ToolCost = 4
		AccessableValue = 5
		ExciseValue = 6
		EcssValue = 7
		SalesTaxType = 8
		SSTType = 9
		SummeryInvoiceValue = 10
	End Enum
	
	Private mintIndex As Short
	Private mblnStatus As Boolean
	Private mstrSelInvoices As String 'to store the selected invoice numbers
    Private mstrNotSelInvoices As String ' to store the deselected invoice numbers
    Dim blnClose As Boolean = False
    Dim blnCheck As Boolean = False
 	
    Private Sub CmdChallanNo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdChallanNo.Click
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Create Date         - 13-APR-2006
        'Arguments           - None
        'Return Value        - None
        'Function            - To Display Help From SalesChallan_Dtl
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Dim strHelpString As String
        Dim strChallanNo() As String
        If Trim(txtLocationCode.Text) = "" Then
            Call ConfirmWindow(10239, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO, 100)
            txtLocationCode.Focus()
            Exit Sub
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        strHelpString = "Select DISTINCT Doc_No,Location_Code from SupplementaryInv_hdr where Location_code ='" & txtLocationCode.Text & "'"
        strHelpString = strHelpString & " AND Bill_Flag=1 AND Cancel_flag = 0 and Unit_Code = '" & gstrUNITID & "'"
        strChallanNo = ctlEMPHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelpString, "Supplementary Invoice No", 2)
        If UBound(strChallanNo) < 0 Then
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            Exit Sub
        End If
        If strChallanNo(0) = "0" Then
            MsgBox("No Supplementary Invoice Available To Display", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            txtChallanNo.Text = ""
        Else
            txtChallanNo.Text = strChallanNo(0)
        End If
        txtChallanNo.Focus()
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub CmdGrp_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.cmdGrpAuthorise.ButtonClickEventArgs) Handles CmdGrp.ButtonClick
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Create Date         - 13-APR-2006
        'Arguments           - None
        'Return Value        - None
        'Function            - To Perform the function according to the Button Clicked
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim strResult As String
        Dim objDrCr As prj_DrCrNote.cls_DrCrNote
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_AUTHORIZE  'CANCELLATION OF INVOICE
                If NotValidData() Then Exit Sub
                If ConfirmWindow(10054, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
                    objDrCr = New prj_DrCrNote.cls_DrCrNote(GetServerDate)
                    mP_Connection.BeginTrans()
                    strsql = "Update supplementaryinv_hdr set Cancel_flag=1,Remarks='" & Trim(txtCancelRemarks.Text) & "' where  Unit_Code = '" & gstrUNITID & "' and Doc_no=" & Trim(txtChallanNo.Text)
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    If PostInFin() Then
                        strResult = objDrCr.ReverseARInvoiceDocument(gstrUNITID, Trim(txtChallanNo.Text), mP_User, getDateForDB(GetServerDate()), Trim(txtCancelRemarks.Text), gstrCURRENCYCODE, , gstrCONNECTIONSTRING)
                        strResult = CheckString(strResult)
                    Else
                        strResult = "Y"
                    End If
                    If strResult <> "Y" Then
                        MsgBox(strResult, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        mP_Connection.RollbackTrans()
                        objDrCr = Nothing
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                        Exit Sub
                    Else
                        mP_Connection.CommitTrans()
                        objDrCr = Nothing
                        MsgBox("Challan Cancelled Successfully!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        Call CmdGrp_ButtonClick(CmdGrp, New UCActXCtl.cmdGrpAuthorise.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_REFRESH))
                    End If
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_REFRESH
                Call RefreshCtrls()
                txtLocationCode.Focus()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                blnCheck = True
                Me.Close()
                blnCheck = False
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub CmdLocCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdLocCodeHelp.Click
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Create Date         - 13-APR-2006
        'Arguments           - None
        'Return Value        - None
        'Function            - To Display Help From Location Master
        '*****************************************************************************************
        Dim strLocationCode() As String
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        strLocationCode = Me.ctlEMPHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select DISTINCT s.Location_Code,l.Description from Location_Mst l,SupplementaryInv_hdr s where s.Location_code = l.Location_code and s.Unit_Code = l.Unit_Code and l.Unit_Code = '" & gstrUNITID & "'", "Accounting Locations", 2)
        If UBound(strLocationCode) < 0 Then
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            Exit Sub
        End If
        If strLocationCode(0) = "0" Then
            MsgBox("No Accounting Location Available to Display.") : txtLocationCode.Text = "" : txtLocationCode.Focus() : Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default) : Exit Sub
        Else
            txtLocationCode.Text = strLocationCode(0)
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOk.Click
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Make the visibility of the frame false containing the list view with selected Invoice No's
        ' Datetime      : 15-APR-2006
        '--------------------------------------------------------------------
        On Error GoTo ErrHandler
        fraInvoice.Visible = False
        Call EnableCtrls(True)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdSelectInvoice_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSelectInvoice.Click
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To display the list of selected invoices against which Supp. Inv. has been made
        ' Datetime      : 15-Apr-2006
        '--------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim RsTemp As ADODB.Recordset
        RsTemp = New ADODB.Recordset
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        If mstrSelInvoices = "" Then
            mstrSelInvoices = "Select RefDoc_No from SupplementaryInv_Dtl Where Doc_No = '" & txtChallanNo.Text & "' and Unit_Code = '" & gstrUNITID & "'"
            RsTemp.Open(mstrSelInvoices, mP_Connection)
            mstrSelInvoices = RsTemp.GetString(ADODB.StringFormatEnum.adClipString, , , "|")
            If VB.Right(mstrSelInvoices, 1) = "|" Then
                mstrSelInvoices = VB.Left(mstrSelInvoices, Len(mstrSelInvoices) - 1)
            End If
        End If
        If FillInvoiceNumber() = True Then
            fraInvoice.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(Me.Width) - VB6.PixelsToTwipsX(fraInvoice.Width)) / 2)
            fraInvoice.Width = VB6.TwipsToPixelsX(2390)
            fraInvoice.Visible = True
            fraInvoice.BringToFront()
            cmdOk.Enabled = True
            Call EnableCtrls(True)
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0048_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '-----------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To check the name of the form in the Window Menu
        ' Datetime      : 15-June-2005
        '---------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0048_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '---------------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Set Property NODEFONTBOLD as per Rule book
        ' Datetime      : 15-Apr-2006
        '---------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0048_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '---------------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Keycode of key pressed
        ' Return Value  : Nil
        ' Function      : Unload the form
        ' Datetime      : 14-Apr-2006
        '-----------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler ' Error Handler
        If KeyCode = System.Windows.Forms.Keys.Escape And Shift = 0 Then Me.Close()
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0048_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '------------------------------------------------------------------------------------------------------'
        ' Author        : Davinder singh
        ' Arguments     : Keyascii
        ' Return Value  : Nil
        ' Function      : To make the enterkey behave like Tab
        ' Datetime      : 14-Apr-2006
        '------------------------------------------------------------------------------------------------------'
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{tab}")
        End If
        GoTo EventExitSub 'To prevent the execution of errhandler
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub frmMKTTRN0048_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '------------------------------------------------------------------------------------------------------'
        ' Author        : Davinder singh
        ' Arguments     : Keycode and shift
        ' Return Value  : Nil
        ' Function      : To invoke the onlinehelp associated with form
        ' Datetime      : 14-Apr-2006
        '------------------------------------------------------------------------------------------------------'
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlFormHeader_Click(ctlFormHeader, New System.EventArgs())
        End If
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0048_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '---------------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Set Property AddFormNameToWindowList &
        ' Function FitToClient as per Rule book
        ' Intialise the Controls on form load
        ' Datetime      : 14-Apr-2006
        '-----------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler ' Error Handler
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader.HeaderString())
        FitToClient(Me, FraMain, ctlFormHeader, CmdGrp, 500)
        Call loadicons()
        Call EnableCtrls(False)
        dtpDateFrom.CustomFormat = gstrDateFormat : dtpDateFrom.Format = DateTimePickerFormat.Custom : dtpDateFrom.Value = GetServerDate()
        dtpDateTo.CustomFormat = gstrDateFormat : dtpDateTo.Format = DateTimePickerFormat.Custom : dtpDateTo.Value = GetServerDate()
        sstbInvoiceDtl.Enabled = True : spdPrevInv.Enabled = True : spdInvDetails.Enabled = True
        Call AddHeadersToGrids()
        Call SetWidthofColumnsinGrid()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0048_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '-------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Cancel as Integer
        ' Return Value  : NIL
        ' Function      : RemoveFormNameFromWindowList
        ' Release form Object Memory from Database.
        ' Datetime      : 14-Apr-2006
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Me.Dispose()
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function loadicons() As Object
        '-------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Cancel as Integer
        ' Return Value  : NIL
        ' Function      : To load the Icons on the Buttons and change the Captions of the Buttons.
        ' Datetime      : 14-Apr-2006
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Me.CmdLocCodeHelp.Image = My.Resources.ico111.ToBitmap
        Me.CmdChallanNo.Image = My.Resources.ico111.ToBitmap
        Me.cmdSelectInvoice.Image = My.Resources.ico223.ToBitmap
        CmdGrp.Caption(0) = "Cancel"
        CmdGrp.Caption(1) = "Refresh"
        CmdGrp.Picture(UCActXCtl.clsDeclares.ButtonEnabledEnum.NEW_BUTTON) = My.Resources.resEmpower.ico123.ToBitmap
        CmdGrp.Picture(UCActXCtl.clsDeclares.ButtonEnabledEnum.UPDATE_BUTTON) = My.Resources.resEmpower.ico121.ToBitmap
        Exit Function 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function EnableCtrls(ByVal blnFlag As Boolean) As Object
        '-------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Boolean
        ' Return Value  : NIL
        ' Function      : To Enable/Disable the controls
        ' Datetime      : 13-APR-2006
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        If blnFlag = False Then
            Call EnableControls(False, Me, True)
            txtLocationCode.Enabled = True
            txtChallanNo.Enabled = True
            CmdLocCodeHelp.Enabled = True
            CmdChallanNo.Enabled = True
            txtCancelRemarks.Enabled = True
            txtLocationCode.MaxLength = 4
            txtChallanNo.MaxLength = 9
            txtCancelRemarks.MaxLength = 100
            txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            txtCancelRemarks.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            lblCVD_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            lblExctax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            lblSAD_per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            lblEcssCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            lblCustCodeDes.Text = ""
            lblCustItemDesc.Text = ""
            lblItemCodeDes.Text = ""
            CmdGrp.Enabled(0) = True
            CmdGrp.Enabled(1) = True
        Else
            If txtLocationCode.Enabled Then
                txtLocationCode.Enabled = False
                txtChallanNo.Enabled = False
                txtCancelRemarks.Enabled = False
                txtCancelRemarks.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                CmdLocCodeHelp.Enabled = False
                CmdChallanNo.Enabled = False
                cmdSelectInvoice.Enabled = False
                CmdGrp.Enabled(0) = False
            Else
                txtLocationCode.Enabled = True
                txtChallanNo.Enabled = True
                txtCancelRemarks.Enabled = True
                txtCancelRemarks.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                CmdLocCodeHelp.Enabled = True
                CmdChallanNo.Enabled = True
                cmdSelectInvoice.Enabled = True
                CmdGrp.Enabled(0) = True
                txtCancelRemarks.Focus()
            End If
        End If
        Exit Function 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub txtCancelRemarks_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCancelRemarks.TextChanged
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Create Date         - 17/04/2006
        'Arguments           - NIL
        'Return Value        - NIL
        'Function            - To Replace the single Quote with blank
        '*****************************************************************************************
        On Error GoTo ErrHandler
        txtCancelRemarks.Text = Replace(txtCancelRemarks.Text, "'", "")
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtCancelRemarks_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCancelRemarks.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Create Date         - 17/04/2006
        'Arguments           - None
        'Return Value        - Bollean Values
        'Function            - To Check the PostInFin flag from Sales_Parameter
        '*****************************************************************************************
        On Error GoTo ErrHandler
        If KeyAscii = 39 Then KeyAscii = 0
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtChallanNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtChallanNo.TextChanged
        '-------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Boolean
        ' Return Value  : NIL
        ' Function      : To refresh the whole form by preserving Location Code
        ' Datetime      : 13-APR-2006
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        txtChallanNo.Text = Replace(txtChallanNo.Text, "'", "")
        If Trim(txtChallanNo.Text) = "" Then
            Me.txtLocationCode.Tag = Me.txtLocationCode.Text
            Call RefreshCtrls()
            Me.txtLocationCode.Text = Me.txtLocationCode.Tag
            Me.txtChallanNo.Focus()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtChallanNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtChallanNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '-------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Boolean
        ' Return Value  : NIL
        ' Function      : To Display the Challan No. help on pressing the F1 key
        ' Datetime      : 13-APR-2006
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then CmdChallanNo.PerformClick()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtChallanNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtChallanNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '-------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Boolean
        ' Return Value  : NIL
        ' Function      : To restrict the User from entering the Non numeric value
        ' Datetime      : 13-APR-2006
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        If Not (KeyAscii = 13 Or KeyAscii = System.Windows.Forms.Keys.Back) Then
            If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtChallanNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtChallanNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Create Date         - 14-Apr-2006
        'Arguments           - Cancel
        'Return Value        - None
        'Function            - To check the Validity of entered Challan No.
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Dim strsql As String
        If Me.CmdGrp.GetActiveButton.Text.ToString.ToUpper = "CLOSE" Then
            Exit Sub
        End If
        strsql = "Select DISTINCT Doc_No,Location_Code from SupplementaryInv_hdr where Location_code ='" & Trim(txtLocationCode.Text) & "' AND Doc_No='" & Trim(txtChallanNo.Text) & "'"
        strsql = strsql & " AND Bill_Flag=1 AND Cancel_flag = 0 and Unit_Code = '" & gstrUNITID & "'"
        If Trim(txtChallanNo.Text) <> "" Then
            If Not CheckExistanceOfFieldData(strsql) Then
                MsgBox("Entered Challan No. is Invalid Or doesn't associated with the selected Location ", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                With txtChallanNo
                    .SelectionStart = 0
                    .SelectionLength = Len(.Text)
                End With
                mblnStatus = True
                Cancel = True
            Else
                If Not DisplayDetailsinViewMode() Then
                    Call ConfirmWindow(10414, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    mblnStatus = True
                    Cancel = True
                    GoTo EventExitSub
                End If
                cmdSelectInvoice.Enabled = True
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtLocationCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLocationCode.TextChanged
        '-------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Boolean
        ' Return Value  : NIL
        ' Function      : To refresh the all controls on the form
        ' Datetime      : 13-APR-2006
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        txtLocationCode.Text = Replace(txtLocationCode.Text, "'", "")
        If Trim(txtLocationCode.Text) = "" Then
            Call RefreshCtrls()
            txtChallanNo.Text = ""
            txtLocationCode.Focus()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLocationCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtLocationCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '-------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Boolean
        ' Return Value  : NIL
        ' Function      : To Show the Location Code's help
        ' Datetime      : 13-APR-2006
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then CmdLocCodeHelp.PerformClick()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function CheckExistanceOfFieldData(ByVal pstrSQL As String) As Boolean
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Create Date         - 14/10/2003
        'Arguments           - pstrSQL - Full Query
        'Return Value        - None
        'Function            - To Check Validity Of Field Data Whether it Exists In The Database Or Not
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Dim rsExistData As ClsResultSetDB
        rsExistData = New ClsResultSetDB
        CheckExistanceOfFieldData = False
        rsExistData.GetResult(pstrSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsExistData.GetNoRows > 0 Then
            CheckExistanceOfFieldData = True
        Else
            CheckExistanceOfFieldData = False
        End If
        rsExistData.ResultSetClose()
        rsExistData = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub txtLocationCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLocationCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '-------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Boolean
        ' Return Value  : NIL
        ' Function      : To enter the letters only in Capital
        ' Datetime      : 13-APR-2006
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        If KeyAscii = 39 Then
            KeyAscii = 0
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtLocationCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtLocationCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Create Date         - 14/10/2003
        'Arguments           - Cancel
        'Return Value        - None
        'Function            - To Validate the entered Location code
        '*****************************************************************************************
        On Error GoTo ErrHandler
        If Trim(txtLocationCode.Text) <> "" Then
            If Not CheckExistanceOfFieldData("Select Location_Code from SupplementaryInv_hdr where Location_code='" & Trim(txtLocationCode.Text) & "' and Unit_Code = '" & gstrUNITID & "'") Then
                MsgBox("Entered location Code is Invalid", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                With txtLocationCode
                    .SelectionStart = 0
                    .SelectionLength = Len(.Text)
                End With
                mblnStatus = True
                Cancel = True
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Function NotValidData() As Boolean
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Create Date         - 15 Apr 2006
        'Arguments           - NIL
        'Return Value        - Boolean
        'Function            - To Check Validity Of the entered data before saving it
        '*****************************************************************************************
        On Error GoTo ErrHandler
        NotValidData = False
        If Trim(txtLocationCode.Text) = "" Then
            MsgBox("Location Code can't be left blank", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            NotValidData = True
            txtLocationCode.Focus()
            Exit Function
        End If
        Call txtLocationCode_Validating(txtLocationCode, New System.ComponentModel.CancelEventArgs(False))
        If mblnStatus Then
            NotValidData = True
            mblnStatus = False
            txtLocationCode.Focus()
            Exit Function
        End If
        If Trim(txtChallanNo.Text) = "" Then
            MsgBox("Challan No. can't be left blank", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            NotValidData = True
            txtChallanNo.Focus()
            Exit Function
        End If
        Call txtChallanNo_Validating(txtChallanNo, New System.ComponentModel.CancelEventArgs(False))
        If mblnStatus Then
            NotValidData = True
            mblnStatus = False
            txtChallanNo.Focus()
            Exit Function
        End If
        If Trim(txtCancelRemarks.Text) = "" Then
            MsgBox("Please Enter the Canceling Remarks", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            NotValidData = True
            txtCancelRemarks.Focus()
            Exit Function
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Sub AddHeadersToGrids()
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Create Date         - 15 Apr 2006
        'Arguments           - None
        'Return Value        - None
        'Function            - To Set Header Labels of Both the Grids
        '*****************************************************************************************
        On Error GoTo ErrHandler
        With spdPrevInv
            .MaxRows = 1
            .MaxCols = 44
            .Row = 0
            .Col = enumPreInvoiceDetails.Select_Invoice : .Text = "Selected Invoice"
            .Col = enumPreInvoiceDetails.Invoice_No : .Text = "Invoice No"
            .Col = enumPreInvoiceDetails.Invoice_Date : .Text = "Invoice Date"
            .Col = enumPreInvoiceDetails.LastSupplementary : .Text = "Supp Inv No"
            .Col = enumPreInvoiceDetails.SupplementaryDate : .Text = "Supp Inv Date"
            .Col = enumPreInvoiceDetails.Quantity : .Text = "Quantity"
            .Col = enumPreInvoiceDetails.Rate : .Text = "Rate"
            .Col = enumPreInvoiceDetails.New_Rate : .Text = "New Rate"
            .Col = enumPreInvoiceDetails.Rate_diff : .Text = "Rate diff"
            .Col = enumPreInvoiceDetails.TotalCustSuppMaterial : .Text = "Cust Supp Material Amount"
            .Col = enumPreInvoiceDetails.NewCustSuppMaterial : .Text = "New Cust Supp Material (PerValue)"
            .Col = enumPreInvoiceDetails.NewTotalCustSuppMaterial : .Text = "New Total Cust Supp Material"
            .Col = enumPreInvoiceDetails.CustSuppMaterial_diff : .Text = "Cust Supp Mat.diff"
            .Col = enumPreInvoiceDetails.ToolCost : .Text = "Tool Cost Amount"
            .Col = enumPreInvoiceDetails.newToolCost : .Text = "New Tool Cost (Per Value)"
            .Col = enumPreInvoiceDetails.NewTotalToolCost : .Text = "New Total Tool Cost"
            .Col = enumPreInvoiceDetails.ToolCost_diff : .Text = "Tool Cost diff"
            .Col = enumPreInvoiceDetails.TotalPacking : .Text = "Packing"
            .Col = enumPreInvoiceDetails.NewPacking : .Text = "New Packing (Per Unit)"
            .Col = enumPreInvoiceDetails.NewTotalPacking : .Text = "New Total Packing"
            .Col = enumPreInvoiceDetails.NewCVDValue : .Text = "New CVD Value"
            .Col = enumPreInvoiceDetails.NewSADValue : .Text = "New SAD Value"
            .Col = enumPreInvoiceDetails.NewExciseValue : .Text = "New Excise Value"
            .Col = enumPreInvoiceDetails.TotalExciseValue : .Text = "Total Excise Value"
            .Col = enumPreInvoiceDetails.NewTotalExciseValue : .Text = "New Total Exc. Value"
            .Col = enumPreInvoiceDetails.TotalExciseValueDiff : .Text = "Total Excise Diff"
            .Col = enumPreInvoiceDetails.TotalEcessValue : .Text = "Total ECSS Value"
            .Col = enumPreInvoiceDetails.NewEcessValue : .Text = "New ECSS Value"
            .Col = enumPreInvoiceDetails.TotalEcssDiff : .Text = "Total ECSS Diff"
            .Col = enumPreInvoiceDetails.SalesTaxValue : .Text = "Sales Tax Value"
            .Col = enumPreInvoiceDetails.NewSalesTaxValue : .Text = "New S.Tax Value"
            .Col = enumPreInvoiceDetails.SalesTaxValueDiff : .Text = "S.Tax Diff"
            .Col = enumPreInvoiceDetails.SSTVlaue : .Text = "SSTax Value"
            .Col = enumPreInvoiceDetails.NewSSTValue : .Text = "New SSTax Value"
            .Col = enumPreInvoiceDetails.SSTVlaueDiff : .Text = "SSTax Diff"
            .Col = enumPreInvoiceDetails.Basic : .Text = "Basic Value"
            .Col = enumPreInvoiceDetails.NewBasic : .Text = "New Basic Value"
            .Col = enumPreInvoiceDetails.BasicDiff : .Text = "Basic Diff"
            .Col = enumPreInvoiceDetails.AccessableValue : .Text = "Assessable Value"
            .Col = enumPreInvoiceDetails.NewAccessableValue : .Text = "NEW Assessable Value"
            .Col = enumPreInvoiceDetails.AccessableValue_Diff : .Text = "Assessable Value Diff"
            .Col = enumPreInvoiceDetails.TotalCurrInvValue : .Text = "New Total Value"
            .Col = enumPreInvoiceDetails.TotalInvoiceValue : .Text = "Total Value Diff"
            .Col = enumPreInvoiceDetails.Flag : .Text = "FLAG"
            .Row = 1
            .Row2 = .MaxRows
            .Col = 1
            .Col2 = .MaxCols
            .BlockMode = True
            .Lock = True
            .BlockMode = False
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
        End With
        With spdInvDetails
            .MaxRows = 1
            .MaxCols = 10
            .Row = 0
            .Col = enumInvoiceSummery.Rate : .Text = "Rate"
            .Col = enumInvoiceSummery.BasicValue : .Text = "Basic"
            .Col = enumInvoiceSummery.CustSuppMat : .Text = "Cust Supp Material"
            .Col = enumInvoiceSummery.ToolCost : .Text = "Tool Cost"
            .Col = enumInvoiceSummery.AccessableValue : .Text = "Accessable Value"
            .Col = enumInvoiceSummery.ExciseValue : .Text = "Total Excise Value"
            .Col = enumInvoiceSummery.EcssValue : .Text = "Total Ecss Value "
            .Col = enumInvoiceSummery.SalesTaxType : .Text = "Sales Tax Value"
            .Col = enumInvoiceSummery.SSTType : .Text = "SST Value"
            .Col = enumInvoiceSummery.SummeryInvoiceValue : .Text = "Total Value"
            .Row = 1
            .Row2 = .MaxRows
            .Col = 1
            .Col2 = .MaxCols
            .BlockMode = True
            .Lock = True
            .BlockMode = False
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Sub SetWidthofColumnsinGrid()
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Create Date         - 14/04/2006
        'Arguments           - None
        'Return Value        - None
        'Function            - To Set Width of Columns in Invoice Details Grid on Form Load
        '*****************************************************************************************
        On Error GoTo ErrHandler
        With spdPrevInv
            .set_ColWidth(enumPreInvoiceDetails.Select_Invoice, 0)
            .Col = enumPreInvoiceDetails.Select_Invoice
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.Invoice_No, 0)
            .Col = enumPreInvoiceDetails.Invoice_No
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.Invoice_Date, 0)
            .Col = enumPreInvoiceDetails.Invoice_Date
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.LastSupplementary, 0)
            .Col = enumPreInvoiceDetails.LastSupplementary
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.SupplementaryDate, 0)
            .Col = enumPreInvoiceDetails.SupplementaryDate
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.Quantity, 0)
            .Col = enumPreInvoiceDetails.Quantity
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.Rate, 0)
            .Col = enumPreInvoiceDetails.Rate
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.New_Rate, 1100)
            .set_ColWidth(enumPreInvoiceDetails.Rate_diff, 0)
            .Col = enumPreInvoiceDetails.Rate_diff
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.TotalPacking, 0)
            .Col = enumPreInvoiceDetails.TotalPacking
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewPacking, 1000)
            .set_ColWidth(enumPreInvoiceDetails.NewTotalPacking, 0)
            .Col = enumPreInvoiceDetails.NewTotalPacking
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.Basic, 0)
            .Col = enumPreInvoiceDetails.Basic
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewBasic, 0)
            .Col = enumPreInvoiceDetails.NewBasic
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.BasicDiff, 0)
            .Col = enumPreInvoiceDetails.BasicDiff
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.TotalCustSuppMaterial, 0)
            .Col = enumPreInvoiceDetails.TotalCustSuppMaterial
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewCustSuppMaterial, 1375)
            .set_ColWidth(enumPreInvoiceDetails.NewTotalCustSuppMaterial, 0)
            .Col = enumPreInvoiceDetails.NewTotalCustSuppMaterial
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.CustSuppMaterial_diff, 0)
            .Col = enumPreInvoiceDetails.CustSuppMaterial_diff
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.ToolCost, 0)
            .Col = enumPreInvoiceDetails.ToolCost
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.newToolCost, 1100)
            .set_ColWidth(enumPreInvoiceDetails.NewTotalToolCost, 0)
            .Col = enumPreInvoiceDetails.NewTotalToolCost
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.ToolCost_diff, 0)
            .Col = enumPreInvoiceDetails.ToolCost_diff
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.AccessableValue, 0)
            .Col = enumPreInvoiceDetails.AccessableValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewAccessableValue, 0)
            .Col = enumPreInvoiceDetails.NewAccessableValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.AccessableValue_Diff, 0)
            .Col = enumPreInvoiceDetails.AccessableValue_Diff
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.TotalExciseValue, 0)
            .Col = enumPreInvoiceDetails.TotalExciseValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewExciseValue, 0)
            .Col = enumPreInvoiceDetails.NewExciseValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.TotalEcessValue, 0)
            .Col = enumPreInvoiceDetails.TotalEcessValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewEcessValue, 0)
            .Col = enumPreInvoiceDetails.NewEcessValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.TotalEcssDiff, 0)
            .Col = enumPreInvoiceDetails.TotalEcssDiff
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewCVDValue, 0)
            .Col = enumPreInvoiceDetails.NewCVDValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewSADValue, 0)
            .Col = enumPreInvoiceDetails.NewSADValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewTotalExciseValue, 0)
            .Col = enumPreInvoiceDetails.NewTotalExciseValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.TotalExciseValueDiff, 0)
            .Col = enumPreInvoiceDetails.TotalExciseValueDiff
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.SalesTaxValue, 0)
            .Col = enumPreInvoiceDetails.SalesTaxValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewSalesTaxValue, 0)
            .Col = enumPreInvoiceDetails.NewSalesTaxValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.SalesTaxValueDiff, 0)
            .Col = enumPreInvoiceDetails.SalesTaxValueDiff
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.SSTVlaue, 0)
            .Col = enumPreInvoiceDetails.SSTVlaue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewSSTValue, 0)
            .Col = enumPreInvoiceDetails.NewSSTValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.SSTVlaueDiff, 0)
            .Col = enumPreInvoiceDetails.SSTVlaueDiff
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.TotalCurrInvValue, 0)
            .Col = enumPreInvoiceDetails.TotalCurrInvValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.TotalInvoiceValue, 0)
            .Col = enumPreInvoiceDetails.TotalInvoiceValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.Flag, 0)
            .Col = enumPreInvoiceDetails.Flag
            .ColHidden = True
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function DisplayDetailsinViewMode() As Boolean
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Create Date         - 15 Apr 2006
        'Arguments           - None
        'Return Value        - Boolean
        'Function            - To Display invoice Details
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Dim strsqlHdr As String
        Dim strSqlDtl As String
        Dim intmaxLoop As Short
        Dim intLoopCounter As Short
        Dim rsSupplementaryHdr As ClsResultSetDB
        Dim rsSupplementaryDtl As New ClsResultSetDB
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        DisplayDetailsinViewMode = True
        rsSupplementaryHdr = New ClsResultSetDB
        strsqlHdr = "select S.Location_Code,S.Account_Code,S.Cust_name,S.Cust_Ref,S.Amendment_No,S.Doc_No,S.Invoice_DateFrom,S.Invoice_DateTo,S.Item_Code,"
        strsqlHdr = strsqlHdr & "S.Cust_Item_Code,S.Currency_Code,S.Rate,S.Basic_Amount,S.Accessible_amount,S.Excise_type,S.CVD_type,S.SAD_type,S.Excise_per,"
        strsqlHdr = strsqlHdr & "S.CVD_per,S.SVD_per,S.TotalExciseAmount,S.CustMtrl_Amount,S.ToolCost_amount,S.SalesTax_Type,S.SalesTax_Per,S.Sales_Tax_Amount,"
        strsqlHdr = strsqlHdr & "S.Surcharge_salesTaxType,S.Surcharge_SalesTax_Per,S.Surcharge_Sales_Tax_Amount,S.Total_amount,S.SuppInv_Remarks,"
        strsqlHdr = strsqlHdr & "S.Remarks "
        strsqlHdr = strsqlHdr & ",S.Ecess_Type,S.Ecess_per,S.Ecess_Amount,I.Description,C.Drg_desc"
        strsqlHdr = strsqlHdr & " from supplementaryINV_hdr S, Item_mst I, CustItem_Mst C where S.Location_code = '" & Trim(txtLocationCode.Text) & "' and S.Doc_no = "
        strsqlHdr = strsqlHdr & Trim(txtChallanNo.Text) & "  and S.Unit_Code = '" & gstrUNITID & "' AND S.Item_code=I.Item_Code AND S.Unit_Code=I.Unit_Code and S.Account_Code=C.Account_Code and S.Unit_Code=C.Unit_Code AND S.Cust_Item_Code=C.Cust_drgno"
        rsSupplementaryHdr.GetResult(strsqlHdr, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSupplementaryHdr.GetNoRows > 0 Then
            spdInvDetails.MaxRows = 1
            spdPrevInv.MaxRows = 1
            Call SetCellTypeofGrids()
            dtpDateFrom.Value = rsSupplementaryHdr.GetValue("Invoice_DateFrom")
            dtpDateTo.Value = rsSupplementaryHdr.GetValue("Invoice_DateTo")
            txtCustCode.Text = rsSupplementaryHdr.GetValue("Account_code")
            lblCustCodeDes.Text = rsSupplementaryHdr.GetValue("Cust_Name")
            lblCustItemDesc.Text = rsSupplementaryHdr.GetValue("Drg_desc")
            lblItemCodeDes.Text = rsSupplementaryHdr.GetValue("Description")
            txtCustPartCode.Text = rsSupplementaryHdr.GetValue("Cust_item_code")
            txtItemCode.Text = rsSupplementaryHdr.GetValue("item_code")
            txtRefNo.Text = rsSupplementaryHdr.GetValue("Cust_ref")
            txtAmendment.Text = rsSupplementaryHdr.GetValue("Amendment_no")
            lblCurrencyDes.Text = rsSupplementaryHdr.GetValue("Currency_code")
            txtCustRefRemarks.Text = rsSupplementaryHdr.GetValue("SuppInv_Remarks")
            txtRemarks.Text = rsSupplementaryHdr.GetValue("Remarks")
            txtCVDCode.Text = rsSupplementaryHdr.GetValue("CVD_Type")
            lblCVD_Per.Text = rsSupplementaryHdr.GetValue("CVD_Per")
            txtSADCode.Text = rsSupplementaryHdr.GetValue("SAD_Type")
            lblSAD_per.Text = rsSupplementaryHdr.GetValue("SVD_Per")
            txtExciseTaxType.Text = rsSupplementaryHdr.GetValue("Excise_Type")
            lblExctax_Per.Text = rsSupplementaryHdr.GetValue("Excise_Per")
            Me.txtEcssCode.Text = rsSupplementaryHdr.GetValue("Ecess_Type")
            Me.lblEcssCode.Text = rsSupplementaryHdr.GetValue("Ecess_Per")
            txtSaleTaxType.Text = rsSupplementaryHdr.GetValue("SalesTax_type")
            lblSaltax_Per.Text = rsSupplementaryHdr.GetValue("SalesTax_Per")
            txtSurchargeTaxType.Text = rsSupplementaryHdr.GetValue("Surcharge_SalesTaxtype")
            lblSurcharge_Per.Text = rsSupplementaryHdr.GetValue("Surcharge_SalesTax_Per")
            With spdInvDetails
                Call .SetText(enumInvoiceSummery.Rate, 1, rsSupplementaryHdr.GetValue("Rate"))
                Call .SetText(enumInvoiceSummery.BasicValue, 1, rsSupplementaryHdr.GetValue("Basic_amount"))
                Call .SetText(enumInvoiceSummery.CustSuppMat, 1, rsSupplementaryHdr.GetValue("CustMtrl_amount"))
                Call .SetText(enumInvoiceSummery.ToolCost, 1, rsSupplementaryHdr.GetValue("ToolCost_amount"))
                Call .SetText(enumInvoiceSummery.AccessableValue, 1, rsSupplementaryHdr.GetValue("Accessible_amount"))
                Call .SetText(enumInvoiceSummery.ExciseValue, 1, rsSupplementaryHdr.GetValue("TotalExciseAmount"))
                Call .SetText(enumInvoiceSummery.EcssValue, 1, rsSupplementaryHdr.GetValue("Ecess_amount"))
                Call .SetText(enumInvoiceSummery.SalesTaxType, 1, rsSupplementaryHdr.GetValue("Sales_tax_Amount"))
                Call .SetText(enumInvoiceSummery.SSTType, 1, rsSupplementaryHdr.GetValue("Surcharge_Sales_tax_amount"))
                Call .SetText(enumInvoiceSummery.SummeryInvoiceValue, 1, rsSupplementaryHdr.GetValue("Total_amount"))
                .Enabled = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = 1
                .Row2 = .MaxRows
                .BlockMode = True
                .Lock = True
                .BlockMode = False
            End With
        Else
            DisplayDetailsinViewMode = False
        End If
        rsSupplementaryHdr.ResultSetClose()
        If Len(Trim(txtCustCode.Text)) > 0 Then
            strsqlHdr = "SELECT isnull(Bill_Address1,'') as Bill_Address1,isnull(Bill_Address2,'') as Bill_Address2,isnull(Bill_City,'') as Bill_City,isnull(Bill_Pin,'') as Bill_Pin from Customer_Mst where Customer_code ='" & Trim(txtCustCode.Text) & "' and Unit_Code = '" & gstrUNITID & "'"
            rsSupplementaryHdr = New ClsResultSetDB
            rsSupplementaryHdr.GetResult(strsqlHdr)
            If rsSupplementaryHdr.GetNoRows > 0 Then
                lblAddressDes.Text = IIf(Trim(rsSupplementaryHdr.GetValue("Bill_Address1")) = "", "", rsSupplementaryHdr.GetValue("Bill_Address1"))
                lblAddressDes.Text = lblAddressDes.Text & IIf(Trim(rsSupplementaryHdr.GetValue("Bill_Address2")) = "", "", " ," & rsSupplementaryHdr.GetValue("Bill_Address2"))
                lblAddressDes.Text = lblAddressDes.Text & IIf(Trim(rsSupplementaryHdr.GetValue("Bill_City")) = "", "", " ," & rsSupplementaryHdr.GetValue("Bill_City"))
                lblAddressDes.Text = lblAddressDes.Text & IIf(Trim(rsSupplementaryHdr.GetValue("Bill_Pin")) = "", "", " - " & rsSupplementaryHdr.GetValue("Bill_Pin"))
            End If
            rsSupplementaryHdr.ResultSetClose()
            rsSupplementaryHdr = Nothing
        End If
        strSqlDtl = "Select SelectInvoice = 1,RefDoc_No,RefDoc_Date,LastSupplementary,SuppInvdate,Item_code,Cust_Item_Code,PrevRate,Rate,Rate_diff,Quantity,PrevPacking_Amount,Packing_Per,"
        strSqlDtl = strSqlDtl & "Packing_amountDiff,PrevBasic_Amount,Basic_Amount,Basic_AmountDiff,PrevAccessible_amount,Accessible_amount,"
        strSqlDtl = strSqlDtl & "Accessible_amountDiff,PrevTotalExciseAmount,CVD_Amount,SAD_amount,Excise_amount,TotalExciseAmount,TotalExciseAmountDiff,PrevCustMtrl_Amount,"
        strSqlDtl = strSqlDtl & "CustMtrl_Amount,TotalCustMtrl_Amount,CustMtrl_AmountDiff,PrevToolCost_amount,ToolCost_amount,"
        strSqlDtl = strSqlDtl & "TotalToolCost_amount,ToolCost_amountDiff,PrevSales_Tax_Amount,Sales_Tax_Amount,Sales_Tax_AmountDiff,"
        strSqlDtl = strSqlDtl & "PrevSSTAmount,SST_Amount,SST_AmountDiff,total_amount,total_amountDiff "
        strSqlDtl = strSqlDtl & ",ECESS_Amount,PrevECESS_Amount,ECESS_Amount_Diff"
        strSqlDtl = strSqlDtl & " from SupplementaryInv_dtl where Unit_Code = '" & gstrUNITID & "' and Location_code ='"
        strSqlDtl = strSqlDtl & Trim(txtLocationCode.Text) & "' and Doc_no = " & Trim(txtChallanNo.Text) & vbCrLf
        strSqlDtl = strSqlDtl & "UNION Select SelectInvoice = 0,RefDoc_No,RefDoc_Date,LastSupplementary,SuppInvdate,Item_code,Cust_Item_Code,PrevRate,Rate,Rate_diff,Quantity,PrevPacking_Amount,Packing_Per,"
        strSqlDtl = strSqlDtl & "Packing_amountDiff,PrevBasic_Amount,Basic_Amount,Basic_AmountDiff,PrevAccessible_amount,Accessible_amount,"
        strSqlDtl = strSqlDtl & "Accessible_amountDiff,PrevTotalExciseAmount,CVD_Amount,SAD_amount,Excise_amount,TotalExciseAmount,TotalExciseAmountDiff,PrevCustMtrl_Amount,"
        strSqlDtl = strSqlDtl & "CustMtrl_Amount,TotalCustMtrl_Amount,CustMtrl_AmountDiff,PrevToolCost_amount,ToolCost_amount,"
        strSqlDtl = strSqlDtl & "TotalToolCost_amount,ToolCost_amountDiff,PrevSales_Tax_Amount,Sales_Tax_Amount,Sales_Tax_AmountDiff,"
        strSqlDtl = strSqlDtl & "PrevSSTAmount,SST_Amount,SST_AmountDiff,Total_amount,total_amountDiff "
        strSqlDtl = strSqlDtl & ",ECESS_Amount,PrevECESS_Amount,ECESS_Amount_Diff"
        strSqlDtl = strSqlDtl & " from SuppCreditAdvise_Dtl where Unit_Code = '" & gstrUNITID & "' and Location_code ='"
        strSqlDtl = strSqlDtl & Trim(txtLocationCode.Text) & "' and Doc_no = " & Trim(txtChallanNo.Text)
        strSqlDtl = strSqlDtl & " order by SelectInvoice DESC "
        rsSupplementaryDtl = New ClsResultSetDB
        rsSupplementaryDtl.GetResult(strSqlDtl)
        Call GetInvoiceNumbers(rsSupplementaryDtl)
        cmdSelectInvoice.Enabled = True
        If rsSupplementaryDtl.GetNoRows > 0 Then
            intmaxLoop = rsSupplementaryDtl.GetNoRows
            rsSupplementaryDtl.MoveFirst()
            With spdPrevInv
                spdPrevInv.MaxRows = 1
                Call SetMaxLengthofGrid(4)
                For intLoopCounter = 1 To 1
                    If rsSupplementaryDtl.GetValue("SelectInvoice") = 1 Then
                        .Col = enumPreInvoiceDetails.Select_Invoice
                        .Row = intLoopCounter
                        .Value = System.Windows.Forms.CheckState.Checked
                    Else
                        .Col = enumPreInvoiceDetails.Select_Invoice
                        .Row = intLoopCounter
                        .Value = System.Windows.Forms.CheckState.Unchecked
                    End If
                    Call .SetText(enumPreInvoiceDetails.Invoice_No, intLoopCounter, rsSupplementaryDtl.GetValue("RefDoc_no"))
                    Call .SetText(enumPreInvoiceDetails.Invoice_Date, intLoopCounter, rsSupplementaryDtl.GetValue("RefDoc_Date"))
                    Call .SetText(enumPreInvoiceDetails.LastSupplementary, intLoopCounter, rsSupplementaryDtl.GetValue("LastSupplementary"))
                    Call .SetText(enumPreInvoiceDetails.SupplementaryDate, intLoopCounter, rsSupplementaryDtl.GetValue("SuppInvdate"))
                    Call .SetText(enumPreInvoiceDetails.Quantity, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Quantity")))
                    Call .SetText(enumPreInvoiceDetails.Rate, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevRate")))
                    Call .SetText(enumPreInvoiceDetails.New_Rate, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Rate")))
                    Call .SetText(enumPreInvoiceDetails.Rate_diff, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Rate_diff")))
                    Call .SetText(enumPreInvoiceDetails.TotalPacking, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevPacking_amount")))
                    Call .SetText(enumPreInvoiceDetails.NewPacking, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Packing_Per")))
                    Call .SetText(enumPreInvoiceDetails.NewTotalPacking, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Packing_amountdiff")))
                    Call .SetText(enumPreInvoiceDetails.Basic, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevBasic_amount")))
                    Call .SetText(enumPreInvoiceDetails.NewBasic, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Basic_amount")))
                    Call .SetText(enumPreInvoiceDetails.BasicDiff, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Basic_amountDiff")))
                    Call .SetText(enumPreInvoiceDetails.TotalCustSuppMaterial, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevCustMtrl_amount")))
                    Call .SetText(enumPreInvoiceDetails.NewCustSuppMaterial, intLoopCounter, Val(rsSupplementaryDtl.GetValue("CustMtrl_amount")))
                    Call .SetText(enumPreInvoiceDetails.NewTotalCustSuppMaterial, intLoopCounter, Val(rsSupplementaryDtl.GetValue("TotalCustMtrl_amount")))
                    Call .SetText(enumPreInvoiceDetails.CustSuppMaterial_diff, intLoopCounter, Val(rsSupplementaryDtl.GetValue("CustMtrl_amountDiff")))
                    Call .SetText(enumPreInvoiceDetails.ToolCost, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevToolCost_amount")))
                    Call .SetText(enumPreInvoiceDetails.newToolCost, intLoopCounter, Val(rsSupplementaryDtl.GetValue("ToolCost_amount")))
                    Call .SetText(enumPreInvoiceDetails.NewTotalToolCost, intLoopCounter, Val(rsSupplementaryDtl.GetValue("TotalToolCost_amount")))
                    Call .SetText(enumPreInvoiceDetails.ToolCost_diff, intLoopCounter, Val(rsSupplementaryDtl.GetValue("ToolCost_amountDiff")))
                    Call .SetText(enumPreInvoiceDetails.AccessableValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevAccessible_amount")))
                    Call .SetText(enumPreInvoiceDetails.NewAccessableValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Accessible_amount")))
                    Call .SetText(enumPreInvoiceDetails.AccessableValue_Diff, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Accessible_amountDiff")))
                    Call .SetText(enumPreInvoiceDetails.TotalExciseValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevTotalExciseAmount")))
                    Call .SetText(enumPreInvoiceDetails.NewCVDValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("CVD_Amount")))
                    Call .SetText(enumPreInvoiceDetails.NewSADValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("SAD_Amount")))
                    Call .SetText(enumPreInvoiceDetails.NewExciseValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Excise_amount")))
                    Call .SetText(enumPreInvoiceDetails.NewTotalExciseValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("TotalExciseamount")))
                    Call .SetText(enumPreInvoiceDetails.TotalExciseValueDiff, intLoopCounter, Val(rsSupplementaryDtl.GetValue("TotalExciseamountDiff")))
                    Call .SetText(enumPreInvoiceDetails.TotalEcessValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevECESS_Amount")))
                    Call .SetText(enumPreInvoiceDetails.NewEcessValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("ECESS_Amount")))
                    Call .SetText(enumPreInvoiceDetails.TotalEcssDiff, intLoopCounter, Val(rsSupplementaryDtl.GetValue("ECESS_Amount_Diff")))
                    Call .SetText(enumPreInvoiceDetails.SalesTaxValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevSales_Tax_amount")))
                    Call .SetText(enumPreInvoiceDetails.NewSalesTaxValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Sales_Tax_amount")))
                    Call .SetText(enumPreInvoiceDetails.SalesTaxValueDiff, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Sales_Tax_amountDiff")))
                    Call .SetText(enumPreInvoiceDetails.SSTVlaue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevSSTAmount")))
                    Call .SetText(enumPreInvoiceDetails.NewSSTValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("SST_Amount")))
                    Call .SetText(enumPreInvoiceDetails.SSTVlaueDiff, intLoopCounter, Val(rsSupplementaryDtl.GetValue("SST_AmountDiff")))
                    Call .SetText(enumPreInvoiceDetails.TotalCurrInvValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("total_amount")))
                    Call .SetText(enumPreInvoiceDetails.TotalInvoiceValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("total_amountDiff")))
                    If rsSupplementaryDtl.GetValue("SelectInvoice") = 1 Then
                        Call .SetText(enumPreInvoiceDetails.Flag, intLoopCounter, "S")
                    Else
                        Call .SetText(enumPreInvoiceDetails.Flag, intLoopCounter, "")
                    End If
                    rsSupplementaryDtl.MoveNext()
                Next
                .Enabled = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = 1
                .BlockMode = True
                .Lock = True
                .BlockMode = False
            End With
        Else
            DisplayDetailsinViewMode = False
        End If
        rsSupplementaryDtl.ResultSetClose()
        rsSupplementaryDtl = Nothing
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function GetInvoiceNumbers(ByVal RsInvNos As ClsResultSetDB) As Object
        '**************************************************************************
        'Created by   :   Davinder Singh
        'DateTime     :   15 Apr 2006
        'Description  :   Function to fill the different string for selected and not
        'Arguments    :   RsInvNos :- An recordset of data containing the invoices
        '**************************************************************************
        On Error GoTo ErrHandler
        mstrSelInvoices = ""
        mstrNotSelInvoices = ""
        With RsInvNos
            .MoveFirst()
            Do While Not .EOFRecord
                'if the invoice is selected
                If .GetValue("SelectInvoice") = 1 Then
                    mstrSelInvoices = mstrSelInvoices & "|" & .GetValue("RefDoc_No")
                    'if the invoice is not selected
                ElseIf .GetValue("SelectInvoice") = 0 Then
                    mstrNotSelInvoices = mstrNotSelInvoices & "|" & .GetValue("RefDoc_No")
                End If
                .MoveNext()
            Loop
        End With
        'remove the first character
        If VB.Left(mstrSelInvoices, 1) = "|" Then
            mstrSelInvoices = VB.Right(mstrSelInvoices, Len(mstrSelInvoices) - 1)
        End If
        'remove the first character
        If VB.Left(mstrNotSelInvoices, 1) = "|" Then
            mstrNotSelInvoices = VB.Right(mstrNotSelInvoices, Len(mstrNotSelInvoices) - 1)
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Sub SetCellTypeofGrids()
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Create Date         - 15 Apr 2006
        'Arguments           - None
        'Return Value        - None
        'Function            - To Set Cell Type in Grids
        '*****************************************************************************************
        On Error GoTo ErrHandler
        With spdPrevInv
            .Row = 1
            .Row2 = .MaxRows
            .Col = enumPreInvoiceDetails.Select_Invoice
            .Col2 = enumPreInvoiceDetails.Select_Invoice
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
            .TypeCheckCenter = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.Invoice_No
            .Col2 = enumPreInvoiceDetails.Invoice_No
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.Invoice_Date
            .Col2 = enumPreInvoiceDetails.Invoice_Date
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.LastSupplementary
            .Col2 = enumPreInvoiceDetails.LastSupplementary
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.SupplementaryDate
            .Col2 = enumPreInvoiceDetails.SupplementaryDate
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.Quantity
            .Col2 = enumPreInvoiceDetails.Quantity
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.Rate
            .Col2 = enumPreInvoiceDetails.Rate
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.New_Rate
            .Col2 = enumPreInvoiceDetails.New_Rate
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .BlockMode = False
            .Col = enumPreInvoiceDetails.Rate_diff
            .Col2 = enumPreInvoiceDetails.Rate_diff
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.TotalCustSuppMaterial
            .Col2 = enumPreInvoiceDetails.TotalCustSuppMaterial
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewCustSuppMaterial
            .Col2 = enumPreInvoiceDetails.NewCustSuppMaterial
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewTotalCustSuppMaterial
            .Col2 = enumPreInvoiceDetails.NewTotalCustSuppMaterial
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.CustSuppMaterial_diff
            .Col2 = enumPreInvoiceDetails.CustSuppMaterial_diff
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.ToolCost
            .Col2 = enumPreInvoiceDetails.ToolCost
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.newToolCost
            .Col2 = enumPreInvoiceDetails.newToolCost
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewTotalToolCost
            .Col2 = enumPreInvoiceDetails.NewTotalToolCost
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.ToolCost_diff
            .Col2 = enumPreInvoiceDetails.ToolCost_diff
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.TotalPacking
            .Col2 = enumPreInvoiceDetails.TotalPacking
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewPacking
            .Col2 = enumPreInvoiceDetails.NewPacking
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewTotalPacking
            .Col2 = enumPreInvoiceDetails.NewTotalPacking
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewCVDValue
            .Col2 = enumPreInvoiceDetails.NewCVDValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewSADValue
            .Col2 = enumPreInvoiceDetails.NewSADValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewExciseValue
            .Col2 = enumPreInvoiceDetails.NewExciseValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.TotalExciseValue
            .Col2 = enumPreInvoiceDetails.TotalExciseValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewTotalExciseValue
            .Col2 = enumPreInvoiceDetails.NewTotalExciseValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.TotalExciseValueDiff
            .Col2 = enumPreInvoiceDetails.TotalExciseValueDiff
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.TotalEcessValue
            .Col2 = enumPreInvoiceDetails.TotalEcessValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Value = 0.0#
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewEcessValue
            .Col2 = enumPreInvoiceDetails.NewEcessValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Value = 0.0#
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.TotalEcssDiff
            .Col2 = enumPreInvoiceDetails.TotalEcssDiff
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Value = 0.0#
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.SalesTaxValue
            .Col2 = enumPreInvoiceDetails.SalesTaxValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewSalesTaxValue
            .Col2 = enumPreInvoiceDetails.NewSalesTaxValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.SalesTaxValueDiff
            .Col2 = enumPreInvoiceDetails.SalesTaxValueDiff
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.SSTVlaue
            .Col2 = enumPreInvoiceDetails.SSTVlaue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewSSTValue
            .Col2 = enumPreInvoiceDetails.NewSSTValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.SSTVlaueDiff
            .Col2 = enumPreInvoiceDetails.SSTVlaueDiff
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.Basic
            .Col2 = enumPreInvoiceDetails.Basic
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewBasic
            .Col2 = enumPreInvoiceDetails.NewBasic
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.BasicDiff
            .Col2 = enumPreInvoiceDetails.BasicDiff
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.AccessableValue
            .Col2 = enumPreInvoiceDetails.AccessableValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewAccessableValue
            .Col2 = enumPreInvoiceDetails.NewAccessableValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.AccessableValue_Diff
            .Col2 = enumPreInvoiceDetails.AccessableValue_Diff
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .BlockMode = False
            .Col = enumPreInvoiceDetails.TotalCurrInvValue
            .Col2 = enumPreInvoiceDetails.TotalCurrInvValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .BlockMode = False
            .Col = enumPreInvoiceDetails.TotalInvoiceValue
            .Col2 = enumPreInvoiceDetails.TotalInvoiceValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .BlockMode = False
            .Col = enumPreInvoiceDetails.Flag
            .Col2 = enumPreInvoiceDetails.Flag
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .BlockMode = False
        End With
        With spdInvDetails
            .MaxRows = 1
            .Row = 1
            .Row2 = .MaxRows
            .Col = enumInvoiceSummery.Rate
            .Col2 = enumInvoiceSummery.Rate
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumInvoiceSummery.BasicValue
            .Col2 = enumInvoiceSummery.BasicValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumInvoiceSummery.CustSuppMat
            .Col2 = enumInvoiceSummery.CustSuppMat
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumInvoiceSummery.ToolCost
            .Col2 = enumInvoiceSummery.ToolCost
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumInvoiceSummery.AccessableValue
            .Col2 = enumInvoiceSummery.AccessableValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumInvoiceSummery.ExciseValue
            .Col2 = enumInvoiceSummery.ExciseValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumInvoiceSummery.EcssValue
            .Col2 = enumInvoiceSummery.EcssValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumInvoiceSummery.SalesTaxType
            .Col2 = enumInvoiceSummery.SalesTaxType
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumInvoiceSummery.SSTType
            .Col2 = enumInvoiceSummery.SSTType
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumInvoiceSummery.SummeryInvoiceValue
            .Col2 = enumInvoiceSummery.SummeryInvoiceValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function SetMaxLengthofGrid(ByRef pintMaxDecimalPlaces As Object) As Object
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Create Date         - 14 Apr 2006
        'Arguments           - None
        'Return Value        - None
        'Function            - To Set Max Length of Grid
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Dim strMin As String
        Dim strMax As String
        Dim intLoopCounter As Short
        If pintMaxDecimalPlaces < 2 Then
            pintMaxDecimalPlaces = 2
        End If
        strMin = "0." : strMax = "99999999999999."
        For intLoopCounter = 1 To pintMaxDecimalPlaces
            strMin = strMin & "0"
            strMax = strMax & "9"
        Next
        With spdPrevInv
            .Row = 1
            .Row2 = .MaxRows
            .Col = enumPreInvoiceDetails.New_Rate
            .Col2 = enumPreInvoiceDetails.New_Rate
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = pintMaxDecimalPlaces
            .TypeFloatMin = strMin
            .TypeFloatMax = strMax
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewCustSuppMaterial
            .Col2 = enumPreInvoiceDetails.NewCustSuppMaterial
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = pintMaxDecimalPlaces
            .TypeFloatMin = strMin
            .TypeFloatMax = strMax
            .BlockMode = False
            .Col = enumPreInvoiceDetails.newToolCost
            .Col2 = enumPreInvoiceDetails.newToolCost
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = pintMaxDecimalPlaces
            .TypeFloatMin = strMin
            .TypeFloatMax = strMax
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewPacking
            .Col2 = enumPreInvoiceDetails.NewPacking
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = pintMaxDecimalPlaces
            .TypeFloatMin = strMin
            .TypeFloatMax = strMax
            .BlockMode = False
        End With
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function RefreshCtrls() As Object
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Create Date         - 14 Apr 2006
        'Arguments           - None
        'Return Value        - None
        'Function            - To Reset the controls on the form
        '*****************************************************************************************
        On Error GoTo ErrHandler
        fraInvoice.Visible = False
        Call EnableCtrls(False)
        lblCVD_Per.Text = "0.00"
        lblExctax_Per.Text = "0.00"
        lblSurcharge_Per.Text = "0.00"
        lblSAD_per.Text = "0.00"
        lblSaltax_Per.Text = "0.00"
        lblEcssCode.Text = "0.00"
        lblAddressDes.Text = ""
        lblCurrencyDes.Text = ""
        cmdSelectInvoice.Enabled = False
        dtpDateFrom.Value = GetServerDate()
        dtpDateTo.Value = GetServerDate()
        With spdInvDetails
            .Enabled = True
            .MaxRows = 0
            .MaxRows = 1
        End With
        With spdPrevInv
            .Enabled = True
            .MaxRows = 0
            .MaxRows = 1
        End With
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function FillInvoiceNumber() As Boolean
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Create Date         - 17 Apr 2006
        'Arguments           - None
        'Return Value        - Bollean Values
        'Function            - To display invoice numbers in list and return True if list count > 0
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Dim intMaxRow As Short
        Dim intCounter As Short
        Dim strSelInvoices() As String
        lstInv.Items.Clear()
        lstInv.Columns.Item(0).Width = VB6.TwipsToPixelsX(2000)
        lstInv.Columns.Item(1).Width = VB6.TwipsToPixelsX(1000)
        lstInv.View = System.Windows.Forms.View.Details
        strSelInvoices = Split(mstrSelInvoices, "|")
        intMaxRow = UBound(strSelInvoices) + 1
        lstInv.Columns.Item(1).Width = 0
        For intCounter = 0 To intMaxRow - 1
            If strSelInvoices(intCounter) <> "" Then
                lstInv.Items.Insert(intCounter, strSelInvoices(intCounter))
                lstInv.Items.Item(intCounter).Checked = True
            End If
        Next intCounter
        lstInv.GridLines = True
        lstInv.CheckBoxes = False
        FillInvoiceNumber = True
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function PostInFin() As Boolean
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Create Date         - 17/04/2006
        'Arguments           - None
        'Return Value        - Bollean Values
        'Function            - To Check the PostInFin flag from Sales_Parameter
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Dim Rs As ClsResultSetDB
        Rs = New ClsResultSetDB
        PostInFin = False
        Rs.GetResult("Select PostInFin from Sales_Parameter where Unit_Code = '" & gstrUNITID & "'")
        If Rs.GetValue("PostInFin") = "True" Then
            PostInFin = True
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function CheckString(ByVal pstrResult As String) As String
        '--------------------------------------------------------------------------------------
        'Name       :   CheckString
        'Type       :   Function
        'Author     :   Davinder Singh
        'Arguments  :   pstrResult As String
        'Return     :   String
        'Date Time  :   16 Apr 2006
        'Purpose    :   To get Result return from COM (In case of error it return message)
        '---------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strMessage() As String
        If VB.Left(pstrResult, 1) = "N" Then
            strMessage = SplitIntoColumns(pstrResult)
            If VB.Right(strMessage(1), 1) = "¦" Then
                CheckString = VB.Left(strMessage(1), Len(strMessage(1)) - 1)
            Else
                CheckString = strMessage(1)
            End If
        ElseIf Trim(pstrResult) = "" Then
            CheckString = "N"
        Else
            CheckString = "Y"
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub ctlFormHeader_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader.Click
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To display the form's help
        ' Datetime      : 15-Apr-2006
        '--------------------------------------------------------------------
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
End Class