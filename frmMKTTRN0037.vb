Option Strict Off
Option Explicit On
Imports System.IO

Friend Class frmMKTTRN0037
	Inherits System.Windows.Forms.Form
	'---------------------------------------------------------------------------
	'Copyright          :   MIND Ltd.
	'Form Name          :   frmMKTTRN0037
	'Created By         :   Sourabh Khatri
	'Created on         :   27/08/2004
	'Description        :   SRV Printing.
	'Modified Date      :   25/02/2005
	'Revisied by        :   Brij B Bohara
	'History            :   Added For NHK
	'---------------------------------------------------------------------------------------------------------------------
	'Revised By                 -  Davinder Singh
	'Revision Date              -  29/11/2005
	'Revision History           -  SRV printing was not in sequence when a range of Doc_Nos. are given for printing
    '                           -  Changes made in the Query in function cmdSRVprinting by adding order by class there according to issue ID:-16328
    'MODIFIED BY AJAY SHUKLA ON 10/MAY/2011 FOR MULTIUNIT CHANGE
	'------------------------------------------------------------------------------------------------------------------
	
	Dim mlngFormTag As Short ' Variable For Header String
    'Dim mObjSRVPrinting As New prj_InvoicePrinting.clsInvoicePrinting 'Object for SRV Printing.
    Dim mObjSRVPrinting As New prj_InvoicePrinting.clsInvoicePrinting(gstrDateFormat) 'Object for SRV Printing.
	Dim mstrMarutiCode As String 'Variable to Save Maruti Code
	Dim mdblWaitingTime As Short 'Variable to store Waiting time for printing.
	Dim mblnPrintSuffix As Boolean 'Variable to Store Print Suffix in SRV or Not
	Dim mblnDisplaySRVLotInfo As Boolean 'Added for SRV lot
	
	Private Sub chkBarCodePrinting_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles chkBarCodePrinting.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'----------------------------------------------------
		'Author              - Davinder Singh
		'Create Date         - 28/11/2005
		'Arguments           - Ascii code of key pressed
		'Return Value        - None
		'Function            - To set the focus to command buttons
		'----------------------------------------------------
		On Error GoTo Errorhandler
		If KeyAscii = System.Windows.Forms.Keys.Return Then cmdSRVPrinting.Focus()
		GoTo EventExitSub
Errorhandler: 
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        'Function for Close Report View.
        On Error GoTo Errorhandler
        FraInvoicePreview.Visible = False
        Me.txtInvoiceFrom.Focus()
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdInvoiceFrom_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdInvoiceFrom.Click
        'Function for Showing Help for Starting No of SRV
        Dim rsInvoice As New ClsResultSetDB
        Dim strInvoiceFrom As String
        On Error GoTo Errorhandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.WaitCursor)
        strInvoiceFrom = ShowList(0, (Me.txtInvoiceFrom.MaxLength), , "Doc_No", "Location_Code", "SalesChallan_dtl", "and Account_Code = '" & Trim(mstrMarutiCode) & "' and Invoice_Type <> 'SMP'")
        If strInvoiceFrom = "-1" Then
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
            Call ConfirmWindow(10435, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
        Else
            Me.txtInvoiceFrom.Text = CStr(strInvoiceFrom)
            Me.txtInvoiceTo.Focus()
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
	Private Sub cmdInvoiceTo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdInvoiceTo.Click
		'Function for Showing Help for End No. of SRV
		Dim rsInvoice As New ClsResultSetDB
		Dim strInvoiceTo As String
		On Error GoTo Errorhandler
		Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.WaitCursor)
		
		If Len(Me.txtInvoiceFrom.Text) < 1 Then
			MsgBox(" Please Select Starting No .of SRV", MsgBoxStyle.Information, ResolveResString(100))
			Me.txtInvoiceFrom.Focus() : Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
			Exit Sub
		End If

        strInvoiceTo = ShowList(0, (Me.txtInvoiceFrom.MaxLength), , "Doc_No", "Location_Code", "SalesChallan_dtl", "and Account_Code = '" & Trim(mstrMarutiCode) & "' and Doc_No >= '" & Me.txtInvoiceFrom.Text & "' and Invoice_Type <> 'SMP'")
        If strInvoiceTo = "-1" Then
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
            Call ConfirmWindow(10435, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
        Else
            Me.txtInvoiceTo.Text = CStr(strInvoiceTo)
            Me.chkBarCodePrinting.Focus()
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrint.Click
        On Error GoTo Errorhandler
        Call cmdSRVPrinting_ButtonClick(cmdSRVPrinting, New UCActXCtl.UCfraRepCmd.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_PRINTER))
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdSRVPrinting_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCfraRepCmd.ButtonClickEventArgs) Handles cmdSRVPrinting.ButtonClick
        'Function for Handling Different Button Opration
        Dim rsSRV As New ClsResultSetDB
        Dim strSQL As String
        Dim strMarutiCode As String
        Dim strSRVNo As String
        Dim dblWaitingTime As Double
        Dim varPrint As Object
        Dim strFileName As String
        On Error GoTo Errorhandler
        Select Case e.Button

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_FILE
                MsgBox(" Export Can Not Be Valid Option For SRV ", MsgBoxStyle.Information, ResolveResString(100)) : Exit Sub

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_WINDOW
                'Validate Data
                If ValidateData() Then
                    'Validate only for one invoice No
                    If (Val(Me.txtInvoiceFrom.Text) - Val(Me.txtInvoiceTo.Text) <> 0) Then MsgBox(" Please Select Maximum one Invoice No. ", MsgBoxStyle.Information, ResolveResString(100)) : Me.txtInvoiceFrom.Focus() : Exit Sub


                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                    strSRVNo = Trim(Me.txtInvoiceFrom.Text)
                    'Generate Text File
                    Call SRVGenerating(CInt(strSRVNo))
                    'Show Text file in Rich Text Box
                    rtbInvoicePreview.ScrollBars = RichTextBoxScrollBars.Both
                    rtbInvoicePreview.LoadFile(gstrLocalCDrive & "tmp\S" & Mid(strSRVNo, 3, Len(strSRVNo)) & ".txt", RichTextBoxStreamType.PlainText)
                    rtbInvoicePreview.BackColor = System.Drawing.Color.White
                    cmdPrint.Image = My.Resources.ico231.ToBitmap
                    cmdClose.Image = My.Resources.ico217.ToBitmap
                    FraInvoicePreview.Visible = True
                    FraInvoicePreview.Enabled = True
                    FraInvoicePreview.BringToFront()
                    FraInvoicePreview.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 1050)
                    FraInvoicePreview.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) - 400)
                    FraInvoicePreview.Left = VB6.TwipsToPixelsX(100)
                    FraInvoicePreview.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.ctlHeader.Height) - 50)
                    rtbInvoicePreview.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(FraInvoicePreview.Height) - 1000)
                    rtbInvoicePreview.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(FraInvoicePreview.Width) - 200)
                    rtbInvoicePreview.Left = VB6.TwipsToPixelsX(100)
                    rtbInvoicePreview.Top = VB6.TwipsToPixelsY(900)
                    rtbInvoicePreview.RightMargin = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(rtbInvoicePreview.Width) + 5000)
                    shpInvoice.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(FraInvoicePreview.Width) - VB6.PixelsToTwipsX(shpInvoice.Width)) / 2)
                    cmdPrint.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(shpInvoice.Left) + 100)
                    cmdClose.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(cmdPrint.Left) + VB6.PixelsToTwipsX(cmdPrint.Width) + 100)
                    ReplaceJunkCharacters()
                    cmdPrint.Enabled = True : cmdClose.Enabled = True
                    FraInvoicePreview.Enabled = True : rtbInvoicePreview.Enabled = True : rtbInvoicePreview.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    rtbInvoicePreview.Focus()
                End If
                'End If


                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_PRINTER
                If ValidateData() Then
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.WaitCursor)
                    'Assign Maruti Code and Waiting time in Local variables
                    strMarutiCode = mstrMarutiCode
                    dblWaitingTime = mdblWaitingTime
                    If dblWaitingTime = 0 Then dblWaitingTime = 5000
                    strSQL = "Select Doc_No from SalesChallan_dtl where doc_no between " & Val(Me.txtInvoiceFrom.Text) & " and " & Val(Me.txtInvoiceTo.Text) & " and account_Code = '" & strMarutiCode & "' and cancel_flag = 0 and Invoice_Type <> 'SMP' and UNIT_CODE='" & gstrUNITID & "' order by Doc_No"
                    rsSRV.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    On Error Resume Next
                    Kill((gstrLocalCDrive & "EmproInv\prnToPrint.bat"))
                    On Error GoTo Errorhandler
                    FileOpen(1, gstrLocalCDrive & "EmproInv\prnToprint.bat", OpenMode.Output)
                    PrintLine(1, "Type %1 >prn ")
                    FileClose(1)

                    While rsSRV.EOFRecord <> True
                        Me.lblInfo.Text = ""

                        strSRVNo = rsSRV.GetValue("Doc_No")
                        Call SRVGenerating(CInt(strSRVNo))
                        Me.lblInfo.Text = " Printing SRV No. :- " & strSRVNo
Gotoprinting:
                        'Changed for NHK Bar Code is at the beginning of the page
                        If StrComp(UCase(Trim(gstrUNITID)), "NHK", CompareMethod.Text) = 0 Then
                            'Code for Bar Code Printing
                            If Me.chkBarCodePrinting.CheckState = 1 Then
                                strFileName = gstrLocalCDrive & "tmp\B" & Mid(strSRVNo, 3, Len(strSRVNo)) & ".txt"
                                Call BarCodePrint(strFileName)
                                Sleep((dblWaitingTime))
                            Else
                                strFileName = gstrLocalCDrive & "EmproInv\prnToPrint.bat " & gstrLocalCDrive & "EmproInv\PagefeedwWithoutBarCode.txt"

                                varPrint = Shell("cmd.exe /c " & strFileName, AppWinStyle.Hide)
                                Sleep((dblWaitingTime))
                            End If
                            'Code for SRV Printing
                            strFileName = gstrLocalCDrive & "EmproInv\prnToPrint.bat " & gstrLocalCDrive & "tmp\S" & Mid(strSRVNo, 3, Len(strSRVNo)) & ".txt"

                            varPrint = Shell("cmd.exe /c " & strFileName, AppWinStyle.Hide)
                            Sleep((dblWaitingTime))

                        Else 'Code for SRV Printing
                            strFileName = gstrLocalCDrive & "EmproInv\prnToPrint.bat " & gstrLocalCDrive & "tmp\S" & Mid(strSRVNo, 3, Len(strSRVNo)) & ".txt"

                            varPrint = Shell("cmd.exe /c " & strFileName, AppWinStyle.Hide)
                            'Code for Bar Code Printing
                            Sleep((dblWaitingTime))
                            If Me.chkBarCodePrinting.CheckState = 1 Then
                                strFileName = gstrLocalCDrive & "tmp\B" & Mid(strSRVNo, 3, Len(strSRVNo)) & ".txt"
                                Call BarCodePrint(strFileName)
                                Sleep((dblWaitingTime))
                            Else
                                strFileName = gstrLocalCDrive & "EmproInv\prnToPrint.bat " & gstrLocalCDrive & "EmproInv\PagefeedwWithoutBarCode.txt"

                                varPrint = Shell("cmd.exe /c " & strFileName, AppWinStyle.Hide)
                                Sleep((dblWaitingTime))
                            End If

                        End If
                        'Code for Adjust Page .
                        strFileName = gstrLocalCDrive & "EmproInv\prnToPrint.bat " & gstrLocalCDrive & "EmproInv\Pagefeed.txt"

                        varPrint = Shell("cmd.exe /c " & strFileName, AppWinStyle.Hide)
                        Sleep((dblWaitingTime))
                        'Changes ends here
                        rsSRV.MoveNext()
                    End While

                    rsSRV = Nothing
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
                    Me.lblInfo.Text = ""
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
        End Select
        Exit Sub
Errorhandler:
        If Err.Number = 53 Then
            FileOpen(1, gstrLocalCDrive & "EmproInv\prnToprint.bat", OpenMode.Output)
            PrintLine(1, "Type %1 >prn ")
            FileClose(1) : GoTo Gotoprinting
        Else
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
            Me.lblInfo.Text = ""
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub frmMKTTRN0037_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Dim rsSRV As New ClsResultSetDB 'Added for SRV Lot Information
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mlngFormTag
        frmModules.NodeFontBold(Tag) = True
        Me.txtInvoiceFrom.Focus()
        Call rsSRV.GetResult("Select DisplaySRVLotInformation=isnull(DisplaySRVLotInformation,0)   from sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSRV.RowCount > 0 Then
            If rsSRV.GetValue("DisplaySRVLotInformation") = True Then
                fraLotInfo.Visible = True
                mblnDisplaySRVLotInfo = True
            End If
        Else
            mblnDisplaySRVLotInfo = False
        End If
        rsSRV = Nothing
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0037_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Click
        On Error GoTo ErrHandler
        'Checking the form name in the Windows list
        mdifrmMain.CheckFormName = mlngFormTag
        frmModules.NodeFontBold(Me.Tag) = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0037_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        'Make the node normal font
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0037_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlHeader_Click(ctlHeader, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0037_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '----------------------------------------------------
        'Author              - Davinder Singh
        'Create Date         - 28/11/2005
        'Arguments           - Keycode of key pressed and shift key
        'Return Value        - None
        'Function            - To Unload the form
        '----------------------------------------------------
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.Escape And Shift = 0 Then
            Me.Close()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0037_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo Errorhandler
        Dim rsSRV As New ClsResultSetDB
        mlngFormTag = mdifrmMain.AddFormNameToWindowList(Me.ctlHeader.Tag)
        Call FitToClient(Me, frmMainFrame, ctlHeader, (Me.cmdSRVPrinting))
        Me.txtInvoiceFrom.Text = "" : Me.txtInvoiceTo.Text = "" : Me.chkBarCodePrinting.CheckState = System.Windows.Forms.CheckState.Checked
        Me.frmMainFrame.Visible = True
        Me.FraInvoicePreview.Visible = False
        'Code for saving Maruti Code in Variable for Use
        Call rsSRV.GetResult("Select Maruti_ac=isnull(Maruti_ac,''),WaitingTime=isnull(WaitingTime,0),PrintSuffix=isnull(PrintSuffix,0),  DisplaySRVLotInformation=isnull(DisplaySRVLotInformation,0)   from sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSRV.RowCount > 0 Then
            mstrMarutiCode = rsSRV.GetValue("Maruti_ac")
            mdblWaitingTime = Val(rsSRV.GetValue("WaitingTime"))
            mblnPrintSuffix = CBool(rsSRV.GetValue("PrintSuffix"))
            If rsSRV.GetValue("DisplaySRVLotInformation") = True Then
                fraLotInfo.Visible = True
            End If

        End If

        rsSRV = Nothing
        cmdPrint.Image = My.Resources.ico231.ToBitmap
        cmdPrint.Image = My.Resources.ico230.ToBitmap
        rtbInvoicePreview.ScrollBars = RichTextBoxScrollBars.ForcedHorizontal
        rtbInvoicePreview.ScrollBars = RichTextBoxScrollBars.ForcedVertical

        If Directory.Exists(gstrLocalCDrive + "EmproInv") = False Then
            Directory.CreateDirectory(gstrLocalCDrive + "EmproInv")
        End If
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0037_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        'Removing the form name from list
        mdifrmMain.RemoveFormNameFromWindowList = mlngFormTag
        'Setting the corresponding node's tag
        frmModules.NodeFontBold(Tag) = False
        Me.Dispose()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub opnRegular_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles opnRegular.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo Errorhandler
            If fraLotInfo.Visible = True Then
                If opnRegular.Checked = True Then
                    mblnDisplaySRVLotInfo = True
                End If
            Else
                mblnDisplaySRVLotInfo = False
            End If
            Exit Sub
Errorhandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub opnSmall_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles opnSmall.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo Errorhandler
            If opnSmall.Checked = True Then
                mblnDisplaySRVLotInfo = False
            End If
            Exit Sub
Errorhandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

        End If
    End Sub
    Private Sub txtInvoiceFrom_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtInvoiceFrom.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Errorhandler
        If Shift = 0 And KeyCode = System.Windows.Forms.Keys.F1 Then
            Call cmdInvoiceFrom_Click(cmdInvoiceFrom, New System.EventArgs())
        End If
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtInvoiceFrom_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtInvoiceFrom.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo Errorhandler
        If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13 Then
            KeyAscii = KeyAscii
        Else
            KeyAscii = 0
        End If
        If KeyAscii = 13 Then
            Me.txtInvoiceTo.Focus()
        End If
        GoTo EventExitSub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtInvoiceFrom_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtInvoiceFrom.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim rsInvoice As New ClsResultSetDB
        Dim strSQL As String
        On Error GoTo Errorhandler
        If Len(Trim(Me.txtInvoiceFrom.Text)) > 0 Then
            'Validation Maruti Code exist or not in Sales Parameter
            If Len(mstrMarutiCode) < 1 Then
                MsgBox(" Please Define Maruti Code in Sales Parameter", MsgBoxStyle.Information, ResolveResString(100))
                Me.txtInvoiceFrom.Text = "" : Me.txtInvoiceFrom.Focus() : GoTo EventExitSub
            End If
            'Validation for Invoice No
            strSQL = "Select Doc_no from saleschallan_dtl where doc_no = '" & Trim(Me.txtInvoiceFrom.Text) & "' and account_Code = '" & mstrMarutiCode & "' and cancel_flag = 0 and Invoice_Type <> 'SMP' AND UNIT_CODE='" & gstrUNITID & "'"
            rsInvoice.GetResult(strSQL)
            If rsInvoice.RowCount < 1 Then
                MsgBox("Invoice No. Enter by you is Invalid . Please try again", MsgBoxStyle.Information, ResolveResString(100))
                Me.txtInvoiceFrom.Text = "" : Me.txtInvoiceFrom.Focus()
            End If

            rsInvoice = Nothing
        End If
        GoTo EventExitSub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtInvoiceTo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtInvoiceTo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Errorhandler
        If Shift = 0 And KeyCode = System.Windows.Forms.Keys.F1 Then
            Call cmdInvoiceTo_Click(cmdInvoiceTo, New System.EventArgs())
        End If
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub
    Private Sub txtInvoiceTo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtInvoiceTo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo Errorhandler
        If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13 Then
            KeyAscii = KeyAscii
        Else
            KeyAscii = 0
        End If
        If KeyAscii = 13 Then
            Me.chkBarCodePrinting.Focus()
        End If
        GoTo EventExitSub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtInvoiceTo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtInvoiceTo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim rsInvoice As New ClsResultSetDB
        Dim strSQL As String
        Dim strAccountCode As String
        On Error GoTo Errorhandler

        If Len(Trim(Me.txtInvoiceTo.Text)) > 0 Then
            'Validation Maruti Code exist or not in Sales Parameter
            If Len(mstrMarutiCode) < 1 Then
                MsgBox(" Please define Maruti Code in Sales Parameter", MsgBoxStyle.Information, ResolveResString(100))
                Me.txtInvoiceTo.Text = "" : Me.txtInvoiceFrom.Text = "" : Me.txtInvoiceFrom.Focus() : GoTo EventExitSub
            End If
            'Validation for Invoice No
            strSQL = "Select Doc_no from saleschallan_dtl where doc_no = '" & Trim(Me.txtInvoiceTo.Text) & "' and account_Code = '" & mstrMarutiCode & "' and cancel_flag = 0 and  Doc_No >= '" & Me.txtInvoiceFrom.Text & "' and Invoice_Type <> 'SMP' AND UNIT_CODE='" & gstrUNITID & "'"
            rsInvoice.GetResult(strSQL)
            If rsInvoice.RowCount < 1 Then
                MsgBox("Invoice No. Entered by you is Invalid . Please try again", MsgBoxStyle.Information, ResolveResString(100))
                Me.txtInvoiceTo.Text = "" : Me.txtInvoiceTo.Focus() : GoTo EventExitSub
            End If

            rsInvoice = Nothing
            Me.chkBarCodePrinting.Focus()
        End If
        GoTo EventExitSub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

EventExitSub:
        eventArgs.Cancel = Cancel
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
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "x0", "") 'Remove Draft Character
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function SRVGenerating(ByVal SRVNo As Integer) As Boolean
        Dim strSRVFileName As String
        Dim strBSRVFilename As String
        Dim strSQL As String
        Dim rsSRV As New ClsResultSetDB
        Dim fsObjFile As New Scripting.FileSystemObject
        Dim blnBarCodePrinting As Boolean

        On Error GoTo Errorhandler
        SRVGenerating = True
        strSQL = "Select Location_Code,Account_Code from SalesChallan_dtl where Doc_no = '" & SRVNo & "' AND UNIT_CODE='" & gstrUNITID & "'"
        Call rsSRV.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSRV.RowCount < 1 Then MsgBox(" Location Code or Account Code Not Found ", MsgBoxStyle.Information, ResolveResString(100))
        On Error Resume Next
        fsObjFile.CreateFolder(gstrLocalCDrive & "tmp")
        Kill((gstrLocalCDrive & "tmp\S*.txt"))
        Kill((gstrLocalCDrive & "tmp\B*.txt"))
        On Error GoTo Errorhandler

        mObjSRVPrinting.ConnectionString = gstrCONNECTIONSTRING
        mObjSRVPrinting.Connection()
        mObjSRVPrinting.CompanyName = gstrCOMPANY
        mObjSRVPrinting.Address1 = gstr_WRK_ADDRESS1
        mObjSRVPrinting.Address2 = gstr_WRK_ADDRESS2
        strSRVFileName = gstrLocalCDrive & "tmp\S" & Mid(CStr(SRVNo), 3, Len(CStr(SRVNo))) & ".txt"
        mObjSRVPrinting.FileName = strSRVFileName
        strBSRVFilename = gstrLocalCDrive & "tmp\B" & Mid(CStr(SRVNo), 3, Len(CStr(SRVNo))) & ".txt"
        mObjSRVPrinting.BCFileName = strBSRVFilename
        blnBarCodePrinting = Me.chkBarCodePrinting.CheckState
        Call mObjSRVPrinting.PrintSrv(gstrUNITID, True, blnBarCodePrinting, mblnPrintSuffix, CStr(Val(CStr(SRVNo))), rsSRV.GetValue("Account_Code"), rsSRV.GetValue("Location_Code"))

        Exit Function
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        SRVGenerating = False

    End Function
    Public Function ValidateData() As Boolean
        Dim intSwap As Short
        On Error GoTo Errorhandler
        If Len(Me.txtInvoiceFrom.Text) < 1 Then
            MsgBox("Invoice No. can not be blank.Please Enter Invoice No ", MsgBoxStyle.Information, ResolveResString(100))
            Me.txtInvoiceFrom.Focus() : ValidateData = False : Exit Function
        ElseIf Len(Me.txtInvoiceTo.Text) < 1 Then
            MsgBox("Invoice No. can not be blank.Please Enter Invoice No ", MsgBoxStyle.Information, ResolveResString(100))
            Me.txtInvoiceTo.Focus() : ValidateData = False : Exit Function
        End If

        If Val(Me.txtInvoiceFrom.Text) > Val(Me.txtInvoiceTo.Text) Then
            intSwap = Val(Me.txtInvoiceTo.Text)
            Me.txtInvoiceFrom.Text = Me.txtInvoiceTo.Text
            Me.txtInvoiceTo.Text = CStr(intSwap)
        End If
        ValidateData = True
        Exit Function
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub BarCodePrint(ByRef strFileName As String)
        On Error GoTo Errorhandler
        Dim varPrint As Object
        Dim strString As String
        strString = gstrLocalCDrive & "EmproInv\pdf-dot.bat " & strFileName & " 4 2 2 1"

        varPrint = Shell("cmd.exe /c " & strString, AppWinStyle.Hide)
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ctlHeader_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlHeader.Click
        On Error GoTo Errorhandler
        MsgBox(" Help is Not Available for SRV Printing.", MsgBoxStyle.Information, ResolveResString(100))
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
End Class