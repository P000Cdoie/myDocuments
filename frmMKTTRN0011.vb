Option Strict Off
Option Explicit On
Imports System
Imports System.Data
Imports System.Data.SqlClient


Friend Class frmMKTTRN0011
	Inherits System.Windows.Forms.Form
	'---------------------------------------------------------------------------
	'Copyright          :   MIND Ltd.
	'Form Name          :   frmMKTTRN0011.frm
	'Created By         :   Ananya Nath
	'Created on         :   04/02/2002
	'Description        :   Copy of Sales Order
	'Revision Date      :   02.07.2002 Ananya
	'Revision History   :   SALES ORDER CAN BE COPIED FOR THE SAME ACCOUNT CODE
	'Revision Date      :   28.08.2002 Ananya
	'Revision History   :   An Authorised Base So willl be copied if user selects an Amendment.
	'                       Amendment will be copied as unauthorized one.
	'Revision History   :   14-10-2002, Changes done as per Acct. link and new custitem_mst table structure.
	'Revision History   :   21-11-2002, New Fields like Per Value, Internal No., Revision No. is added, necessary changes is made
	'------------------------------------------------------------------------------------------------
	'Revised By         : Arul mozhi Varman
	'Revied On          : 12-10-2004
	'Revision History   : To Give the provision to make the Copy of sales order for First Amentment
	'-----------------------------------------------------------------------------------------
	'Revised By         : Arul
	'Revised On         : 16-11-2004
	'Revision History   : To Insert the Tool Amortization Required Flag value while Making the Copy of the Sales Order
	'------------------------------------------------------------------------------------------------------
	'Revised By         : Arul
	'Revised On         : 20-01-2005
	'Revision History   : To Insert the Ecess Type & Surcharge while make the copy of sales order
	'------------------------------------------------------------------------------------------------------
	'Revised By         : Arul
	'Revised On         : 11-01-2006
	'Revision History   : To Insert the Show authorization flag in custorder detail table
	'------------------------------------------------------------------------------------------------------
	'Revised By         : Arul
	'Revised On         : 07-04-2006
	'Revision History   : To update the show authorization flag by "1"
	'------------------------------------------------------------------------------------------------------
	'Revised By         : Manoj Kr. Vaish
	'Revised On         : 17-Nov-2007 Issue Id(21478)
	'Revision History   : To show the consignee controls with zone check and database filed addition
    '------------------------------------------------------------------------------------------------------
    'Revised By         : Manoj Kr. Vaish
    'Revised On         : 27 Feb 2009
    'Issue ID           : eMpro-20090223-27780
    'Revision History   : To get the zone name from sales_parameter
    'Modified by Amit Rana on 2011-Apr-25
    '   Modified to support MultiUnit functionality
    '------------------------------------------------------------------------------------------------------
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10229989
    'Revision Date   : 10-aug -2013- 31-aug -2013
    'History         : Multiple So Functionlity 
    '****************************************************************************************
	''Declare Variables
	Dim mintFormIndex As Short
	Dim mRdoCls As New ClsResultSetDB
	Dim mRdocls1 As New ClsResultSetDB
	Dim mbolAmendDate As Boolean
	Dim mstrAmend As String
	Dim mbolCopySo As Boolean
	Dim mbolBaseSo As Boolean
	Dim mstrIntNoBase As String
	Dim mstrIntNoAmend As String
	Dim mintRevBase As Short
    Dim mintRevAmend As Short

    Private Sub cmdhelpAmend_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdhelpAmend.Click
        Dim Index As Short = cmdhelpAmend.GetIndex(eventSender)
        On Error GoTo ErrHandler
        Dim strCheckRec As String
        Select Case Index
            Case 1
                With Me.txtamend
                    If Len(.Text) = 0 Then
                        strCheckRec = ShowList(1, .MaxLength, "", "cust_ord_hdr.amendment_no", "cust_ord_hdr.cust_ref", "cust_ord_hdr", " and cust_ord_hdr.account_code = '" & txtCustFrom.Text & "'  and cust_ord_hdr.cust_ref = '" & txtSalesRefFrom.Text & "' and cust_ord_hdr.amendment_no <> '' ")
                        If strCheckRec = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                            .Focus()
                            Exit Sub
                        End If
                    Else
                        strCheckRec = ShowList(1, .MaxLength, "" & txtamend.Text & "", "cust_ord_hdr.amendment_no", "cust_ord_hdr.cust_ref", "cust_ord_hdr", " and cust_ord_hdr.account_code = '" & txtCustFrom.Text & "'  and cust_ord_hdr.cust_ref = '" & txtSalesRefFrom.Text & "' and cust_ord_hdr.amendment_no <> '' ")
                        If strCheckRec = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                            .Focus()
                            Exit Sub
                        End If
                    End If
                    .Text = strCheckRec
                    Me.txtCustTo.Focus()
                End With
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdhelpCons_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdhelpCons.Click
        Dim Index As Short = cmdhelpCons.GetIndex(eventSender)
        On Error GoTo ErrHandler
        Dim strCheckRec As String
        Select Case Index
            Case 0 'TxtconsFrom
                With Me.txtconsfrom
                    If Len(.Text) = 0 Then
                        strCheckRec = ShowList(1, .MaxLength, "", "cust_ord_hdr.consignee_code", "customer_mst.cust_name", "cust_ord_hdr, customer_mst", " and customer_mst.customer_code = cust_ord_hdr.consignee_code and customer_mst.Unit_Code = cust_ord_hdr.Unit_Code and cust_ord_hdr.cust_ref <> '' and cust_ord_hdr.account_code ='" & txtCustFrom.Text & "' and ((isnull(customer_mst.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= customer_mst.deactive_date))", , , , , , "cust_ord_hdr.Unit_Code")
                        If strCheckRec = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                            txtconsfrom.Focus()
                            Exit Sub
                        End If
                    Else
                        strCheckRec = ShowList(1, .MaxLength, "" & txtconsfrom.Text & "", "cust_ord_hdr.consignee_code", "customer_mst.cust_name", "cust_ord_hdr, customer_mst", " and customer_mst.customer_code = cust_ord_hdr.consignee_code and customer_mst.Unit_code = cust_ord_hdr.Unit_code and cust_ord_hdr.cust_ref <> '' and cust_ord_hdr.account_code ='" & txtCustFrom.Text & "' and ((isnull(customer_mst.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= customer_mst.deactive_date))", , , , , , "cust_ord_hdr.Unit_code")
                        If strCheckRec = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                            txtconsfrom.Focus()
                            Exit Sub
                        End If
                    End If
                    .Text = strCheckRec
                    Call ProcName((txtconsfrom.Text), 3)
                    If txtconsfrom.Text <> "" Then Me.cmdGrpCopySO.Enabled(1) = True
                    Me.txtSalesRefFrom.Focus()
                End With
            Case 1 'txtconsto
                With Me.txtconsto
                    If Len(.Text) = 0 Then
                        strCheckRec = ShowList(1, .MaxLength, "", "cust_ord_hdr.consignee_code", "customer_mst.cust_name", "customer_mst,cust_ord_hdr ", "and cust_ord_hdr.account_code ='" & txtCustTo.Text & "' and customer_mst.customer_code =cust_ord_hdr.consignee_code and customer_mst.Unit_Code =cust_ord_hdr.Unit_Code and ((isnull(customer_mst.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= customer_mst.deactive_date))", , , , , , "cust_ord_hdr.Unit_Code")
                        If strCheckRec = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                            txtconsto.Focus()
                            Exit Sub
                        End If
                    Else
                        strCheckRec = ShowList(1, .MaxLength, "" & txtconsto.Text & "", "cust_ord_hdr.consignee_code", "customer_mst.cust_name", "customer_mst,cust_ord_hdr", "and customer_mst.customer_code =cust_ord_hdr.consignee_code and customer_mst.unit_code = cust_ord_hdr.unit_code and  cust_ord_hdr.account_code ='" & txtCustTo.Text & "' and ((isnull(customer_mst.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= customer_mst.deactive_date))", , , , , , "cust_ord_hdr.unit_code")
                        If strCheckRec = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                            txtconsto.Focus()
                            Exit Sub
                        End If
                    End If
                    .Text = strCheckRec
                    Call ProcName((txtconsto.Text), 4)
                    Me.txtSalesRefTo.Focus()
                End With
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen) 'Change the Mouse Pointer of the Screen
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmdhelpCust_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdhelpCust.Click
        Dim Index As Short = cmdhelpCust.GetIndex(eventSender)
        On Error GoTo ErrHandler
        Dim strCheckRec As String
        Select Case Index
            Case 2 'TxtCustFrom
                With Me.txtCustFrom
                    If Len(.Text) = 0 Then
                        strCheckRec = ShowList(1, .MaxLength, "", "cust_ord_hdr.account_code", "customer_mst.cust_name", "cust_ord_hdr, customer_mst", " and customer_mst.customer_code = cust_ord_hdr.account_code and customer_mst.Unit_Code = cust_ord_hdr.Unit_Code and cust_ord_hdr.cust_ref <> '' and ((isnull(customer_mst.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= customer_mst.deactive_date)) ", , , , , , "cust_ord_hdr.Unit_code")
                        If strCheckRec = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                            txtCustFrom.Focus()
                            Exit Sub
                        End If
                    Else
                        strCheckRec = ShowList(1, .MaxLength, "" & txtCustFrom.Text & "", "cust_ord_hdr.account_code", "customer_mst.cust_name", "cust_ord_hdr, customer_mst", " and customer_mst.customer_code = cust_ord_hdr.account_code and customer_mst.Unit_Code = cust_ord_hdr.Unit_Code and cust_ord_hdr.cust_ref <> '' and ((isnull(customer_mst.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= customer_mst.deactive_date))", , , , , , "cust_ord_hdr.Unit_Code")
                        If strCheckRec = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                            txtCustFrom.Focus()
                            Exit Sub
                        End If
                    End If
                    .Text = strCheckRec
                    Call ProcName((txtCustFrom.Text), 1)
                    If txtCustFrom.Text <> "" Then Me.cmdGrpCopySO.Enabled(1) = True
                    Me.txtSalesRefFrom.Focus()
                End With
            Case 3 'txtcustto
                With Me.txtCustTo
                    If Len(.Text) = 0 Then
                        strCheckRec = ShowList(1, .MaxLength, "", "customer_mst.customer_code", "customer_mst.cust_name", "customer_mst ", "and ((isnull(customer_mst.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= customer_mst.deactive_date))")
                        If strCheckRec = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                            txtCustTo.Focus()
                            Exit Sub
                        End If
                    Else
                        strCheckRec = ShowList(1, .MaxLength, "" & txtCustTo.Text & "", "customer_mst.customer_code", "customer_mst.cust_name", "customer_mst", " and ((isnull(customer_mst.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= customer_mst.deactive_date))")
                        If strCheckRec = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                            txtCustTo.Focus()
                            Exit Sub
                        End If
                    End If
                    .Text = strCheckRec
                    Call ProcName((txtCustTo.Text), 2)
                    Me.txtSalesRefTo.Focus()
                End With
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen) 'Change the Mouse Pointer of the Screen
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmdhelpSales_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdhelpSales.Click
        Dim Index As Short = cmdhelpSales.GetIndex(eventSender)
        On Error GoTo ErrHandler
        Dim strSQL As String
        Dim strCheckRec As String
        Select Case Index
            Case 0 'salesRefFrom
                With Me.txtSalesRefFrom
                    If Len(.Text) = 0 Then
                        If GetMktCodeExecutionZone() = "SOUTH" Then
                            strCheckRec = ShowList(1, .MaxLength, "", "cust_ord_hdr.cust_ref", "cust_ord_hdr.account_code", "cust_ord_hdr", " and cust_ord_hdr.account_code = '" & txtCustFrom.Text & "' and cust_ord_hdr.consignee_code='" & txtconsfrom.Text & "' ")
                        Else
                            strCheckRec = ShowList(1, .MaxLength, "", "cust_ord_hdr.cust_ref", "cust_ord_hdr.account_code", "cust_ord_hdr", " and cust_ord_hdr.account_code = '" & txtCustFrom.Text & "'")
                        End If
                        If strCheckRec = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                            txtSalesRefFrom.Focus()
                            Exit Sub
                        End If
                    Else
                        If GetMktCodeExecutionZone() = "SOUTH" Then
                            strCheckRec = ShowList(1, .MaxLength, "" & txtSalesRefFrom.Text & "", "cust_ord_hdr.cust_ref", "cust_ord_hdr.account_code", "cust_ord_hdr", " and cust_ord_hdr.account_code = '" & txtCustFrom.Text & "' and cust_ord_hdr.consignee_code ='" & txtconsfrom.Text & "'")
                        Else
                            strCheckRec = ShowList(1, .MaxLength, "" & txtSalesRefFrom.Text & "", "cust_ord_hdr.cust_ref", "cust_ord_hdr.account_code", "cust_ord_hdr", " and cust_ord_hdr.account_code = '" & txtCustFrom.Text & "'")
                        End If
                        If strCheckRec = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                            txtSalesRefFrom.Focus()
                            Exit Sub
                        End If
                    End If
                    .Text = strCheckRec
                End With
                If Me.txtSalesRefFrom.Text <> "" Then
                    strSQL = "select amendment_no from cust_ord_hdr Where Unit_Code='" & gstrUNITID & "' And account_code = '" & txtCustFrom.Text & "' and cust_ref = '" & txtSalesRefFrom.Text & "' and amendment_no <> '' "
                    If mRdoCls.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset) And mRdoCls.GetNoRows > 0 Then
                        Me.txtamend.Enabled = True : Me.cmdhelpAmend(1).Enabled = True : Me.txtamend.Focus() : txtamend.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        Me.txtAmendTo.Enabled = True : txtAmendTo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    Else
                        Me.txtCustTo.Focus()
                    End If
                Else
                    Me.txtSalesRefFrom.Focus()
                End If
            Case 1 'salesRefTo
                With Me.txtSalesRefTo
                    If Len(.Text) = 0 Then
                        strCheckRec = ShowList(1, .MaxLength, "", "cust_ord_hdr.cust_ref", "cust_ord_hdr.account_code", "cust_ord_hdr", " and cust_ord_hdr.account_code = '" & txtCustTo.Text & "' ")
                        If strCheckRec = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                            txtSalesRefTo.Focus()
                            Exit Sub
                        End If
                    Else
                        strCheckRec = ShowList(1, .MaxLength, "" & txtSalesRefTo.Text & "", "cust_ord_hdr.cust_ref", "cust_ord_hdr.account_code", "cust_ord_hdr", " and cust_ord_hdr.account_code = '" & txtCustTo.Text & "' ")
                        If strCheckRec = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                            txtSalesRefTo.Focus()
                            Exit Sub
                        End If
                    End If
                    .Text = strCheckRec
                    Me.cmdGrpCopySO.Enabled(0) = True
                    If Me.txtAmendTo.Enabled = False Then Me.cmdGrpCopySO.Focus() Else Me.txtAmendTo.Focus()
                End With
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0011_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        'Checking the form name in the Windows list
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        'This is to avoid the execution of the error handler
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0011_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        Dim rsTemp As New ClsResultSetDB
        Dim strSQL As String
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Tag) = False
        gblnCancelUnload = False
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0011_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        On Error GoTo ErrHandler
        Dim strSQL As String
        strSQL = "Drop Table #tmpcust_ord_hdr "
        strSQL = Trim(strSQL) & vbCrLf & "Drop Table #tmpcust_ord_dtl "
        strSQL = Trim(strSQL) & vbCrLf & "Drop Table  #tmpcustitem_mst "
        mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Me.Dispose()
        'Removing the form name from list
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        frmModules.NodeFontBold(Tag) = False
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0011_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If (KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0) Then Call ctlFormHeader1_Click(ctlFormHeader1, New System.EventArgs()) 'F4 key is used to call empHelp
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0011_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                System.Windows.Forms.SendKeys.SendWait("{TAB}") 'If user press the Enter Key ,the focus will be advanced        Case vbKeyEscape  'If user press Escape than valCancel will be callked.
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
    Private Sub frmMKTTRN0011_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        Dim intCount As Short
        Call ProcCreateTab()
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FitToClient(Me, fraMain, ctlFormHeader1, cmdGrpCopySO)
        Me.cmdhelpAmend(1).Image = My.Resources.ico111.ToBitmap
        Me.cmdhelpCust(2).Image = My.Resources.ico111.ToBitmap
        Me.cmdhelpCust(3).Image = My.Resources.ico111.ToBitmap
        Me.cmdhelpSales(0).Image = My.Resources.ico111.ToBitmap
        Me.cmdhelpCons(0).Image = My.Resources.ico111.ToBitmap
        Me.cmdhelpCons(1).Image = My.Resources.ico111.ToBitmap
        With Me
            cmdGrpCopySO.Enabled(0) = True
            cmdGrpCopySO.Enabled(1) = False
            cmdGrpCopySO.Caption(1) = "Refresh"
            cmdGrpCopySO.Enabled(2) = False
            cmdGrpCopySO.Caption(0) = "Copy"
            .cmdhelpCust(2).Enabled = True
            .cmdhelpCust(3).Enabled = True
            .cmdhelpSales(0).Enabled = True
            .txtamend.Enabled = False : .txtamend.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            .cmdhelpAmend(1).Enabled = False
            .txtAmendTo.Enabled = False : .txtAmendTo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        End With
        If GetMktCodeExecutionZone() = "SOUTH" Then
            txtconsfrom.Enabled = True
            txtconsfrom.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            txtconsto.Enabled = True
            txtconsto.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            cmdhelpCons(0).Enabled = True
            cmdhelpCons(1).Enabled = True
        Else
            txtconsfrom.Enabled = False
            txtconsfrom.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtconsto.Enabled = False
            txtconsto.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            cmdhelpCons(0).Enabled = False
            cmdhelpCons(1).Enabled = False
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    '*********************************************'
    'Author:                Ananya Nath
    'Arguments:             None
    'Return Value   :       True/False
    'Description    :       Returns True, if entered/selected Currency code has taken part in transaction
    '*********************************************'
    Public Function valReferenceNo(ByRef strCustCode As String, ByRef strrefno As String) As Boolean ' Checks for existance of rec. in transaction
        On Error GoTo ErrHandler
        Dim strTrans As Boolean
        Dim strSQL As String
        Dim stracccode As String
        Dim strAmendNo As String
        strSQL = "select cust_ref,account_code from cust_ord_hdr Where Unit_Code='" & gstrUNITID & "' And account_code = '" & strCustCode & "' and cust_ref = '" & strrefno & "' "
        If mRdoCls.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset) And mRdoCls.GetNoRows > 0 Then
            valReferenceNo = True
        Else
            valReferenceNo = False
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    '*********************************************'
    'Author:                Ananya Nath
    'Arguments:             None
    'Return Value   :       True/False
    'Description    :       Returns True, if Customer Code is existing, else False.
    '*********************************************
    Public Function ValCustomerCode(ByRef strCustCode As String, ByRef strCustFrom As String) As Boolean ' Checks for Customer Code
        On Error GoTo ErrHandler
        Dim strSQL As String
        If strCustFrom = "1" Then
            strSQL = "select account_code from cust_ord_hdr Where Unit_Code='" + gstrUNITID + "' And account_code = '" & strCustCode & "'"
        Else
            strSQL = "select customer_code from customer_mst Where Unit_Code='" + gstrUNITID + "' And customer_code = '" & strCustCode & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
        End If
        If mRdoCls.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset) And mRdoCls.GetNoRows > 0 Then
            ValCustomerCode = True
        Else
            ValCustomerCode = False
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub txtamend_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtamend.GotFocus
        On Error GoTo ErrHandler
        txtamend.SelectionStart = 0
        txtamend.SelectionLength = Len(txtamend.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtamend_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtamend.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then Call cmdhelpAmend_Click(cmdhelpAmend.Item(1), New System.EventArgs())
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtamend_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtamend.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtamend_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtamend.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If Me.txtamend.Text <> "" Then
            If mRdoCls.GetResult("Select account_code from cust_ord_hdr Where Unit_Code='" + gstrUNITID + "' And account_code = '" & txtCustFrom.Text & "' and cust_ref = '" & txtSalesRefFrom.Text & "' and amendment_no = '" & txtamend.Text & "'", ADODB.CursorTypeEnum.adOpenKeyset) And mRdoCls.GetNoRows <= 0 Then
                ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO) : Me.txtamend.Text = "" : Me.txtamend.Focus() : Cancel = True : GoTo EventExitSub
            End If
        End If
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtAmendTo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAmendTo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtAmendTo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtAmendTo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim strAmendTo As String
        Dim strSQL As String
        mbolAmendDate = True
        If Me.txtCustFrom.Text <> "" And Me.txtCustTo.Text <> "" And Me.txtSalesRefFrom.Text <> "" Then
            '***********If Amendmend To is not entered then copy it as base SO************
            If Me.txtAmendTo.Text = "" Then mstrAmend = "" : mbolAmendDate = False
            ''Amend To is entered then
            If Me.txtAmendTo.Text <> "" Then
                ' amend No. and Sales Ref. No. both can not be same as Ref. SO
                If Me.txtCustFrom.Text = Me.txtCustTo.Text And Me.txtSalesRefTo.Text = txtSalesRefFrom.Text And txtamend.Text = txtAmendTo.Text Then
                    MsgBox("Ref No. and Amend both can not be same as Ref. Sales Order", MsgBoxStyle.Information, "empower")
                    mbolAmendDate = False
                    Me.txtAmendTo.Text = ""
                    Me.txtAmendTo.Focus()
                    Cancel = True
                    GoTo EventExitSub
                Else 'Check the combination is existing or not
                    strSQL = "select cust_ref,account_code from cust_ord_hdr Where Unit_Code='" + gstrUNITID + "' And account_code = '" & txtCustTo.Text & "' and cust_ref = '" & txtSalesRefTo.Text & "' and amendment_no = '" & txtAmendTo.Text & "' "
                    If mRdoCls.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset) And mRdoCls.GetNoRows > 0 Then
                        MsgBox("Can not Insert Duplicate Record ", MsgBoxStyle.Information, "empower")
                        mbolAmendDate = False
                        Me.txtAmendTo.Text = ""
                        Me.txtAmendTo.Focus()
                        Cancel = True
                        GoTo EventExitSub
                    Else
                        mstrAmend = Me.txtAmendTo.Text
                        mbolAmendDate = True
                    End If
                End If
            End If
        End If
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtCustFrom_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustFrom.GotFocus
        On Error GoTo ErrHandler
        txtCustFrom.SelectionStart = 0
        txtCustFrom.SelectionLength = Len(txtCustFrom.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustFrom_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustFrom.TextChanged
        On Error GoTo ErrHandler
        lblCustFrom.Text = ""
        Me.txtSalesRefFrom.Text = ""
        Me.txtamend.Text = ""
        Me.cmdGrpCopySO.Enabled(1) = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustFrom_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustFrom.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then Call cmdhelpCust_Click(cmdhelpCust.Item(2), New System.EventArgs()) 'Listing of Customer Codes will be displayed
        Exit Sub ' if user presses F1 Key , while cursor is in CustCode Field.
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustFrom_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustFrom.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler 'For Double and Single Code
        If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCustFrom_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustFrom.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If txtCustFrom.Text = "" Then GoTo EventExitSub
        If ValCustomerCode((txtCustFrom.Text), "1") = False Then
            ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO) : Me.txtCustFrom.Text = "" : Me.txtCustFrom.Focus() : Cancel = True : GoTo EventExitSub
        Else
            Call ProcName((txtCustFrom.Text), 1)
            cmdGrpCopySO.Enabled(1) = True
        End If
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtCustTo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustTo.GotFocus
        On Error GoTo ErrHandler
        txtCustTo.SelectionStart = 0
        txtCustTo.SelectionLength = Len(txtCustTo.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustTo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustTo.TextChanged
        On Error GoTo ErrHandler
        lblCustTo.Text = ""
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustTo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustTo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then Call cmdhelpCust_Click(cmdhelpCust.Item(3), New System.EventArgs()) 'Listing of Customer Codes will be displayed
        Exit Sub ' if user presses F1 Key , while cursor is in CustCode Field.
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustTo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustTo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCustTo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustTo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If ValCustomerCode((txtCustTo.Text), "2") = False And txtCustTo.Text <> "" Then
            ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO) : Me.txtCustTo.Text = "" : Me.txtCustTo.Focus() : Cancel = True : GoTo EventExitSub
        Else
            Call ProcName((txtCustTo.Text), 2)
        End If
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtSalesRefFrom_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSalesRefFrom.GotFocus
        On Error GoTo ErrHandler
        txtSalesRefFrom.SelectionStart = 0
        txtSalesRefFrom.SelectionLength = Len(txtSalesRefFrom.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtSalesRefFrom_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSalesRefFrom.TextChanged
        On Error GoTo ErrHandler
        If Me.txtSalesRefFrom.Text = "" Then
            Me.txtamend.Text = "" : Me.txtamend.Enabled = False : txtamend.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdhelpAmend(1).Enabled = False
            Me.txtAmendTo.Text = "" : Me.txtAmendTo.Enabled = False : txtAmendTo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtSalesRefFrom_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSalesRefFrom.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then Call cmdhelpSales_Click(cmdhelpSales.Item(0), New System.EventArgs()) 'Listing of Customer Codes will be displayed
        Exit Sub ' if user presses F1 Key , while cursor is in CustCode Field.
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtSalesRefFrom_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSalesRefFrom.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtSalesRefFrom_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtSalesRefFrom.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim strSQL As String
        If txtSalesRefFrom.Text = "" Then GoTo EventExitSub
        If Me.txtCustFrom.Text <> "" Then
            strSQL = "select cust_ref from cust_ord_hdr Where Unit_Code='" + gstrUNITID + "' And account_code = '" & txtCustFrom.Text & "' and cust_ref = '" & txtSalesRefFrom.Text & "' "
            If mRdoCls.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows <= 0 Then
                ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO) : Me.txtSalesRefFrom.Text = "" : Me.txtSalesRefFrom.Focus() : Cancel = True : GoTo EventExitSub
            End If
        Else
            strSQL = "select amendment_no from cust_ord_hdr Where Unit_Code='" + gstrUNITID + "' And account_code = " & txtCustFrom.Text & " and cust_ref = '" & txtSalesRefFrom.Text & "' and amendment_no <> '' "
            If mRdoCls.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset) And mRdoCls.GetNoRows > 0 Then
                Me.txtamend.Enabled = True : Me.cmdhelpAmend(1).Enabled = True : Me.txtamend.Focus() : txtamend.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            Else
                Me.txtCustTo.Focus()
            End If
        End If
        If Me.txtCustFrom.Text = "" Then
            MsgBox(" Pls. select Customer Code which one has to be copied ", MsgBoxStyle.Information, "empower")
            Me.txtSalesRefFrom.Text = ""
            Me.txtCustFrom.Focus()
            Cancel = True
            GoTo EventExitSub
        End If
        If Me.txtSalesRefFrom.Text <> "" Then
            strSQL = "select amendment_no from cust_ord_hdr Where Unit_Code='" + gstrUNITID + "' And account_code = '" & txtCustFrom.Text & "' and cust_ref = '" & txtSalesRefFrom.Text & "' and amendment_no <> '' "
            If mRdoCls.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset) And mRdoCls.GetNoRows > 0 Then
                Me.txtamend.Enabled = True : Me.cmdhelpAmend(1).Enabled = True : txtamend.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                Me.txtAmendTo.Enabled = True : txtAmendTo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : System.Windows.Forms.SendKeys.Send("") : GoTo EventExitSub
            Else
                Me.txtCustTo.Focus()
            End If
        Else
            Me.txtSalesRefFrom.Focus()
        End If
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtSalesRefTo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSalesRefTo.GotFocus
        On Error GoTo ErrHandler
        txtSalesRefTo.SelectionStart = 0
        txtSalesRefTo.SelectionLength = Len(txtSalesRefTo.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtSalesRefTo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSalesRefTo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtSalesRefTo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtSalesRefTo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If Me.txtSalesRefTo.Text = "" Then GoTo EventExitSub
        If Me.txtamend.Enabled = False Then
            If Me.txtCustFrom.Text = Me.txtCustTo.Text And Me.txtSalesRefFrom.Text = Me.txtSalesRefTo.Text And Len(txtamend.Text) <> 0 Then
                MsgBox(" Sales Reference No. can not be same ", MsgBoxStyle.Information, "empower") : Me.txtSalesRefTo.Text = ""
                Me.txtSalesRefTo.Focus()
                Cancel = True
                GoTo EventExitSub
            Else
                Me.txtAmendTo.Enabled = True : txtAmendTo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            End If
            If Me.txtAmendTo.Text <> "" Then
                If valReferenceNo((txtCustTo.Text), (txtSalesRefTo.Text)) = True Then
                    MsgBox(" Sales Reference No. is existing ", MsgBoxStyle.Information, "empower") : Me.txtSalesRefTo.Text = ""
                    Me.txtSalesRefTo.Focus()
                    Cancel = True
                    GoTo EventExitSub
                End If
            End If
        End If
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Public Function ProcCopy() As Boolean
        '*********************************************'
        'Author:                Ananya Nath
        'Arguments:             Custfrom ,SalesRefFrom , AmendFrom, CustTo, SalesTRefTo
        'Return Value   :       True/False
        'Description    :       Returns True, if insertion is done properly, else False.
        '*********************************************'
        On Error GoTo ErrHandler
        Dim strSQL As String
        Dim strAmendNo As String
        Dim intRows As Short
        Dim intCounter As Short
        'GST CHANGES
        'Dim strArr(32) As String
        'Dim strArr(39) As String
        Dim strArr(43) As String
        'GST CHANGES
        Dim strarr1(9) As String
        Dim strsql1 As String
        Dim strAmendTo As String
        Dim STRSQLTAX As String

        Dim rsGetDate As ClsResultSetDB
        Dim m_strSql As String
        Dim GST_COLUMNS_AV As ClsResultSetDB

        mbolCopySo = True
        'Checking for existing Items in the cust_ord_dtl
        If Me.txtamend.Text = "" Then mbolAmendDate = False
        strSQL = "select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty ,external_salesorder_no,HSNSACCODE,ISHSNORSAC,CGSTTXRT_TYPE,SGSTTXRT_TYPE,UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS,InternalAmendment,InternalAmendEnt_dt,InternalAmend_Ent_UserId from cust_ord_dtl Where Unit_Code='" + gstrUNITID + "' And account_code = '" & txtCustFrom.Text & "' and cust_ref = '" & txtSalesRefFrom.Text & "' and amendment_no = '" & txtamend.Text & "'"
        If mRdocls1.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdocls1.GetNoRows > 0 Then
            intRows = mRdocls1.GetNoRows
            mRdocls1.MoveFirst()
            For intCounter = 1 To intRows
                strArr(0) = mRdocls1.GetValue("item_code")
                strArr(1) = mRdocls1.GetValue("rate")
                strArr(2) = mRdocls1.GetValue("order_qty")
                strArr(3) = 0
                strArr(4) = mRdocls1.GetValue("active_flag")
                strArr(5) = mRdocls1.GetValue("cust_mtrl")
                strArr(6) = mRdocls1.GetValue("cust_drgno")
                strArr(7) = mRdocls1.GetValue("packing")
                strArr(8) = mRdocls1.GetValue("others")
                strArr(9) = mRdocls1.GetValue("excise_duty")
                strArr(10) = mRdocls1.GetValue("OPENSO")
                strArr(10) = IIf(strArr(10) = "False", 0, 1)
                strArr(11) = mRdocls1.GetValue("salestax_type")
                strArr(12) = mRdocls1.GetValue("cust_drg_desc")
                strArr(13) = mRdocls1.GetValue("tool_cost")
                strArr(14) = CStr(0) 'Authorized Flag
                strArr(15) = mRdocls1.GetValue("PerValue")
                If txtAmendTo.Text = "" Then
                    strArr(16) = mstrIntNoBase
                    strArr(17) = CStr(mintRevBase)
                Else
                    strArr(16) = mstrIntNoAmend
                    strArr(17) = CStr(mintRevAmend)
                End If
                strArr(18) = mRdocls1.GetValue("Remarks")
                strArr(19) = mRdocls1.GetValue("MRP")
                If strArr(19) = "" Then strArr(19) = CStr(0)
                strArr(20) = mRdocls1.GetValue("Abantment_code")
                strArr(21) = mRdocls1.GetValue("AccessibleRateforMRP")
                If strArr(21) = "" Then strArr(21) = CStr(0)
                strArr(22) = IIf(IsDBNull(mRdocls1.GetValue("Packing_Type")), "", mRdocls1.GetValue("Packing_Type"))
                If strArr(22) = "" Then strArr(22) = "PKT0"
                strArr(23) = mRdocls1.GetValue("TOOL_AMOR_FLAG")
                strArr(23) = IIf(strArr(23) = "True", "1", "0")
                strArr(24) = CStr(Val(mRdocls1.GetValue("ShowInAuth")))
                strArr(25) = IIf(IsDBNull(mRdocls1.GetValue("ADD_Excise_Duty")), "", mRdocls1.GetValue("ADD_Excise_Duty"))
                strArr(26) = CStr(Val(mRdocls1.GetValue("external_salesorder_no")))
                'GST CHANGES
                strArr(33) = CStr(mRdocls1.GetValue("HSNSACCODE"))
                strArr(34) = CStr(mRdocls1.GetValue("ISHSNORSAC"))


                GST_COLUMNS_AV = New ClsResultSetDB
                Dim strSql_GST_AV As String = ""

                strSql_GST_AV = "select * from dbo.UFN_GST_ITEMWISETAXES('" & gstrUNITID & "','" & Trim(Me.txtCustTo.Text) & "','" & strArr(0) & "',getdate(),getdate())"
                GST_COLUMNS_AV.GetResult(strSql_GST_AV)

                If GST_COLUMNS_AV.GetNoRows > 0 Then
                    strArr(35) = GST_COLUMNS_AV.GetValue("CGST_TXRT_HEAD")
                    strArr(36) = GST_COLUMNS_AV.GetValue("SGST_TXRT_HEAD")
                    strArr(37) = GST_COLUMNS_AV.GetValue("UGST_TXRT_HEAD")
                    strArr(38) = GST_COLUMNS_AV.GetValue("IGST_TXRT_HEAD")
                    strArr(39) = GST_COLUMNS_AV.GetValue("COMPENSATION_CESS")
                End If
           


                'GST CHANGES
                '*******************************************************************************************************
                'Search for combination of Customer and Item code in the CustItem_mst, if not found then copy as it is *
                '*******************************************************************************************************
                strSQL = ""
                strSQL = " select cust_drgno,drg_desc from custitem_mst Where Unit_Code='" + gstrUNITID + "' And account_code = '" & txtCustTo.Text & "' and item_code = '" & strArr(0) & "'"
                If mRdoCls.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset) And mRdoCls.GetNoRows = 0 Then
                    strSQL = ""
                    'Cust Item combination should not exist in the temp table also
                    strSQL = " select cust_drgno,drg_desc from #tmpcustitem_mst where UNIT_CODE = '" & gstrUNITID & "' AND account_code = '" & txtCustTo.Text & "' and item_code = '" & strArr(0) & "'"
                    If mRdoCls.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset) And mRdoCls.GetNoRows = 0 Then
                        strsql1 = "select Account_Code,Cust_Drgno,Drg_Desc,Item_code,Item_Desc,Active,Ent_dt,BinQuantity,VarModel,SCHUPLDREQD,CUST_MTRL,TOOL_COST,Shop_Name,Gate_No,Container,Container1,Container_qty,Commodity,Packing_Code,DELIVERY_PATTERN,POF from custitem_mst Where Unit_Code='" + gstrUNITID + "' And account_code = '" & txtCustFrom.Text & "' and item_code = '" & strArr(0) & "' "
                        If mRdoCls.GetResult(strsql1, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
                            strarr1(0) = mRdoCls.GetValue("cust_drgno")
                            strarr1(1) = mRdoCls.GetValue("drg_desc")
                            strarr1(2) = mRdoCls.GetValue("item_code")
                            strarr1(3) = mRdoCls.GetValue("item_desc")
                            strarr1(4) = mRdoCls.GetValue("active")
                            strarr1(4) = IIf(strarr1(4) = "True", "1", "0")
                            strarr1(5) = mRdoCls.GetValue("binquantity")
                            strarr1(6) = mRdoCls.GetValue("varmodel")
                            strarr1(7) = mRdoCls.GetValue("schupldreqd")
                            strarr1(7) = IIf(strarr1(7) = "True", "1", "0") '
                            strarr1(8) = CStr(Val(mRdoCls.GetValue("CUST_MTRL")))
                            strarr1(9) = CStr(Val(mRdoCls.GetValue("TOOL_COST")))
                            strSQL = ""
                            strSQL = " insert into #tmpcustitem_mst(Account_Code,Cust_Drgno,Drg_Desc,Item_code,Item_Desc,Shop_Name,Gate_No,Container,Container1,Container_qty,Active,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,binquantity,VarModel,SCHUPLDREQD,CUST_MTRL,TOOL_COST,Unit_Code,Product_Start_date, Product_End_date) values( '" & txtCustTo.Text & "' ,'" & strarr1(0) & "', '" & strarr1(1) & "','" & strarr1(2) & "','" & strarr1(3) & "','" & mRdoCls.GetValue("shop_name") & "','" & mRdoCls.GetValue("gate_no") & "','" & mRdoCls.GetValue("container") & "','" & mRdoCls.GetValue("container1") & "'," & Val(mRdoCls.GetValue("container_qty")) & "," & strarr1(4) & ", "
                            strSQL = strSQL & " getdate(), '" & mP_User & "', getdate(),'" & mP_User & "','" & strarr1(5) & "','" & strarr1(6) & "','" & strarr1(7) & "'," & strarr1(8) & "," & strarr1(9) & ",'" + gstrUNITID + "', getdate(), getdate())"
                            mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If
                    Else
                        strArr(6) = mRdoCls.GetValue("cust_drgno")
                        strArr(12) = mRdoCls.GetValue("drg_desc")
                    End If
                End If
                '*****************************************
                '*  Insertion in the #tmpcust_ord_dtl Table  *
                '*****************************************
                strSQL = ""
                'GST CHANGES
                'strSQL = " insert into #tmpcust_ord_dtl(Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty,Unit_code,external_salesorder_no)"
                strSQL = " insert into #tmpcust_ord_dtl(Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty,Unit_code,external_salesorder_no,HSNSACCODE,ISHSNORSAC,CGSTTXRT_TYPE,SGSTTXRT_TYPE,UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS,InternalAmendment,InternalAmendEnt_dt,InternalAmend_Ent_UserId)"
                'GST CHANGES
                'strSQL = strSQL & " values ('" & txtCustTo.Text & "', '" & txtSalesRefTo.Text & "', '" & txtAmendTo.Text & "','" & strArr(0) & "'," & strArr(1) & ", " & strArr(2) & "," & strArr(3) & ",'A'," & strArr(5) & ",'" & strArr(6) & "'," & strArr(7) & "," & strArr(8) & ",'" & strArr(9) & "','" & strArr(12) & "'," & strArr(13) & "," & strArr(14) & ", getdate(),'" & mP_User & "', getdate(),'" & mP_User & "'," & strArr(10) & ",'" & strArr(11) & "'," & strArr(15) & ",'" & strArr(16) & "'," & strArr(17) & ",'" & strArr(18) & "'," & Val(strArr(19)) & ",'" & strArr(20) & "'," & strArr(21) & ",'" & strArr(22) & "'," & strArr(23) & "," & strArr(24) & ",'" & strArr(25) & "','" + gstrUNITID + "','" + strArr(26) & "')"
                strSQL = strSQL & " values ('" & txtCustTo.Text & "', '" & txtSalesRefTo.Text & "', '" & txtAmendTo.Text & "','" & strArr(0) & "'," & strArr(1) & ", " & strArr(2) & "," & strArr(3) & ",'A'," & strArr(5) & ",'" & strArr(6) & "'," & strArr(7) & "," & strArr(8) & ",'" & strArr(9) & "','" & strArr(12) & "'," & strArr(13) & "," & strArr(14) & ", getdate(),'" & mP_User & "', getdate(),'" & mP_User & "'," & strArr(10) & ",'" & strArr(11) & "'," & strArr(15) & ",'" & strArr(16) & "'," & strArr(17) & ",'" & strArr(18) & "'," & Val(strArr(19)) & ",'" & strArr(20) & "'," & strArr(21) & ",'" & strArr(22) & "'," & strArr(23) & "," & strArr(24) & ",'" & strArr(25) & "','" + gstrUNITID + "','" + strArr(26) & "'"
                strSQL = strSQL & ",'" & strArr(33) & "','" & strArr(34) & "','" & strArr(35) & "','" & strArr(36) & "','" & strArr(37) & "','" & strArr(38) & "','" & strArr(39) & "','','','')"
                With mP_Connection
                    .Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End With
                mRdocls1.MoveNext()
            Next
        End If
        '************************************
        ' Checking data in the Cust_ord_hdr *
        '************************************
        strSQL = ""
        strSQL = "SELECT Account_Code,Cust_Ref,Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code,Valid_Date,Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized,Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type,SalesTax_Type,OpenSO,AddCustSupp,PerValue,InternalSONo,RevisionNo,Surcharge_code,Future_SO,ECESS_Code,Consignee_Code,ADDVAT_TYPE , CT2_Reqd_In_SO,exportsotype FROM CUST_ORD_HDR Where Unit_Code='" + gstrUNITID + "' And account_code = '" & txtCustFrom.Text & "' and cust_ref = '" & txtSalesRefFrom.Text & "' and amendment_no = '" & txtamend.Text & "'"
        If mRdoCls.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
            mRdoCls.MoveFirst() ' Copying in to an array
            strArr(0) = VB6.Format(mRdoCls.GetValue("order_date"), "dd-MMM-yyyy")
            If GetPlantName() = "HILEX" Then
                If mbolAmendDate = True Then
                    strArr(1) = VB6.Format(getDateForDB(GetServerDate()), "dd-MMM-yyyy")
                Else
                    strArr(1) = ""
                End If
            Else
                If mbolAmendDate = True Then
                    strArr(1) = VB6.Format(mRdoCls.GetValue("amendment_date"), "dd-MMM-yyyy")
                Else
                    strArr(1) = ""
                End If
            End If
            
            strArr(2) = mRdoCls.GetValue("active_flag")
            strArr(3) = mRdoCls.GetValue("OpenSO")
            strArr(3) = IIf(strArr(3) = "True", 1, 0)
            strArr(4) = mRdoCls.GetValue("AddCustSupp")
            strArr(4) = IIf(strArr(4) = "True", 1, 0)
            strArr(5) = mRdoCls.GetValue("currency_code")
            'strArr(6) = VB6.Format(mRdoCls.GetValue("valid_date"), "dd-MMM-yyyy")
            m_strSql = "Select Financial_EndDate from company_mst where unit_code='" & gstrUNITID & "'"
            rsGetDate = New ClsResultSetDB
            rsGetDate.GetResult(m_strSql)
            strArr(6) = rsGetDate.GetValue("Financial_EndDate")

            strArr(7) = VB6.Format(mRdoCls.GetValue("effect_date"), "dd-MMM-yyyy")
            strArr(8) = mRdoCls.GetValue("term_payment")
            strArr(9) = mRdoCls.GetValue("special_remarks")
            strArr(10) = mRdoCls.GetValue("pay_remarks")
            strArr(11) = mRdoCls.GetValue("price_remarks")
            strArr(12) = mRdoCls.GetValue("packing_remarks")
            strArr(13) = mRdoCls.GetValue("frieght_remarks")
            strArr(14) = mRdoCls.GetValue("transport_remarks")
            strArr(15) = mRdoCls.GetValue("octorai_remarks")
            strArr(16) = mRdoCls.GetValue("mode_despatch")
            strArr(17) = mRdoCls.GetValue("delivery")
            strArr(18) = ""
            strArr(19) = ""
            strArr(20) = ""
            strArr(21) = CStr(0)
            strArr(22) = mRdoCls.GetValue("reason")
            strArr(23) = mRdoCls.GetValue("po_type")
            strArr(24) = mRdoCls.GetValue("salestax_type")
            strArr(25) = mRdoCls.GetValue("PerValue")
            If txtAmendTo.Text = "" Then
                strArr(26) = mstrIntNoBase
                strArr(27) = CStr(mintRevBase)
            Else
                strArr(26) = mstrIntNoAmend
                strArr(27) = CStr(mintRevAmend)
            End If
            strArr(28) = mRdoCls.GetValue("surcharge_code")
            If mRdoCls.GetValue("future_so") = False Then
                strArr(29) = CStr(0)
            Else
                strArr(29) = CStr(1)
            End If
            strArr(30) = mRdoCls.GetValue("ECESS_Code")
            strArr(31) = Trim(txtconsto.Text)
            strArr(32) = mRdoCls.GetValue("CT2_Reqd_In_SO")
            strArr(40) = mRdoCls.GetValue("EXPORTSOTYPE")

            '*********************************
            'Insertion in tmpcustordhdr table
            '*********************************
            strSQL = ""
            strSQL = " Insert into #tmpcust_ord_hdr (Account_Code,Cust_Ref,Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code,Valid_Date,Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized,Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type,SalesTax_Type,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO,AddCustSupp,PerValue,InternalSONo,RevisionNo,surcharge_code,Future_SO,ECESS_Code,Consignee_Code,Unit_code,CT2_Reqd_In_SO,exportsotype) "
            strSQL = strSQL & " values ('" & txtCustTo.Text & "', '" & txtSalesRefTo.Text & "','" & txtAmendTo.Text & "', '" & strArr(0) & "'," & IIf(strArr(1) = "", "'01 JAN 1900'", "'" & strArr(1) & "'") & ",'A', '" & strArr(5) & "','" & strArr(6) & "','" & strArr(7) & "','" & strArr(8) & "','" & strArr(9) & "','" & strArr(10) & "','" & strArr(11) & "','" & strArr(12) & "','" & strArr(13) & "','" & strArr(14) & "','" & strArr(15) & "','" & strArr(16) & "','" & strArr(17) & "','" & strArr(18) & "','" & strArr(19) & "','" & strArr(20) & "','" & strArr(21) & "','" & strArr(22) & "','" & strArr(23) & "','" & strArr(24) & "',getdate(),'" & mP_User & "',getdate(),'" & mP_User & "'," & strArr(3) & ",'" & strArr(4) & "'," & strArr(25) & ",'" & strArr(26) & "'," & strArr(27) & ",'" & strArr(28) & "'," & strArr(29) & ",'" & strArr(30) & "','" & strArr(31) & "','" + gstrUNITID + "','" & strArr(32) & "','" & strArr(40) & "' )"
            With mP_Connection
                .Execute("SET DATEFORMAT 'DMY'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                .Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords) ' insertion in the Cust_ord_hdr
            End With
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mbolCopySo = False
    End Function
    Public Function ProcName(ByRef strcode As String, ByRef intNo As Short) As Object
        '*********************************************'
        'Author:                Ananya Nath
        'Arguments:             CustCode,  1 for Custfrom and 2 for CustTo
        'Return Value   :       True/False
        'Description    :       Displays the corresponding Account Description.
        '*********************************************'
        On Error GoTo ErrHandler
        Dim strSQL As String
        If strcode = "" Then Exit Function
        strSQL = " select cust_name from customer_mst Where Unit_Code='" + gstrUNITID + "' And customer_code = '" & strcode & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date)) "
        If mRdoCls.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset) And mRdoCls.GetNoRows > 0 Then
            Select Case intNo
                Case 1
                    lblCustFrom.Text = mRdoCls.GetValue("cust_name")
                Case 2
                    lblCustTo.Text = mRdoCls.GetValue("cust_name")
                Case 3
                    lblconsfrom.Text = mRdoCls.GetValue("Cust_name")
                Case 4
                    lblconsto.Text = mRdoCls.GetValue("cust_name")
            End Select
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    '*********************************************'
    'Author         :       Ananya Nath
    'Arguments      :       None
    'Return Value   :       True or False
    'Description    :       Used to validate all the mandatory fields have been properly entered.
    '*********************************************'
    Public Function ValBeforesave() As Boolean ' Checks for validity
        On Error GoTo ErrHandler
        Dim strControls As String
        Dim strFocus1 As System.Windows.Forms.Control
        Dim lNo As Integer
        ValBeforesave = True
        lNo = 1
        strControls = ResolveResString(10059) & vbCrLf
        'CustFrom
        If Len(Trim(Me.txtCustFrom.Text)) < 1 Then
            strControls = strControls & vbCrLf & lNo & ". Customer Name [Copy From]." 'Add message to String
            lNo = lNo + 1
            strFocus1 = Me.txtCustFrom
            ValBeforesave = False
        Else
            If ValCustomerCode((txtCustFrom.Text), "1") = False Then
                strControls = strControls & vbCrLf & lNo & ". Customer Name [Copy From]." 'Add message to String
                lNo = lNo + 1
                strFocus1 = Me.txtCustFrom
                ValBeforesave = False
            End If
        End If
        'CustRefFrom
        If Len(Me.txtSalesRefFrom.Text) < 1 Then
            strControls = strControls & vbCrLf & lNo & ". Sales Reference [Copy From]." 'Add message to String
            lNo = lNo + 1
            If strFocus1 Is Nothing Then strFocus1 = txtSalesRefFrom
            ValBeforesave = False
        Else
            If valReferenceNo((txtCustFrom.Text), (txtSalesRefFrom.Text)) = False Then
                strControls = strControls & vbCrLf & lNo & ". Sales Reference [Copy From]." 'Add message to String
                lNo = lNo + 1
                If strFocus1 Is Nothing Then strFocus1 = txtSalesRefFrom
                ValBeforesave = False
            End If
        End If
        'amendment
        If Me.txtamend.Text <> "" Then
            If mRdoCls.GetResult("Select account_code from cust_ord_hdr Where Unit_Code='" + gstrUNITID + "' And account_code = '" & txtCustFrom.Text & "' and cust_ref = '" & txtSalesRefFrom.Text & "' and amendment_no = '" & txtamend.Text & "'", ADODB.CursorTypeEnum.adOpenKeyset) And mRdoCls.GetNoRows <= 0 Then
                strControls = strControls & vbCrLf & lNo & ". Amendment No. [Copy From]." 'Add message to String
                lNo = lNo + 1
                If strFocus1 Is Nothing Then strFocus1 = txtamend
                ValBeforesave = False
            End If
        End If
        'Cust. to
        If Len(Trim(Me.txtCustTo.Text)) < 1 Then
            strControls = strControls & vbCrLf & lNo & ". Customer Name [Copy To]." 'Add message to String
            lNo = lNo + 1
            If strFocus1 Is Nothing Then strFocus1 = txtCustTo
            ValBeforesave = False
        Else
            If ValCustomerCode((txtCustTo.Text), "2") = False Then
                strControls = strControls & vbCrLf & lNo & ". Customer Name [Copy To]." 'Add message to String
                lNo = lNo + 1
                If strFocus1 Is Nothing Then strFocus1 = txtCustTo
                ValBeforesave = False
            End If
        End If
        'Sales Ref. To
        If Len(Me.txtSalesRefTo.Text) < 1 Then
            strControls = strControls & vbCrLf & lNo & ". Sales Reference [Copy To]." 'Add message to String
            lNo = lNo + 1
            If strFocus1 Is Nothing Then strFocus1 = txtSalesRefTo
            ValBeforesave = False
        Else
            If valReferenceNo((txtCustTo.Text), (txtSalesRefTo.Text)) = True And Me.txtAmendTo.Text = "" Then
                strControls = strControls & vbCrLf & lNo & ". Sales Reference [Copy To]." 'Add message to String
                lNo = lNo + 1
                If strFocus1 Is Nothing Then strFocus1 = txtSalesRefTo
                ValBeforesave = False
            End If
        End If
        'amendment
        If Me.txtAmendTo.Text <> "" Then
            If mRdoCls.GetResult("Select account_code from cust_ord_hdr Where Unit_Code='" + gstrUNITID + "' And account_code = '" & txtCustTo.Text & "' and cust_ref = '" & txtSalesRefTo.Text & "' and amendment_no = '" & txtAmendTo.Text & "'", ADODB.CursorTypeEnum.adOpenKeyset) And mRdoCls.GetNoRows > 0 Then
                strControls = strControls & vbCrLf & lNo & ". Amendment No. [Copy To]." 'Add message to String
                lNo = lNo + 1
                If strFocus1 Is Nothing Then strFocus1 = txtAmendTo
                ValBeforesave = False
            End If
        End If
        If ValBeforesave = False Then 'If any invalid field is there than set the focus on that field(after displaying message).
            MsgBox(strControls, MsgBoxStyle.Information, "empower")
            strFocus1.Focus()
        End If
        strFocus1 = Nothing
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub CopyBaseSO()
        ''*********************************************'
        ''Author:                Ananya Nath
        ''Arguments:             None
        ''Return Value   :       None
        ''Description    :       If uesr wants to copy an amendment no. then base so willl be copied.
        ''*********************************************'
        On Error GoTo ErrHandler
        Dim strSQL As String
        Dim strAmendNo As String
        Dim intRows As Short
        Dim intCounter As Short
        Dim strArr(40) As String
        Dim strarr1(9) As String
        Dim strsql1 As String
        Dim strAmendTo As String

        Dim rsGetDate As ClsResultSetDB
        Dim m_strSql As String
        mbolBaseSo = True
        '***************************************************************************
        'If Base SO is already existing then No need to copy it.
        '***************************************************************************
        strSQL = "select pay_remarks from cust_ord_hdr Where Unit_Code='" + gstrUNITID + "' And account_code = '" & txtCustTo.Text & "' and cust_ref = '" & txtSalesRefTo.Text & "' and amendment_no ='' "
        If mRdocls1.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdocls1.GetNoRows > 0 Then
            Exit Sub
        End If
        '******************************************************************************
        'Checking for existing Items in the cust_ord_dtl
        '******************************************************************************
        strSQL = "select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty,HSNSACCODE,ISHSNORSAC,CGSTTXRT_TYPE,SGSTTXRT_TYPE,UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS from cust_ord_dtl Where Unit_Code='" + gstrUNITID + "' And account_code = '" & txtCustFrom.Text & "' and cust_ref = '" & txtSalesRefFrom.Text & "' and amendment_no ='' "
        If mRdocls1.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdocls1.GetNoRows > 0 Then
            intRows = mRdocls1.GetNoRows
            mRdocls1.MoveFirst()
            For intCounter = 1 To intRows
                strArr(0) = mRdocls1.GetValue("item_code")
                strArr(1) = mRdocls1.GetValue("rate")
                strArr(2) = mRdocls1.GetValue("order_qty")
                strArr(3) = 0
                strArr(4) = mRdocls1.GetValue("active_flag")
                strArr(5) = mRdocls1.GetValue("cust_mtrl")
                strArr(6) = mRdocls1.GetValue("cust_drgno")
                strArr(7) = mRdocls1.GetValue("packing")
                strArr(8) = mRdocls1.GetValue("others")
                strArr(9) = mRdocls1.GetValue("excise_duty")
                strArr(10) = mRdocls1.GetValue("OPENSO")
                strArr(10) = IIf(strArr(10) = "False", 0, 1)
                strArr(11) = mRdocls1.GetValue("salestax_type")
                strArr(12) = mRdocls1.GetValue("cust_drg_desc")
                strArr(13) = mRdocls1.GetValue("tool_cost")
                strArr(14) = CStr(1) 'Authorized Flag
                strArr(15) = mRdocls1.GetValue("PerValue")
                strArr(16) = mstrIntNoBase
                strArr(17) = CStr(0)
                strArr(18) = mRdocls1.GetValue("Remarks")
                strArr(19) = mRdocls1.GetValue("MRP")
                If strArr(19) = "" Then strArr(19) = CStr(0)
                strArr(20) = mRdocls1.GetValue("Abantment_code")
                strArr(21) = mRdocls1.GetValue("AccessibleRateforMRP")
                If strArr(21) = "" Then strArr(21) = CStr(0)
                strArr(22) = IIf(IsDBNull(mRdocls1.GetValue("Packing_Type")), "", mRdocls1.GetValue("Packing_Type"))
                If strArr(22) = "" Then strArr(22) = "PKT0"
                strArr(23) = mRdocls1.GetValue("TOOL_AMOR_FLAG")
                strArr(23) = IIf(strArr(23) = "True", "1", "0")
                strArr(24) = CStr(Val(mRdocls1.GetValue("ShowInAuth")))
                strArr(25) = IIf(IsDBNull(mRdocls1.GetValue("ADD_Excise_Duty")), "", mRdocls1.GetValue("ADD_Excise_Duty"))
                'GST CHANGES
                strArr(33) = CStr(mRdocls1.GetValue("HSNSACCODE"))
                strArr(34) = CStr(mRdocls1.GetValue("ISHSNORSAC"))
                strArr(35) = CStr(mRdocls1.GetValue("CGSTTXRT_TYPE"))
                strArr(36) = CStr(mRdocls1.GetValue("SGSTTXRT_TYPE"))
                strArr(37) = CStr(mRdocls1.GetValue("UTGSTTXRT_TYPE"))
                strArr(38) = CStr(mRdocls1.GetValue("IGSTTXRT_TYPE"))
                strArr(39) = CStr(mRdocls1.GetValue("COMPENSATION_CESS"))
                'strArr(40) = CStr(mRdocls1.GetValue("EXPORTSOTYPE"))

                'GTS CHANGES
                '*****************Search for combination of Customer and Item code in the CustItem_mst, if not found then as it is
                strSQL = ""
                strSQL = " select cust_drgno,drg_desc from custitem_mst Where Unit_Code='" + gstrUNITID + "' And account_code = '" & txtCustTo.Text & "' and item_code = '" & strArr(0) & "'"
                If mRdoCls.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset) And mRdoCls.GetNoRows = 0 Then
                    strsql1 = " select Account_Code,Cust_Drgno,Drg_Desc,Item_code,Item_Desc,Active,BinQuantity,VarModel,SCHUPLDREQD,CUST_MTRL,TOOL_COST,Shop_Name,Gate_No,Container,Container1,Container_qty,Commodity,Packing_Code,DELIVERY_PATTERN,POF from custitem_mst Where Unit_Code='" + gstrUNITID + "' And account_code = '" & txtCustFrom.Text & "' and item_code = '" & strArr(0) & "' "
                    If mRdoCls.GetResult(strsql1, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic) And mRdoCls.GetNoRows > 0 Then
                        strarr1(0) = mRdoCls.GetValue("cust_drgno")
                        strarr1(1) = mRdoCls.GetValue("drg_desc")
                        strarr1(2) = mRdoCls.GetValue("item_code")
                        strarr1(3) = mRdoCls.GetValue("item_desc")
                        strarr1(4) = mRdoCls.GetValue("active")
                        strarr1(4) = IIf(strarr1(4) = "True", "1", "0")
                        strarr1(5) = CStr(Val(mRdoCls.GetValue("BinQuantity")))
                        strarr1(6) = mRdoCls.GetValue("varmodel")
                        strarr1(7) = mRdoCls.GetValue("SCHUPLDREQD")
                        strarr1(7) = IIf(strarr1(7) = "True", "1", "0")
                        strarr1(8) = CStr(Val(mRdoCls.GetValue("CUST_MTRL")))
                        strarr1(9) = CStr(Val(mRdoCls.GetValue("TOOL_COST")))
                        strSQL = ""
                        strSQL = " insert into #tmpcustitem_mst(Account_Code,Cust_Drgno,Drg_Desc,Item_code,Item_Desc,Shop_Name,Gate_No,Container,Container1,Container_qty,Active,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,binquantity,VarModel,SCHUPLDREQD,CUST_MTRL,TOOL_COST,Unit_Code,Product_Start_date, Product_End_date) values( '" & txtCustTo.Text & "' ,'" & strarr1(0) & "', '" & strarr1(1) & "','" & strarr1(2) & "','" & strarr1(3) & "','" & mRdoCls.GetValue("shop_name") & "','" & mRdoCls.GetValue("gate_no") & "','" & mRdoCls.GetValue("container") & "','" & mRdoCls.GetValue("container1") & "'," & Val(mRdoCls.GetValue("container_qty")) & "," & strarr1(4) & ", "
                        strSQL = strSQL & " getdate(), '" & mP_User & "', getdate(),'" & mP_User & "'," & strarr1(5) & ",'" & strarr1(6) & "'," & strarr1(7) & "," & strarr1(8) & "," & strarr1(9) & ",'" + gstrUNITID + "',getdate(),getdate())"
                        mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                Else
                    strArr(6) = mRdoCls.GetValue("cust_drgno")
                    strArr(12) = mRdoCls.GetValue("drg_desc")
                End If
                '*****************************************
                '*  Insertion in the #tmpcust_ord_dtl Table  *
                '*****************************************
                strSQL = ""
                strSQL = " insert into #tmpcust_ord_dtl(Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty,Unit_Code,HSNSACCODE,ISHSNORSAC,CGSTTXRT_TYPE,SGSTTXRT_TYPE,UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS,InternalAmendment,InternalAmendEnt_dt,InternalAmend_Ent_UserId)"
                'GST CHANGES
                'strSQL = strSQL & " values ('" & txtCustTo.Text & "', '" & txtSalesRefTo.Text & "', '', '" & strArr(0) & "'," & strArr(1) & ", " & strArr(2) & "," & strArr(3) & ",'" & strArr(4) & "'," & strArr(5) & ",'" & strArr(6) & "'," & strArr(7) & "," & strArr(8) & ",'" & strArr(9) & "','" & strArr(12) & "'," & strArr(13) & "," & strArr(14) & ", getdate(),'" & mP_User & "', getdate(),'" & mP_User & "'," & strArr(10) & ",'" & strArr(11) & "'," & strArr(15) & ",'" & strArr(16) & "'," & strArr(17) & ",'" & strArr(18) & "'," & Val(strArr(19)) & ",'" & strArr(20) & "'," & strArr(21) & ",'" & strArr(22) & "'," & strArr(23) & "," & strArr(24) & ",'" & strArr(25) & "','" + gstrUNITID + "')"
                strSQL = strSQL & " values ('" & txtCustTo.Text & "', '" & txtSalesRefTo.Text & "', '', '" & strArr(0) & "'," & strArr(1) & ", " & strArr(2) & "," & strArr(3) & ",'" & strArr(4) & "'," & strArr(5) & ",'" & strArr(6) & "'," & strArr(7) & "," & strArr(8) & ",'" & strArr(9) & "','" & strArr(12) & "'," & strArr(13) & "," & strArr(14) & ", getdate(),'" & mP_User & "', getdate(),'" & mP_User & "'," & strArr(10) & ",'" & strArr(11) & "'," & strArr(15) & ",'" & strArr(16) & "'," & strArr(17) & ",'" & strArr(18) & "'," & Val(strArr(19)) & ",'" & strArr(20) & "'," & strArr(21) & ",'" & strArr(22) & "'," & strArr(23) & "," & strArr(24) & ",'" & strArr(25) & "','" + gstrUNITID + "'"
                strSQL = strSQL & ",'" & strArr(33) & "','" & strArr(34) & "','" & strArr(35) & "','" & strArr(36) & "','" & strArr(37) & "','" & strArr(38) & "','" & strArr(39) & "','','','')"
                'GST CHANGES
                With mP_Connection
                    .Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End With
                mRdocls1.MoveNext()
            Next
        End If
        strSQL = ""
        strSQL = " select Account_Code,Cust_Ref,Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code,Valid_Date,Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized,Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type,SalesTax_Type,OpenSO,AddCustSupp,PerValue,InternalSONo,RevisionNo,Surcharge_code,Future_SO,ECESS_Code,Consignee_Code,ADDVAT_TYPE ,EXPORTSOTYPE from cust_ord_hdr Where Unit_Code='" + gstrUNITID + "' And account_code = '" & txtCustFrom.Text & "' and cust_ref = '" & txtSalesRefFrom.Text & "' and amendment_no ='' "
        If mRdoCls.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
            mRdoCls.MoveFirst() ' Copying in to an array
            strArr(0) = VB6.Format(mRdoCls.GetValue("order_date"), "dd-MMM-yyyy")
            strArr(1) = ""
            strArr(2) = mRdoCls.GetValue("active_flag")
            strArr(3) = mRdoCls.GetValue("OpenSO")
            strArr(3) = IIf(strArr(3) = "True", 1, 0)
            strArr(4) = mRdoCls.GetValue("AddCustSupp")
            strArr(4) = IIf(strArr(4) = "True", 1, 0)
            strArr(5) = mRdoCls.GetValue("currency_code")
            'strArr(6) = VB6.Format(mRdoCls.GetValue("valid_date"), "dd-MMM-yyyy")
            m_strSql = "Select Financial_EndDate from company_mst where unit_code='" & gstrUNITID & "'"
            rsGetDate = New ClsResultSetDB
            rsGetDate.GetResult(m_strSql)
            strArr(6) = rsGetDate.GetValue("Financial_EndDate")

            strArr(7) = VB6.Format(mRdoCls.GetValue("effect_date"), "dd-MMM-yyyy")
            strArr(8) = mRdoCls.GetValue("term_payment")
            strArr(9) = mRdoCls.GetValue("special_remarks")
            strArr(10) = mRdoCls.GetValue("pay_remarks")
            strArr(11) = mRdoCls.GetValue("price_remarks")
            strArr(12) = mRdoCls.GetValue("packing_remarks")
            strArr(13) = mRdoCls.GetValue("frieght_remarks")
            strArr(14) = mRdoCls.GetValue("transport_remarks")
            strArr(15) = mRdoCls.GetValue("octorai_remarks")
            strArr(16) = mRdoCls.GetValue("mode_despatch")
            strArr(17) = mRdoCls.GetValue("Delivery")
            strArr(18) = mRdoCls.GetValue("First_authorized")
            strArr(19) = mRdoCls.GetValue("Second_authorized")
            strArr(20) = mRdoCls.GetValue("Third_authorized")
            strArr(21) = CStr(1)
            strArr(22) = mRdoCls.GetValue("reason")
            strArr(23) = mRdoCls.GetValue("po_type")
            strArr(24) = mRdoCls.GetValue("salestax_type")
            strArr(25) = mRdoCls.GetValue("PerValue")
            strArr(26) = mstrIntNoBase
            strArr(27) = CStr(0)
            strArr(28) = mRdoCls.GetValue("Surcharge_code")
            If mRdoCls.GetValue("future_so") = False Then
                strArr(29) = CStr(0)
            Else
                strArr(29) = CStr(1)
            End If
            strArr(30) = mRdoCls.GetValue("ECESS_Code")
            strArr(31) = Trim(txtconsto.Text)
            'strArr(40) = mRdoCls1.GetValue("EXPORTSOTYPE")
            '*********************************
            'Insertion in #tmpcust_ord_hdr table
            '*********************************
            strSQL = ""
            strSQL = " Insert into #tmpcust_ord_hdr (Account_Code,Cust_Ref,Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code,Valid_Date,Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized,Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type,SalesTax_Type,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO,AddCustSupp,PerValue,InternalSONo,RevisionNo,surcharge_code,Future_SO,ECESS_Code,Consignee_Code,Unit_code,exportsotype) "
            strSQL = strSQL & " values ('" & txtCustTo.Text & "', '" & txtSalesRefTo.Text & "', '', '" & strArr(0) & "'," & IIf(strArr(1) = "", "NULL", "'" & strArr(1) & "'") & ",'" & strArr(2) & "', '" & strArr(5) & "','" & strArr(6) & "','" & strArr(7) & "','" & strArr(8) & "','" & strArr(9) & "','" & strArr(10) & "','" & strArr(11) & "','" & strArr(12) & "','" & strArr(13) & "','" & strArr(14) & "','" & strArr(15) & "','" & strArr(16) & "','" & strArr(17) & "','" & strArr(18) & "','" & strArr(19) & "','" & strArr(20) & "','" & strArr(21) & "','" & strArr(22) & "','" & strArr(23) & "','" & strArr(24) & "',getdate(),'" & mP_User & "',getdate(),'" & mP_User & "'," & strArr(3) & ",'" & strArr(4) & "'," & strArr(25) & ",'" & strArr(26) & "'," & strArr(27) & ",'" & strArr(28) & "'," & strArr(29) & ",'" & strArr(30) & "','" & strArr(31) & "','" + gstrUNITID + "','" & strArr(40) & "')"
            With mP_Connection
                .Execute("SET DATEFORMAT 'DMY'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                .Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords) ' insertion in the #tmpcust_ord_hdr
            End With
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mbolBaseSo = False
    End Sub
    ''*********************************************'
    ''Author:                Ananya Nath
    ''Arguments:             None
    ''Return Value   :       None
    ''Description    :       Creates # tables.
    ''*********************************************'

    Public Function ProcCreateTab() As Object
        On Error GoTo ErrHandler
        Dim strSQL As String
        strSQL = "SET ROWCOUNT 1" & vbCrLf
        strSQL = strSQL & " select * into #tmpcust_ord_hdr from cust_ord_hdr Where Unit_Code='" + gstrUNITID + "'" & vbCrLf
        strSQL = strSQL & " select * into #tmpcust_ord_dtl from cust_ord_dtl Where Unit_Code='" + gstrUNITID + "'" & vbCrLf
        'To Update the show authorization by 1
        strSQL = strSQL & " Update #tmpcust_ord_dtl set ShowInAuth = 1 where unit_code = '" & gstrUNITID & "'" & vbCrLf
        strSQL = strSQL & " select * into #tmpcustitem_mst from custitem_mst Where Unit_Code='" + gstrUNITID + "'" & vbCrLf
        strSQL = strSQL & " SET ROWCOUNT 0 "

        mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Public Function GenerateDocumentNumber(ByVal pstrTableName As String, ByVal pstrDocNofield As String, ByRef pstrDateFieldName As String, ByVal pstrWantedDate As String) As String
        ''*********************************************'
        ''Author:                Nisha
        ''Arguments:             Tablename, fieldname, ent_dt, serverdate
        ''Return Value   :       Doc No.[Interbnal No.]
        ''Description    :       Generates Internal No.
        ''*********************************************'
        On Error GoTo ErrHandler
        Dim strCheckDOcNo As String 'Gets the Doc Number from Back End
        Dim strTempSeries As String 'Find the Numeric series in Doc No
        Dim NewTempSeries As String 'Generate a NEW Series
        Dim rsDocumentNoSO As ClsResultSetDB
        rsDocumentNoSO = New ClsResultSetDB
        If Len(Trim(pstrWantedDate)) > 0 Then 'For Post Dated Docs
            'No need to check for Previously made documents for After Dates
            mP_Connection.Execute("Set dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            rsDocumentNoSO.GetResult("Select DocNo = Max(convert(int,substring(" & pstrDocNofield & ",9,7))) from " & pstrTableName & " Where Unit_Code='" + gstrUNITID + "' And datePart(mm,ent_dt) = datePart(mm,'" & pstrWantedDate & "') and datePart(yyyy,ent_dt) = datePart(yyyy,'" & pstrWantedDate & "')")
            strCheckDOcNo = rsDocumentNoSO.GetValue("DocNo")
        End If
        If Len(Trim(strCheckDOcNo)) > 0 Then 'That is the Document is Made for that Period
            'Add 1 to it
            strTempSeries = CStr(CDbl(strCheckDOcNo) + 1)
            If Val(strTempSeries) < 9999 Then
                strTempSeries = New String("0", 4 - Len(strTempSeries)) & strTempSeries 'Concatenate Zeroes before the Number
            End If
            strCheckDOcNo = CStr(DatePart(Microsoft.VisualBasic.DateInterval.Year, CDate(pstrWantedDate))) & "-"
            strCheckDOcNo = strCheckDOcNo & New String("0", 2 - Len(CStr(DatePart(Microsoft.VisualBasic.DateInterval.Month, CDate(pstrWantedDate))))) & CStr(DatePart(Microsoft.VisualBasic.DateInterval.Month, CDate(pstrWantedDate))) & "-"
            strCheckDOcNo = strCheckDOcNo & strTempSeries
            GenerateDocumentNumber = strCheckDOcNo
        Else 'The Document has not been made for that Period
            NewTempSeries = NewTempSeries & CStr(DatePart(Microsoft.VisualBasic.DateInterval.Year, CDate(pstrWantedDate))) & "-"
            NewTempSeries = NewTempSeries & New String("0", 2 - Len(CStr(DatePart(Microsoft.VisualBasic.DateInterval.Month, CDate(pstrWantedDate))))) & CStr(DatePart(Microsoft.VisualBasic.DateInterval.Month, CDate(pstrWantedDate))) & "-"
            NewTempSeries = NewTempSeries & "0001"
            GenerateDocumentNumber = NewTempSeries 'The Number Is Generated
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Function
    End Function
    Public Function GenerateRevisionNo() As Short
        ''*********************************************'
        ''Author:                Nisha
        ''Arguments:             None
        ''Return Value   :       Revision No.
        ''Description    :       Generates Revision No.
        ''*********************************************'
        On Error GoTo ErrHandler
        Dim rsRevisionNo As ClsResultSetDB
        rsRevisionNo = New ClsResultSetDB
        rsRevisionNo.GetResult("Select Revision = Max(RevisionNo) from Cust_ord_hdr Where Unit_Code='" + gstrUNITID + "' And cust_ref = '" & txtSalesRefTo.Text & "'")
        GenerateRevisionNo = IIf(IsDBNull(rsRevisionNo.GetValue("Revision")), 1, Val(rsRevisionNo.GetValue("Revision")) + 1)
        rsRevisionNo = Nothing
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Function
    End Function
    Public Sub InteralNo()
        ''****************************************************************
        ''Author:                Ananya Nath
        ''Arguments:             None
        ''Return Value   :       None
        ''Description    :       Returns Internal No(s) and Revision No.
        ''*****************************************************************
        On Error GoTo ErrHandler
        Dim strSQL As String
        If txtAmendTo.Text = "" Then
            mstrIntNoBase = GenerateDocumentNumber("cust_ord_hdr", "InternalSono", "ent_dt", getDateForDB(GetServerDate()))
            mintRevBase = 0
        Else
            strSQL = "select cust_ref,internalsono, max(revisionno) from cust_ord_hdr Where Unit_Code='" + gstrUNITID + "' And account_code = '" & txtCustTo.Text & "' and cust_ref = '" & txtSalesRefTo.Text & "' group by cust_ref,internalsono"
            If mRdoCls.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows <= 0 Then
                mstrIntNoBase = GenerateDocumentNumber("cust_ord_hdr", "InternalSono", "ent_dt", getDateForDB(GetServerDate()))
                mintRevBase = 0
                mstrIntNoAmend = mstrIntNoBase
                mintRevAmend = 1
            Else
                mstrIntNoAmend = mRdoCls.GetValue("internalsono")
                mintRevAmend = GenerateRevisionNo()
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLPMKTTRN0011.htm")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdGrpCopySO_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.cmdGrpAuthorise.ButtonClickEventArgs) Handles cmdGrpCopySO.ButtonClick
        On Error GoTo ErrHandler
        Dim strSQL As String
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_AUTHORIZE
                If ValBeforesave() = True Then
                    If Me.txtamend.Enabled = True Then
                        If (Me.txtamend.Text <> "" And Me.txtAmendTo.Text = "") Or (Me.txtamend.Text = "" And Me.txtAmendTo.Text <> "") Then
                            If MsgBox("Amendment No. of copied Sales Order is blank, Do you Wish to Contd...", 321, "empower") = MsgBoxResult.Cancel Then
                                If Me.txtamend.Text = "" Then Me.txtamend.Focus() Else Me.txtAmendTo.Focus()
                                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                                Exit Sub
                            End If
                        End If
                    End If
                    '***************************
                    '*   Delete from tmptables *
                    '***************************
                    strSQL = "delete from #tmpcustitem_mst where unit_code = '" & gstrUNITID & "' " & vbCrLf & " delete from #tmpcust_ord_dtl where unit_code = '" & gstrUNITID & "' " & vbCrLf & " delete from #tmpcust_ord_hdr where unit_code = '" & gstrUNITID & "' "
                    mP_Connection.BeginTrans()
                    mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    mP_Connection.CommitTrans()
                    Call InteralNo()
                    '****************************************************************
                    'IF CUSTFROM, TO AND REFE. AND AMEND EXISTS THEN COPY the Base SO
                    '****************************************************************
                    If Me.txtAmendTo.Text <> "" Then
                        strSQL = "select Account_Code,Cust_Ref,Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code,Valid_Date,Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized,Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type,SalesTax_Type,OpenSO,AddCustSupp,PerValue,InternalSONo,RevisionNo,Surcharge_code,Future_SO,ECESS_Code,Consignee_Code,ADDVAT_TYPE from cust_ord_hdr Where Unit_Code='" + gstrUNITID + "' And account_code = '" & txtCustTo.Text & "' and cust_ref = '" & txtSalesRefTo.Text & "'"
                        If mRdoCls.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows <= 0 Then
                            Call CopyBaseSO()
                        Else : mbolBaseSo = True
                        End If
                    Else : mbolBaseSo = True
                    End If
                    '***************************
                    '* Calling for other Cases *
                    '***************************
                    Call ProcCopy()
                    Dim strMessage As String = ""
                    If mbolCopySo = True And mbolBaseSo = True Then

                        strSQL = "INSERT INTO CUSTITEM_MST SELECT * FROM #tmpcustitem_mst where unit_code = '" & gstrUNITID & "'" & vbCrLf & "INSERT INTO CUST_ORD_HDR SELECT * FROM #tmpcust_ord_hdr where unit_code = '" & gstrUNITID & "' " & vbCrLf & " INSERT INTO CUST_ORD_DTL SELECT * FROM #tmpcust_ord_dtl where unit_code = '" & gstrUNITID & "' "
                        mP_Connection.BeginTrans()
                            mP_Connection.Execute("SET DATEFORMAT 'DMY'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            mP_Connection.CommitTrans()


                        '' added by priti on 20 May for MTL customer drg no inactive issue
                        If gstrUNITID = "MS1" Then
                            Dim strExist As Boolean = True

                            strSQL = "Select * from cust_ord_dtl  a,CustItem_Mst b  where a.Account_Code=b.Account_Code and a.Item_Code=b.Item_code and active=1 and a.unit_code = '" & gstrUNITID & "' and a.account_code='" & txtCustTo.Text & "' and a.cust_ref='" & txtSalesRefTo.Text & "' and Amendment_No ='" & txtAmendTo.Text & "'"
                            If mRdoCls.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows <= 0 Then
                                strExist = False
                            End If

                            mP_Connection.BeginTrans()
                            strSQL = "Delete a from cust_ord_dtl  a,CustItem_Mst b  where a.Account_Code=b.Account_Code And a.Item_Code=b.Item_code And active=0 And a.unit_code = '" & gstrUNITID & "' and a.account_code='" & txtCustTo.Text & "' and a.cust_ref='" & txtSalesRefTo.Text & "'  and Amendment_No ='" & txtAmendTo.Text & "'"
                            mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                            If strExist = False Then
                                strSQL = "Delete from cust_ord_hdr  where unit_code = '" & gstrUNITID & "' and account_code='" & txtCustTo.Text & "'  and Amendment_No ='" & txtAmendTo.Text & "'"
                                mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                strMessage = "Sale Order cannot be copied. No active customer drg no in SO ."
                            End If
                            mP_Connection.CommitTrans()
                        End If
                        If strMessage <> "" Then
                            MsgBox(strMessage, MsgBoxStyle.Information, "empower")
                        Else
                            MsgBox("Sales Order is Successfully Copied", MsgBoxStyle.Information, "empower")
                        End If
                        '' End by priti on 20 May for MTL customer drg no inactive issue


                    Else
                        MsgBox("Could not Copy the Sales Order", MsgBoxStyle.Information, "empower")
                    End If
                    Me.txtCustFrom.Focus()
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_REFRESH
                Me.txtCustFrom.Text = ""
                Me.txtCustTo.Text = ""
                Me.txtamend.Text = "" : Me.txtamend.Enabled = False : Me.cmdhelpAmend(1).Enabled = False : txtamend.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                Me.txtSalesRefFrom.Text = ""
                Me.txtSalesRefTo.Text = ""
                Me.txtCustFrom.Focus()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
                Me.Dispose()
        End Select
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        mP_Connection.RollbackTrans()
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
    End Sub
End Class