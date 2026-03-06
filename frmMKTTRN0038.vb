Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0038
	Inherits System.Windows.Forms.Form
	'---------------------------------------------------------------------------
	'Copyright          :   MIND Ltd.
	'Form Name          :   frmMKTTRN0038
	'Created By         :   Sourabh Khatri
	'Created on         :   16/11/2004
	'Modified Date      :
    'Description        :   Form shall work to cancel all the pending schedule for the selected date range and selected customers.
    'MODIFIED BY AJAY SHUKLA ON 10/MAY/2011 FOR MULTIUNIT CHANGE
	'---------------------------------------------------------------------------
	Dim mlngFormTag As Short ' Variable For Header String
    Private Sub cmdScheduleCancellation_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles cmdScheduleCancellation.ButtonClick
        On Error GoTo Errorhandler
        Dim Intcounter As Short
        Dim cmdObject As New ADODB.Command
        Dim strString As String
        Dim rsObject As New ClsResultSetDB
        Dim intCounter1 As Short
        Select Case e.Button

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                Call InitializeControls()

                Call optAllCustomer_CheckedChanged(optAllCustomer, New System.EventArgs())

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                Call frmMKTTRN0038_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.WaitCursor)
                If ValidateData() Then

                    With cmdObject
                        .let_ActiveConnection(mP_Connection)
                        .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        .CommandTimeout = 0
                    End With
                    mP_Connection.BeginTrans()
                    If Me.optAllCustomer.Checked = True Then
                        strString = "Select distinct Customer_Code from Customer_mst WHERE UNIT_CODE='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                        rsObject.GetResult(strString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If Not rsObject.EOFRecord Then
                            rsObject.MoveFirst()
                            While Not rsObject.EOFRecord

                                strString = "SCHEDULECANCELLATION('" & gstrUNITID & "','" & rsObject.GetValue("Customer_Code") & "')"
                                cmdObject.CommandText = strString
                                cmdObject.Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                rsObject.MoveNext()
                            End While
                            rsObject.ResultSetClose()
                            rsObject = Nothing
                        Else
                            MsgBox(" Customer not Exist in Customer Master.Please Check Customer Master", MsgBoxStyle.Information, "eMPro")
                            Exit Sub
                        End If

                    Else
                        For Intcounter = 0 To Me.LstCustomerSel.Items.Count - 1

                            If Me.LstCustomerSel.Items.Item(Intcounter).Checked = True Then

                                strString = "SCHEDULECANCELLATION('" & gstrUNITID & "','" & Me.LstCustomerSel.Items.Item(Intcounter).Text & "')"
                                cmdObject.CommandText = strString
                                cmdObject.Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                            End If
                        Next
                    End If
                    mP_Connection.CommitTrans()

                    cmdObject = Nothing
                    MsgBox(" Transaction Completed Successfully ", MsgBoxStyle.Information, "eMPro")
                    Me.txtSearch.Text = "" : Me.optAllCustomer.Checked = True
                    Me.cmdScheduleCancellation.Revert()
                    Call InitializeControls()

                End If
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
        End Select

        Exit Sub
Errorhandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
        mP_Connection.RollbackTrans()
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub frmMKTTRN0038_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '----------------------------------------------------
        'Author              - Sourabh Khatri
        'Create Date         - 16/11/2004
        'Arguments           - None
        'Return Value        - None
        'Function            - To intialise required
        '----------------------------------------------------
        On Error GoTo Errorhandler
        mdifrmMain.CheckFormName = mlngFormTag
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub frmMKTTRN0038_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo errHandler
        'Make the node normal font
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0038_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo errHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                'If user press the ESC Key ,the Form will be in View Mode
                If Me.cmdScheduleCancellation.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call Me.cmdScheduleCancellation.Revert()
                        Call InitializeControls()
                    End If
                End If
        End Select
        GoTo EventExitSub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub frmMKTTRN0038_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo Errorhandler
        '----------------------------------------------------
        'Author              - Sourabh Khatri
        'Create Date         - 16/11/2004
        'Arguments           - None
        'Return Value        - None
        'Function            - To intialise required
        '----------------------------------------------------
        mlngFormTag = mdifrmMain.AddFormNameToWindowList(Me.ctlFormHeader.Tag)
        Call FitToClient(Me, frmMain, (Me.ctlFormHeader), (Me.cmdScheduleCancellation))
        Call InitializeControls()
        Call Me.cmdScheduleCancellation.ShowButtons(True, False, False, False)
        'Call PopulateListView
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0038_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo errHandler
        'Removing the form name from list
        mdifrmMain.RemoveFormNameFromWindowList = mlngFormTag
        'Setting the corresponding node's tag
        frmModules.NodeFontBold(Tag) = False
        Me.Dispose()
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub InitializeControls()
        '----------------------------------------------------
        'Author              - Sourabh Khatri
        'Create Date         - 16/11/2004
        'Arguments           - None
        'Return Value        - None
        'Function            - To intialise required
        '----------------------------------------------------
        Me.optAllCustomer.Checked = True
        Me.optSerachCOde.Enabled = False
        Me.optSerachName.Enabled = False
        Me.txtSearch.Enabled = False
        Me.txtSearch.Text = ""
        Me.LstCustomerSel.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
    End Sub
    Public Sub PopulateListView()
        Dim rsObject As ClsResultSetDB
        Dim strString As String
        Dim Intcounter As Short
        On Error GoTo Errorhandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
        strString = "Select Customer_Code,Cust_Name from Customer_mst WHERE UNIT_CODE='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date)) order by customer_Code"
        rsObject = New ClsResultSetDB
        rsObject.GetResult(strString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

        With Me.LstCustomerSel
            .Columns.Clear()
            .Columns.Insert(0, "", " Customer Code ", -2)
            .Columns.Insert(1, "", " Customer Name ", -2)
            .Items.Clear()
            .Sort()
            .LabelEdit = False
            .View = System.Windows.Forms.View.Details
            If Me.optAllCustomer.Checked = True Then
                .Enabled = False
            Else
                .Enabled = True
            End If
            .CheckBoxes = True
        End With
        If rsObject.RowCount > 0 Then

            rsObject.MoveFirst()
            With Me.LstCustomerSel
                For Intcounter = 0 To rsObject.RowCount - 1
                    .Items.Insert(Intcounter, rsObject.GetValue("customer_Code"))
                    .Items.Item(Intcounter).SubItems.Add(rsObject.GetValue("Cust_Name"))
                    rsObject.MoveNext()
                Next
                .Columns.Item(0).Width = 100
                .Columns.Item(1).Width = 400
            End With
        End If
        rsObject = Nothing
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Exit Sub
Errorhandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub optAllCustomer_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAllCustomer.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo Errorhandler
            Dim Intcounter As Short
            Me.LstCustomerSel.Enabled = False
            Me.LstCustomerSel.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Me.LstCustomerSel.Items.Clear()
            Exit Sub
Errorhandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub

    Private Sub optSelectedCustomer_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSelectedCustomer.CheckedChanged
        If eventSender.Checked Then
            Dim Intcounter As Short
            On Error GoTo Errorhandler
            If Me.cmdScheduleCancellation.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                Me.LstCustomerSel.Items.Clear()
                PopulateListView()
                If Me.LstCustomerSel.Items.Count > 0 Then
                    If Me.cmdScheduleCancellation.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                        Me.optAllCustomer.Checked = True : Exit Sub
                    End If
                    Me.LstCustomerSel.Enabled = True
                    Me.optSerachCOde.Enabled = True
                    Me.optSerachName.Enabled = True
                    Me.txtSearch.Enabled = True
                    Me.frmSeracgOptions.Enabled = True
                    Me.frmSerachOp.Enabled = True
                    Me.LstCustomerSel.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    Me.optSerachCOde.Checked = True
                End If
                For Intcounter = 0 To Me.LstCustomerSel.Items.Count - 1
                    Me.LstCustomerSel.Items.Item(Intcounter).Checked = False
                Next
            End If
            Exit Sub
Errorhandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub SearchValue(ByRef lvwView As System.Windows.Forms.ListView, ByRef txtBox As System.Windows.Forms.TextBox, ByRef optButton1 As System.Windows.Forms.RadioButton, ByRef optButton2 As System.Windows.Forms.RadioButton)
        '*******************************************************************************
        'Author             :   Sourabh Khatri
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Add Required code on Form Load
        'Comments           :   NA
        'Creation Date      :   11/02/2004
        '*******************************************************************************
        On Error GoTo Errorhandler
        Dim Intcounter As Short

        With lvwView
            For Intcounter = 0 To .Items.Count - 1
                If .Items.Item(Intcounter).Font.Bold = True Then
                    .Items.Item(Intcounter).Font = VB6.FontChangeBold(.Items.Item(Intcounter).Font, False)
                    .Refresh()
                End If
                If .Items.Item(Intcounter).SubItems.Item(1).Font.Bold = True Then
                    .Items.Item(Intcounter).SubItems.Item(1).Font = VB6.FontChangeBold(.Items.Item(Intcounter).SubItems.Item(1).Font, False)
                    .Refresh()
                End If
            Next
            If optButton1.Checked = True Then
                If Len(txtBox.Text) < 1 Then Exit Sub
                For Intcounter = 0 To lvwView.Items.Count - 1
                    If Trim(UCase(Mid(.Items.Item(Intcounter).Text, 1, Len(txtBox.Text)))) = Trim(UCase(txtBox.Text)) Then
                        .Items.Item(Intcounter).Font = VB6.FontChangeBold(.Items.Item(Intcounter).Font, True)
                        Call .Items.Item(Intcounter).EnsureVisible()
                        .Refresh()
                        Exit Sub
                    End If
                Next
            End If

            If optButton2.Checked = True Then
                If Len(txtBox.Text) < 1 Then Exit Sub
                For Intcounter = 0 To lvwView.Items.Count - 1
                    If Trim(UCase(Mid(.Items.Item(Intcounter).SubItems.Item(1).Text, 1, Len(txtBox.Text)))) = Trim(UCase(txtBox.Text)) Then
                        .Items.Item(Intcounter).Font = VB6.FontChangeBold(.Items.Item(Intcounter).Font, True)
                        Call .Items.Item(Intcounter).EnsureVisible()
                        .Refresh()
                        Exit Sub
                    End If
                Next
            End If
        End With
        Exit Sub
Errorhandler:
        If Err.Number = 35600 Then Resume Next
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub optSerachCOde_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSerachCOde.CheckedChanged
        If eventSender.Checked Then
            If Me.txtSearch.Enabled = True Then
                Me.txtSearch.Text = ""
            End If
        End If
    End Sub

    Private Sub optSerachName_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSerachName.CheckedChanged
        If eventSender.Checked Then
            If Me.txtSearch.Enabled = True Then
                Me.txtSearch.Text = ""
            End If
        End If
    End Sub

    Private Sub TxtSearch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSearch.TextChanged
        Call SearchValue((Me.LstCustomerSel), (Me.txtSearch), (Me.optSerachCOde), (Me.optSerachName))
    End Sub
    Public Function ValidateData() As Boolean
        On Error GoTo Errorhandler
        Dim lstrControls As String
        Dim lNo As Integer
        Dim Intcounter As Short
        Dim intCounterItems As Short
        Dim lctrFocus As System.Windows.Forms.Control
        Dim strString As String
        Dim rsObject As New ClsResultSetDB
        Dim strMsg As String

        ValidateData = True
        lstrControls = ResolveResString(10059) & vbCrLf
        lNo = 1
        For Intcounter = 0 To Me.LstCustomerSel.Items.Count - 1

            If Me.LstCustomerSel.Items.Item(Intcounter).Checked = True Then
                Intcounter = 0 : Exit For
            End If
        Next
        If Intcounter <> 0 Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Select At Least One Customer For Cancellation"
            lNo = lNo + 1
            ValidateData = False
        End If

        'Validation for Check invoice exist or not after the last date of schedule cancellation ( If exist schedule cancellation will not work )
        For Intcounter = 0 To Me.LstCustomerSel.Items.Count - 1

            If Me.LstCustomerSel.Items.Item(Intcounter).Checked = True Then

                strString = " Select item_Code from mkt_invdshistory where customer_Code = '" & Trim(Me.LstCustomerSel.Items.Item(Intcounter).Text) & "' and cancellation_flag = 0 and ent_dt > Convert(varchar(11),Dateadd(D,1,Convert(varchar(12),(dateadd(d,-day(getdate()),getdate())))),106) AND UNIT_CODE='" & gstrUNITID & "'"
                rsObject.GetResult(strString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rsObject.RowCount > 0 Then

                    lstrControls = lstrControls & vbCrLf & lNo & ".Invoicing has been done for Customer '" & Me.LstCustomerSel.Items.Item(Intcounter).Text & "' in current month. Schedule can't be rollover for customer :- " & Me.LstCustomerSel.Items.Item(Intcounter).Text
                    lNo = lNo + 1
                    ValidateData = False
                End If
            End If
        Next
        If ValidateData = False Then
            MsgBox(lstrControls, MsgBoxStyle.Information, "eMPro")
        End If
		Exit Function
Errorhandler: 
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Sub ctlFormHeader_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader.Click
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Set Property CheckFormName as per Rule book
        ' Datetime      : 16 april 2005
        '--------------------------------------------------------------------
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        Exit Sub    'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub
End Class