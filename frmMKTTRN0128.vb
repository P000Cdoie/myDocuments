Option Strict Off
Option Explicit On
Imports System.Data.SqlClient

Friend Class frmMKTTRN0128
    Inherits System.Windows.Forms.Form
    Dim mintIndex As Short 'Variable used to Store the Index of the Form in the List View
    Dim mStrDes As String ' Variable to Store Description

    Private Sub cmdGrpPacking_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles cmdASN.ButtonClick
        Try
            Select Case e.Button 'Checks the Button pressed by the user
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD  'ADD BUTTON Pressed
                    Me.txtInvoiceNo.Tag = Me.txtInvoiceNo.Text 'To store the value of the current code so that we revert back if user cancels the operation
                    cmdCustHelp.Enabled = True
                    cmdInvoiceHelp.Enabled = True
                    Me.txtInvoiceNo.Enabled = False
                    txtCustomerCode.Enabled = False
                    txtASN.Enabled = True
                    cmdInvoiceHelp.Enabled = True
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE  'SAVE BUTTON Pressed
                    SAVEDATA()
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL  'CANCEL BUTTON Pressed
                    Call frmMFGMST0015_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE  'CLOSE BUTTON Pressed
                    Me.Close() 'Close the form
            End Select
            Exit Sub 'To Prevent the execution of the following lines if no errors occurred

        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
        Exit Sub
    End Sub
    Private Sub ctlItem_c1_Change(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles txtInvoiceNo.Change
        On Error GoTo ErrHandler
        If CDbl(Trim(CStr(Len(Me.txtInvoiceNo.Text)))) = 0 Then
            If Me.cmdASN.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                Me.txtASN.Text = ""
                Me.cmdASN.Revert()
                Me.cmdASN.Enabled(1) = False
                Me.cmdASN.Enabled(2) = False
                Me.cmdASN.Enabled(5) = False 'Disable Print Butto
            End If
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub ctlItem_c1_KeyPress(ByVal Sender As System.Object, ByVal e As CtlGeneral.KeyPressEventArgs) Handles txtInvoiceNo.KeyPress
        Dim KeyAscii As Short = e.KeyAscii
        On Error GoTo ErrHandler
        If KeyAscii = System.Windows.Forms.Keys.Space Then
            Beep()
            KeyAscii = 0
        End If
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub ctlItem_c1_KeyUp(ByVal Sender As System.Object, ByVal e As CtlGeneral.KeyUpEventArgs) Handles txtInvoiceNo.KeyUp
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.Shift
        On Error GoTo ErrHandler 'If error occurred, then go to the Error Handler
        ' if F1 key is pressed
        If Me.cmdHelp(0).Enabled Then 'If the Help Button is Enabled, Display Listing
            If KeyCode = 112 Then
                'Call cmdHelp_Click(cmdHelp.Item(0), New System.EventArgs()) 'Help should be invoked if F1 key is pressed
                Me.cmdASN.Focus()
            End If
        End If
        Exit Sub 'To Prevent the execution of the following lines if no errors occurred
ErrHandler:  'The Error Handling Code Starts here
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen) 'Change the Mouse Pointer of the Screen
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Sub
    End Sub
    Private Sub ctlItem_c1_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtInvoiceNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Select Case Me.cmdASN.Mode 'To Check the Mode in which the user is working
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT  'MODE FOR VIEWING DATA
                If Len(Me.txtInvoiceNo.Text) > 0 Then 'Checking if Item Field is not Blank
                    If Not txtInvoiceNo.ExistsRec Then 'Checking if the Record Exists
                        Call RefreshFrm() 'If Invalid Item, Refresh Form
                        Me.cmdASN.Enabled(1) = False
                        Me.cmdASN.Enabled(2) = False
                        Me.txtInvoiceNo.Text = ""
                        txtInvoiceNo.Focus() 'Set the Focus back to the Item Field
                        Cancel = True
                        GoTo EventExitSub
                    Else 
                        If Me.cmdASN.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then 'If user is in VIEW MODE
                            Me.cmdASN.Enabled(1) = True 'Enable Edit Button
                            Me.cmdASN.Enabled(2) = True 'Enable Delete Button
                        Else 'If user is not in VIEW MODE
                            Me.cmdASN.Enabled(1) = True 'Enable Only Edit Button
                        End If
                    End If
                End If
        End Select
        GoTo EventExitSub 'To Prevent the execution of the following lines if no errors occurred
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmMFGMST0015_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        Me.txtInvoiceNo.Focus() 'To Set the Focus
        mdifrmMain.CheckFormName = mintIndex
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default) 'To change the mouse pointer
        'Me.txtInvoiceNo.ToolTip = "Enter Packing Style Code Maximum of 4 characters long."
        'Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMFGMST0015_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMFGMST0015_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlFormHeader1_Click(ctlFormHeader1, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Public Sub frmMFGMST0015_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = System.Windows.Forms.Keys.Escape Then
            'If user press the ESC Key ,the Form will be unloaded
            If Me.cmdASN.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then 'If the User is in the VIEW MODE
                If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then 'Get confirmation before cancelling the current operation
                    Call RefreshFrm() 'Clear the Form

                    'Me.cmdHelp(0).Enabled = True 'Enable the Help Button
                    Me.cmdASN.Enabled(1) = False
                    Me.cmdASN.Enabled(2) = False 'Disable Edit and Delete Button
                    Me.cmdASN.Enabled(5) = False 'Disable Print Button
                    'Me.txtInvoiceNo.Text = Me.txtInvoiceNo.Tag 'To revert back to the old value
                    'Call ctlItem_c1_Validating(txtInvoiceNo, New System.ComponentModel.CancelEventArgs(True)) 'To Display details as per the code displayed
                Else
                    Me.txtInvoiceNo.Focus()
                    GoTo EventExitSub
                End If
            End If
        ElseIf KeyAscii = System.Windows.Forms.Keys.Return Then  'Enter key pressed
            System.Windows.Forms.SendKeys.Send("{TAB}") 'The control is forwarded to next control
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub frmMFGMST0015_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Try
            Call Initialize_controls() 'Function called to Initialize the controls on the Form
            cmdASN.ShowButtons(True, False, False, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub
    Private Sub Initialize_controls()
        Try
            Call FillLabelFromResFile(Me) 'To Fill label description from Resource file
            Call FitToClient(Me, Frame1, ctlFormHeader1, cmdASN) 'To fit the form in the MDI
            Me.txtASN.Enabled = False
            mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag) 'To add the form to the Window List Menu
            Exit Sub
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub
    Private Sub RefreshFrm()
        Try
            Call Me.cmdASN.Revert()
            Me.txtInvoiceNo.Enabled = False

            'Me.cmdHelp(0).Enabled = False
            cmdCustHelp.Enabled = False
            cmdInvoiceHelp.Enabled = False
            txtCustomerCode.Enabled = False
            txtCustomerCode.Text = ""
            txtInvoiceNo.Text = ""
            txtASN.Text = ""
            cmdCustHelp.Focus()
            Exit Sub
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub
    Private Function SAVEDATA() As Boolean
        Try
            If txtCustomerCode.Text.Trim = "" Then
                MsgBox("Kindly select customer Code First !", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return False
                Exit Function
            ElseIf txtInvoiceNo.Text.Trim = "" Then
                MsgBox("Kindly select Invoice No !", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return False
                Exit Function
            End If
            Using sqlCmd As SqlCommand = New SqlCommand
                With sqlCmd
                    .CommandText = "USP_SAVE_MANUAL_ASN"
                    .CommandTimeout = 0
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@CUSTOMERCODE", SqlDbType.VarChar, 8).Value = txtCustomerCode.Text.Trim
                    .Parameters.Add("@INVOICENO", SqlDbType.Float).Value = txtInvoiceNo.Text
                    .Parameters.Add("@ASN_ACKNNO", SqlDbType.VarChar, 100).Value = txtASN.Text
                    .Parameters.Add("@USER_ID", SqlDbType.VarChar, 20).Value = mP_User
                    SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                    MsgBox("ASN Updated successfully", MsgBoxStyle.Information, ResolveResString(100))
                End With
            End Using
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try

    End Function
    Private Sub frmMFGMST0015_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            frmModules.NodeFontBold(Me.Tag) = False
            mdifrmMain.RemoveFormNameFromWindowList = mintIndex 'To Remove the FORM from the Windows Menu
            Me.Dispose()
            Exit Sub
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Private Sub AddRecord()
      
    End Sub
   
    
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm") '("HLPCSTMS0001.htm")
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub

    Private Sub txtCustomerCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyUp
        Try
            If e.KeyCode = Keys.F1 Then
                cmdCustHelp.PerformClick()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

   
    Private Sub txtCustomerCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.TextChanged
        Try
            LblCustomerName.Text = ""
            txtInvoiceNo.Text = ""
            txtASN.Text = ""
            Exit Sub
        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub txtCustomerCode_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
       
    End Sub

    Private Sub cmdCustHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCustHelp.Click
        Dim strCustomer() As String
        Dim strQry As String = ""
        Try
            strQry = "select customer_code As [Customer],cust_name as [Customer Name] from VW_MANUAL_ASN_CUST_HELP  where unit_code='" & gstrUNITID & "'"
            strCustomer = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
            If IsNothing(strCustomer) = True Then Exit Sub
            If strCustomer.GetUpperBound(0) <> -1 Then
                If (Len(strCustomer(0)) >= 1) And strCustomer(0) = "0" Then
                    MsgBox("No Record found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                    Exit Sub
                Else
                    txtCustomerCode.Text = strCustomer(0)
                    LblCustomerName.Text = strCustomer(1)
                   
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub cmdInvoiceHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdInvoiceHelp.Click
        Dim strCustomer() As String
        Dim strQry As String = ""
        Try
            If txtCustomerCode.Text.Trim = "" Then
                MsgBox("Kindly select Customer Code First !", MsgBoxStyle.Exclamation, ResolveResString(100))
                Exit Sub
            End If
            strQry = "select * from VW_MANUAL_ASN_INVOICE_HELP  where unit_code='" & gstrUNITID & "' and customer_code='" & txtCustomerCode.Text & "' "
            strCustomer = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
            If IsNothing(strCustomer) = True Then Exit Sub
            If strCustomer.GetUpperBound(0) <> -1 Then
                If (Len(strCustomer(0)) >= 1) And strCustomer(0) = "0" Then
                    MsgBox("No Record found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                    Exit Sub
                Else
                    txtInvoiceNo.Text = strCustomer(0)
                End If
                txtASN.Text = ""
            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub
End Class