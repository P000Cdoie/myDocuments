Option Strict Off
Option Explicit On

Imports System
Imports System.Data
Imports System.Data.SqlClient

Imports VB = Microsoft.VisualBasic

Friend Class frmMKTTRN0087
    Inherits System.Windows.Forms.Form

#Region "Comments"
    '***************************************************************************************
    'COPYRIGHT(C)   : MOTHERSON SUMI INFOTECH & DESIGN LTD. 
    'FORM NAME      : FRMMKTTRN0087 - ARE-3 Acknowledgement
    'CREATED BY     : Abhinav Kumar 
    'CREATED DATE   : 22 Jan 2015
    'Issue ID       : 10736222
    '***************************************************************************************
#End Region


#Region "Form level variable Declarations"
    Dim mintFormtag As Short
    Dim mctlError As System.Windows.Forms.Control
#End Region

#Region "Form Events"
    Private Sub frmMKTTRN0087_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        Try
            mdifrmMain.CheckFormName = mintFormtag
            frmModules.NodeFontBold(Me.Tag) = True
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub frmMKTTRN0087_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate

        Try
            frmModules.NodeFontBold(Me.Tag) = False
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub frmMKTTRN0087_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Try
            Call FitToClient(Me, fraBase, ctlFormHeader1, GrpButtons, 550)
            Call EnableControls(False, Me, True)
            MdiParent = prjMPower.mdifrmMain
            gblnCancelUnload = False : gblnFormAddEdit = False
            mintFormtag = mdifrmMain.AddFormNameToWindowList(Me.ctlFormHeader1.Tag)
            fraBase.Enabled = True

            DTPAcknowledgeDate.Format = DateTimePickerFormat.Custom
            DTPAcknowledgeDate.CustomFormat = gstrDateFormat
            DTPAcknowledgeDate.Value = GetServerDateNew()
            DTPAcknowledgeDate.MaxDate = GetServerDateNew()

            TxtARENo.Enabled = True
            CmdHlpARE3.Enabled = True

            GrpButtons.Revert()
            ctlFormHeader1.HeaderString = "ARE-3 Acknowledgement"
            Me.Text = "MKTTRN0087 - ARE-3 Acknowledgement"
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub frmMKTTRN0087_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000

        Try
            If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
                Call ctlFormHeader1_Click(ctlFormHeader1, New System.EventArgs())
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub frmMKTTRN0087_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Try
            Select Case KeyAscii
                Case System.Windows.Forms.Keys.Escape
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        RefreshForm()
                    Else
                        Me.ActiveControl.Focus()
                    End If
            End Select
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            e.KeyChar = Chr(KeyAscii)
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        End Try

    End Sub

    Private Sub frmMKTTRN0087_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Try
            mdifrmMain.RemoveFormNameFromWindowList = mintFormtag
            frmModules.NodeFontBold(Me.Tag) = False
            Me.Dispose()
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub
#End Region

#Region "Other Controls Events"
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        Try
            Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub TxtCT2No_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCT2No.KeyPress
        e.Handled = True
    End Sub

    Private Sub CmdARE3Help_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdHlpARE3.Click

        Dim strVal() As String
        Dim strSQL As String = String.Empty

        Try
            strSQL = "Select A.ARE_No,A.INV_NO, A.INV_DATE,B.Cust_Name,A.Cust_Code, A.CT2_NO from ARE3_MST A " & _
                " inner join Customer_mst B on (A.UNIT_CODE=B.UNIT_CODE and A.CUST_CODE=B.Customer_Code)" & _
                " where A.ACK_RECEIVED = 0 and A.CANCELLED = 0 and A.Unit_Code = '" & gstrUNITID & "'"

            strVal = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, ResolveResString(100))

            If Not strVal Is Nothing Then
                If UBound(strVal) > 0 And strVal(0) <> "0" Then
                    TxtARENo.Text = strVal(0)
                    TxtInvoiceNo.Text = strVal(1)
                    TxtInvoiceDate.Text = strVal(2)
                    TxtCustomerName.Text = strVal(3)
                    TxtCustCode.Text = strVal(4)
                    TxtCT2No.Text = strVal(4)
                Else
                    MessageBox.Show("No ARENo pending for acknowledgement", ResolveResString(100), MessageBoxButtons.OK)
                    RefreshForm()
                End If
            Else
                MessageBox.Show("No ARENo Selected.", ResolveResString(100), MessageBoxButtons.OK)
                RefreshForm()
            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub
#End Region

#Region "Routines"
    Private Sub RefreshForm()

        Try
            GrpButtons.Revert()
            fraBase.Enabled = True
            Call EnableControls(False, Me, True)
            CmdHlpARE3.Enabled = True

            DTPAcknowledgeDate.Format = DateTimePickerFormat.Custom
            DTPAcknowledgeDate.CustomFormat = gstrDateFormat
            DTPAcknowledgeDate.Value = GetServerDateNew()
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub
#End Region

    Private Sub GrpButtons_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtnEditGrp.ButtonClickEventArgs) Handles GrpButtons.ButtonClick

        Dim strSQL As String = String.Empty
        Try
            Select Case e.Button
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                    If TxtARENo.Text.Trim.Length > 0 Then
                        CmdHlpARE3.Enabled = False
                        ChkARE3.Checked = True
                        DTPAcknowledgeDate.Enabled = True
                    Else
                        MessageBox.Show("Select atleast one ARENo", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        CmdHlpARE3.Focus()
                        GrpButtons.Revert()
                        Exit Sub
                    End If

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                    Select Case GrpButtons.Mode
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                            If ValidRecord() = False Then Exit Sub
                            strSQL = "Update ARE3_MST set ACK_RECEIVED = 1, ACK_DATE = '" & getDateForDB(DTPAcknowledgeDate.Value) & "'"
                            strSQL += " where unit_Code = '" & gstrUNITID & "' and ARE_NO = '" & TxtARENo.Text.Trim & "' and INV_No = '" & TxtInvoiceNo.Text.Trim & "'"
                            SqlConnectionclass.ExecuteNonQuery(strSQL)
                            MessageBox.Show("ARE3 No is acknowledged.", ResolveResString(100), MessageBoxButtons.OK)
                            RefreshForm()
                    End Select
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                    Call frmMKTTRN0087_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                    Me.Close()
            End Select
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            gblnCancelUnload = True : gblnFormAddEdit = True
        End Try

    End Sub
    Public Function ValidRecord() As Boolean
        Dim strSql As String = String.Empty
        ValidRecord = False

        If DTPAcknowledgeDate.Value > GetServerDate() Then
            MessageBox.Show("ARE3 Ack Received Date cannot be greater than Todays Date!", ResolveResString(100), MessageBoxButtons.OK)
            DTPAcknowledgeDate.Focus()
            Exit Function
        End If
        If DTPAcknowledgeDate.Value < TxtInvoiceDate.Text.Trim Then
            MessageBox.Show("ARE3 Ack Received Date cannot be less than Invoice Date!", ResolveResString(100), MessageBoxButtons.OK)
            DTPAcknowledgeDate.Focus()
            Exit Function
        End If

        ValidRecord = True

    End Function
End Class