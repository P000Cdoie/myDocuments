Option Strict Off
Option Explicit On
Imports System.Data.SqlClient
Imports System.IO

'-----------------------------------------------------------------------------
'Copyright (c)  -  MIND
'Name of module -  frmMKTTRN0109.vb
'Created By     -  Shubhra Verma   
'Created Date   -  Jun 2019
'Description    -  View / Print Signed PDFs
'-----------------------------------------------------------------------------

Friend Class frmMKTTRN0109
    Inherits System.Windows.Forms.Form

    Dim strUploadedFileName As String
    Dim mintIndex As Short
    Dim mflag As Short

    Private Sub cmdclose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdclose.Click
        Try

            Me.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub frmMKTTRN0109_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Try

            mdifrmMain.CheckFormName = mintIndex
            frmModules.NodeFontBold(Tag) = True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub frmMKTTRN0109_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        Try

            frmModules.NodeFontBold(Me.Tag) = False

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub frmMKTTRN0109_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Try

            mintIndex = mdifrmMain.AddFormNameToWindowList(ctlDigitallySigned.Text)
            FitToClient(Me, framain, ctlDigitallySigned, frabuttons)
            Call FillLabelFromResFile(Me)
            Call getDocTypeText()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub frmMKTTRN0109_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            frmModules.NodeFontBold(Me.Tag) = False
            mdifrmMain.RemoveFormNameFromWindowList = mintIndex

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub txtcustomerhelp_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtcustomerhelp.KeyDown

        Dim strSQL As String = String.Empty
        Dim strHelp() As String

        Try

            txtInvFrom.Text = "" : txtInvTo.Text = ""
            grdInvData.DataSource = Nothing

            strSQL = "SELECT ACCOUNT_CODE ,CUST_NAME FROM VW_SignedPDFs_Customer WHERE UNIT_CODE ='" & gstrUNITID & "'"
            strHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL)

            If IsNothing(strHelp) = True Then Exit Sub
            If strHelp.GetUpperBound(0) <> -1 Then
                If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                    MsgBox("No Record found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                    Exit Sub
                Else
                    txtcustomerhelp.Text = strHelp(0)
                    lblcustname.Text = strHelp(1)
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub txtcustomerhelp_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtcustomerhelp.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Try
            If KeyAscii = 13 Then
                Call txtcustomerhelp_Validating(txtcustomerhelp, New System.ComponentModel.CancelEventArgs(False))
                If Len(Trim(txtcustomerhelp.Text)) <> 0 Then
                    System.Windows.Forms.SendKeys.Send("{tab}")
                End If
            End If
            If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = System.Windows.Forms.Keys.Back Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 22 Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
                KeyAscii = 0
            ElseIf KeyAscii = System.Windows.Forms.Keys.Return Then
                If Len(Trim(txtcustomerhelp.Text)) <> 0 Then
                    Call txtcustomerhelp_Validating(txtcustomerhelp, New System.ComponentModel.CancelEventArgs(False))
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            eventArgs.KeyChar = Chr(KeyAscii)
            If KeyAscii = 0 Then
                eventArgs.Handled = True
            End If
        End Try
    End Sub

    Private Sub txtcustomerhelp_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtcustomerhelp.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Try
            '  Dim oRs As ADODB.Recordset
            Dim strSQL As String
            Dim oDr As SqlDataReader

            If Len(txtcustomerhelp.Text) > 0 Then
                txtcustomerhelp.Text = Replace(txtcustomerhelp.Text, "'", "")
                strSQL = "select a.customer_Code, a.cust_name from Customer_Mst a where a.customer_code='" & Trim(txtcustomerhelp.Text) & "' and a.Unit_code = '" & gstrUNITID & "' and ((isnull(a.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= a.deactive_date))"
                oDr = SqlConnectionclass.ExecuteReader(strSQL)

                If oDr.HasRows Then
                    oDr.Read()
                    lblcustname.Text = oDr("cust_name").ToString()
                Else
                    MsgBox("Invalid Customer Code", MsgBoxStyle.Information, ResolveResString(100))
                    lblcustname.Text = ""
                    txtcustomerhelp.Text = ""
                    txtcustomerhelp.Enabled = True
                    txtcustomerhelp.Focus()
                End If
            Else
                lblcustname.Text = ""
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            eventArgs.Cancel = Cancel
        End Try
    End Sub

    Private Sub getDocTypeText()
        Try
            Dim strQry As String = ""
            Dim SQLAdp As New SqlDataAdapter
            Dim DsDistinctRows As New DataSet

            strQry = "select DOC_TYPE_TEXT from dbo.ufn_getSignedInvoiceTypeText ('" + gstrUNITID + "','DOC_TYPE')"
            SQLAdp = New SqlDataAdapter(strQry, SqlConnectionclass.GetConnection)
            SQLAdp.Fill(DsDistinctRows)
            cmbInvoiceType.DataSource = DsDistinctRows.Tables(0)
            cmbInvoiceType.ValueMember = "DOC_TYPE_TEXT"
            cmbInvoiceType.DisplayMember = "DOC_TYPE_TEXT"
            cmbInvoiceType.SelectedIndex = -1

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub getInvoice()

        Try

            Dim strQry As String = ""
            Dim SQLAdp As New SqlDataAdapter
            Dim DsDistinctRows As New DataSet

            strQry = "select * from dbo.ufn_getSignedInvoiceTypeText ('" + gstrUNITID + "','InvoiceNo') where doc_type_text = '" + cmbInvoiceType.SelectedItem + "' )"
            SQLAdp = New SqlDataAdapter(strQry, SqlConnectionclass.GetConnection)
            SQLAdp.Fill(DsDistinctRows)

            cmbInvoiceType.DataSource = DsDistinctRows.Tables(0)
            cmbInvoiceType.ValueMember = "DOC_TYPE_TEXT"
            cmbInvoiceType.DisplayMember = "DOC_TYPE_TEXT"
            cmbInvoiceType.SelectedIndex = -1

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub DownloadPDFs(ByVal strDocNo As String)
        Try
            Dim strFilePath, strQry As String
            Dim buffer As Byte()

            strQry = "Select top 1 PDF_BINARY from SALESCHALLAN_SIGNED_PDFS  where UNIT_CODE = '" + gstrUNITID + "' and DOC_NO = '" + strDocNo + "' and doc_type_text = '" + cmbInvoiceType.SelectedValue + "'"
            buffer = SqlConnectionclass.ExecuteScalar(strQry)
            ''Praveen Digital Sign Changes 
            If cmbInvoiceType.SelectedValue = "ORIGINAL FOR BUYER" Then
                strFilePath = System.IO.Path.GetTempPath() + strDocNo + "_org.pdf"
            ElseIf cmbInvoiceType.SelectedValue = "DUPLICATE FOR TRANSPORTER" Then
                strFilePath = System.IO.Path.GetTempPath() + strDocNo + "_dup.pdf"
            ElseIf cmbInvoiceType.SelectedValue = "TRIPLICATE FOR ASSESSEE" Then   ''Praveen Digital Sign Changes 
                strFilePath = System.IO.Path.GetTempPath() + strDocNo + "_tri.pdf"
            ElseIf cmbInvoiceType.SelectedValue = "EXTRA COPY" Then            ''Praveen Digital Sign Changes 
                strFilePath = System.IO.Path.GetTempPath() + strDocNo + "_ext.pdf"
            Else
                strFilePath = System.IO.Path.GetTempPath() + strDocNo + ".pdf"
            End If

            Try
                If System.IO.File.Exists(strFilePath) Then
                    System.IO.File.Delete(strFilePath)
                End If
            Catch ex As Exception

            End Try
            System.IO.File.WriteAllBytes(strFilePath, buffer)

            Dim act As Action(Of String) = New Action(Of String)(AddressOf openPDFFile)
            act.BeginInvoke(strFilePath, Nothing, Nothing)

            Try
                System.IO.File.Delete(System.IO.Path.GetTempPath())
            Catch ex As Exception

            End Try
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Shared Sub openPDFFile(ByVal strFilePath As String)
        Try
            Using p As New System.Diagnostics.Process
                p.StartInfo = New System.Diagnostics.ProcessStartInfo(strFilePath)
                p.Start()
                p.WaitForExit()
                Try
                    System.IO.File.Delete(strFilePath)
                Catch ex As Exception
                    '   MessageBox.Show("error in openPDFFile function", ResolveResString(100), MessageBoxButtons.OK)
                End Try
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub cmdviewInvoice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdviewInvoice.Click
        Try

            getSignedInvoices()

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub getSignedInvoices()
        Try
            Dim strQry As String
            Dim SQLAdp As SqlDataAdapter
            Dim DsDistinctRows As New DataSet

            If txtcustomerhelp.Text = "" Then
                grdInvData.DataSource = Nothing
                MessageBox.Show("Select Customer.", ResolveResString(100), MessageBoxButtons.OK)
                Exit Sub
            End If

            If cmbInvoiceType.SelectedIndex = -1 Then
                grdInvData.DataSource = Nothing
                MessageBox.Show("Select Document Type.", ResolveResString(100), MessageBoxButtons.OK)
                Exit Sub
            End If

            strQry = "select * from dbo.ufn_getSignedInvoices('" + gstrUNITID + "','" + txtcustomerhelp.Text + "',convert(date,'" + dtfromdate.Value + "',103),convert(date,'" + dttodate.Value + "',103),'" + cmbInvoiceType.SelectedValue + "','" + txtInvFrom.Text + "','" + txtInvTo.Text + "')"

            SQLAdp = New SqlDataAdapter(strQry, SqlConnectionclass.GetConnection)
            SQLAdp.Fill(DsDistinctRows)
            grdInvData.DataSource = DsDistinctRows.Tables(0)
            grdInvData.AllowUserToResizeRows = False

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub grdInvData_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grdInvData.CellContentClick
        Try
            Dim row As Integer
            Dim strDocNo As String

            row = e.RowIndex

            If row = -1 Then Exit Sub

            If row = grdInvData.RowCount - 1 Then
                Exit Sub
            End If

            strDocNo = grdInvData.Item(2, e.RowIndex).Value

            DownloadPDFs(strDocNo)

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click
        Try

            txtcustomerhelp.Text = ""
            cmbInvoiceType.SelectedIndex = -1
            lblcustname.Text = ""
            dtfromdate.Value = GetServerDate()
            dttodate.Value = GetServerDate()
            txtInvFrom.Text = ""
            txtInvTo.Text = ""
            grdInvData.DataSource = Nothing

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub cmbInvoiceType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbInvoiceType.SelectedIndexChanged
        Try
            grdInvData.DataSource = Nothing
            If txtcustomerhelp.Text <> "" And cmbInvoiceType.SelectedIndex <> -1 Then
                getSignedInvoices()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub txtInvFrom_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtInvFrom.KeyDown
        Dim strSQL As String = String.Empty
        Dim strHelp() As String

        Try
            If e.KeyCode = System.Windows.Forms.Keys.F1 Then
                If cmbInvoiceType.SelectedIndex = -1 Then
                    MessageBox.Show("Select Document Type", ResolveResString(100), MessageBoxButtons.OK)
                    cmbInvoiceType.Focus()
                Else
                    strSQL = "select InvoiceNo, InvDate from dbo.ufn_getSignedInvoiceTypeText ('" + gstrUNITID + "','InvoiceNo')" & _
                        " where doc_type_text = '" + cmbInvoiceType.SelectedValue + "'" & _
                        " and InvDate between '" + getDateForDB(dtfromdate.Value) + "' and '" + getDateForDB(dttodate.Value) + "' "

                    strHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL)

                    If IsNothing(strHelp) = True Then Exit Sub
                    If strHelp.GetUpperBound(0) <> -1 Then
                        If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                            MsgBox("No Record found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                            Exit Sub
                        Else
                            txtInvFrom.Text = strHelp(0)
                            txtInvTo.Text = strHelp(0)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub txtInvTo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtInvTo.KeyDown
        Dim strSQL As String = String.Empty
        Dim strHelp() As String

        Try
            If e.KeyCode = System.Windows.Forms.Keys.F1 Then
                If cmbInvoiceType.SelectedIndex = -1 Then
                    MessageBox.Show("Select Document Type", ResolveResString(100), MessageBoxButtons.OK)
                    cmbInvoiceType.Focus()
                Else
                    strSQL = "select InvoiceNo, InvDate from dbo.ufn_getSignedInvoiceTypeText ('" + gstrUNITID + "','InvoiceNo')" & _
                        " where doc_type_text = '" + cmbInvoiceType.SelectedValue + "'" & _
                        " and InvDate between '" + getDateForDB(dtfromdate.Value) + "' and '" + getDateForDB(dttodate.Value) + "' "

                    strHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL)

                    If IsNothing(strHelp) = True Then Exit Sub
                    If strHelp.GetUpperBound(0) <> -1 Then
                        If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                            MsgBox("No Record found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                            Exit Sub
                        Else
                            txtInvTo.Text = strHelp(0)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub dtfromdate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtfromdate.ValueChanged, dttodate.ValueChanged
        Try

            txtInvFrom.Text = ""
            txtInvTo.Text = ""
            grdInvData.DataSource = Nothing

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

#Region "Bulk Download Pdf"
    'Modify By Anupam Kumar On the Date of 26 Sep 23.
    Dim FolderPath As String = String.Empty
    Dim dset As New DataSet

    Private Function getSignedInvoicesBulk() As Boolean
        Try
            Dim ValidField As Boolean = False
            Dim strQry As String
            Dim SQLAdp As SqlDataAdapter

            If txtcustomerhelp.Text = "" Then
                MessageBox.Show("Select Customer.", ResolveResString(100), MessageBoxButtons.OK)
                Return False
            End If

            If cmbInvoiceType.SelectedIndex = -1 Then
                MessageBox.Show("Select Document Type.", ResolveResString(100), MessageBoxButtons.OK)
                Return False
            End If

            dset = New DataSet

            strQry = "select * from dbo.ufn_getSignedInvoices('" + gstrUNITID + "','" + txtcustomerhelp.Text + "',convert(date,'" + dtfromdate.Value + "',103),convert(date,'" + dttodate.Value + "',103),'" + cmbInvoiceType.SelectedValue + "','" + txtInvFrom.Text + "','" + txtInvTo.Text + "');"
            strQry += "SELECT top 1 ISNULL(IS_FILE_NAME_AS_SMAC_FORMAT,0) AS IS_FILE_NAME_AS_SMAC_FORMAT FROM CommonDigital_EINVOICING_CONFIG (NOLOCK) WHERE IS_ACTIVE=1 AND UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtcustomerhelp.Text & "';"
            SQLAdp = New SqlDataAdapter(strQry, SqlConnectionclass.GetConnection)
            SQLAdp.Fill(dset)

            If dset IsNot Nothing AndAlso dset.Tables(0).Rows.Count > 0 Then
                ValidField = True
            Else
                MessageBox.Show("Record not found.", ResolveResString(100), MessageBoxButtons.OK)
            End If

            Return ValidField
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Function

    Private Sub DownloadBulkPDFs(ByVal strFilePath As String, ByVal strDocNo As String, ByVal InvoiceDate As String, Optional ByVal IS_FILE_NAME_AS_SMAC_FORMATA As Boolean = False)
        Try
            Dim strQry As String = String.Empty
            Dim FileType As String = String.Empty
            Dim buffer As Byte()

            strQry = "Select top 1 PDF_BINARY from SALESCHALLAN_SIGNED_PDFS  where UNIT_CODE = '" + gstrUNITID + "' and DOC_NO = '" + strDocNo + "' and doc_type_text = '" + cmbInvoiceType.SelectedValue + "'"
            buffer = SqlConnectionclass.ExecuteScalar(strQry)
            'If cmbInvoiceType.SelectedValue = "ORIGINAL FOR BUYER" Then
            '    FileType = "_org.pdf"
            'ElseIf cmbInvoiceType.SelectedValue = "DUPLICATE FOR TRANSPORTER" Then
            '    FileType = "_dup.pdf"
            'ElseIf cmbInvoiceType.SelectedValue = "TRIPLICATE FOR ASSESSEE" Then   ''Praveen Digital Sign Changes 
            '    FileType = "_tri.pdf"
            'ElseIf cmbInvoiceType.SelectedValue = "EXTRA COPY" Then            ''Praveen Digital Sign Changes 
            '    FileType = "_ext.pdf"
            'Else
            '    FileType = ".pdf"
            'End If
            FileType = ".pdf"
            If IS_FILE_NAME_AS_SMAC_FORMATA Then
                strFilePath = Path.Combine(strFilePath, strDocNo & "_" & InvoiceDate & "_" & txtcustomerhelp.Text & FileType)
            Else
                strFilePath = Path.Combine(strFilePath, strDocNo & FileType)
            End If


            Try
                If System.IO.File.Exists(strFilePath) Then
                    System.IO.File.Delete(strFilePath)
                End If
                System.IO.File.WriteAllBytes(strFilePath, buffer)
            Catch ex As Exception
                Return
            End Try


            'Dim act As Action(Of String) = New Action(Of String)(AddressOf openPDFFile)
            'act.BeginInvoke(strFilePath, Nothing, Nothing)


        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub


    Private Sub Btn_DownloadAllInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_DownloadAllInvoice.Click
        Dim IsRecordExists As Boolean = False
        Dim IS_FILE_NAME_AS_SMAC_FORMATA As Boolean = False
        If getSignedInvoicesBulk() Then
            If (SelectFolderBrowserDialog.ShowDialog() = DialogResult.OK) Then
                FolderPath = SelectFolderBrowserDialog.SelectedPath
                If (dset.Tables(1).Rows.Count > 0) Then
                    IS_FILE_NAME_AS_SMAC_FORMATA = IIf(dset.Tables(1).Rows(0)(0).ToString().Equals("1"), True, False)
                End If
                For Each rw As DataRow In dset.Tables(0).Rows
                    Dim InvoiceNo = rw.Item(1).ToString()
                    Dim InvoiceDate = rw.Item(2).ToString().Substring(0, 10).Replace("/", "")
                    DownloadBulkPDFs(FolderPath, InvoiceNo, InvoiceDate, IS_FILE_NAME_AS_SMAC_FORMATA)
                    IsRecordExists = True
                Next
            End If
        End If
        If IsRecordExists Then
            MessageBox.Show("Invoice Download Path:- " & FolderPath, "Invoice Downloaded", MessageBoxButtons.OK)
        End If
    End Sub

#End Region

End Class