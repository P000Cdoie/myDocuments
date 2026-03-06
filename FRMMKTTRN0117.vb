Imports System.IO
Imports System.Data.SqlClient

'*********************************************************************************************************************
'Copyright(c)       - MIND
'Name of Module     - Mahindra EDI/View/Print Digitally Signed Invoices
'Name of Form       - FRMMKTTRN0117  , Mahindra EDI/View/Print Digitally Signed Invoices
'Created by         - Ashish sharma
'Created Date       - 08 SEP 2021
'description        - Mahindra EDI/View/Print Digitally Signed Invoices (New Development) To send invoices to Mahindra by EDI and View / Print digitally signed invoices
'*********************************************************************************************************************
Public Class FRMMKTTRN0117
    Private Const OriginalForBuyer As String = "ORIGINAL FOR BUYER"
    Private Const DuplicateForTransporter As String = "DUPLICATE FOR TRANSPORTER"
    Private MahindraEDIPathForDigitalSignedInvoicePDF As String = String.Empty
    Private Enum GridInvoiceDetails
        Selection = 0
        View
        DocNo
        DocDate
        DocType
        DocStatus
        AccountCode
        AccountName
        InvoiceType
        ASNNo
    End Enum
    Private Sub FRMMKTTRN0117_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            dtpDateFrom.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FRMMKTTRN0117_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Call FitToClient(Me, GrpMain, ctlHeader, GrpBoxButtons, 500)
            Me.MdiParent = mdifrmMain
            Dim currentDate As Date = GetServerDate()
            dtpDateFrom.Value = DateAdd(DateInterval.Day, -1, currentDate)
            dtpDateTo.Value = currentDate
            MahindraEDIPathForDigitalSignedInvoicePDF = Convert.ToString(SqlConnectionclass.ExecuteScalar("SELECT TOP 1 [PATH] FROM MAHINDRA_ASN_CONFIGURATION (NOLOCK) WHERE UNIT_CODE='" & gstrUNITID & "' AND [TYPE]='PDF'"))
            FillDocumentType()
            ConfigureGridColumnns()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub FillDocumentType()
        Dim dtDocumentType As New DataTable
        Try
            Dim strQry As String = "SELECT DISTINCT UPPER(DOC_TYPE_TEXT) AS DOC_TYPE_TEXT FROM SALESCHALLAN_SIGNED_PDFS (NOLOCK) WHERE UNIT_CODE = '" & gstrUNITID & "' ORDER BY DOC_TYPE_TEXT DESC"
            dtDocumentType = SqlConnectionclass.GetDataTable(strQry)
            If dtDocumentType IsNot Nothing AndAlso dtDocumentType.Rows.Count > 0 Then
                cmbInvoiceType.DataSource = dtDocumentType
                cmbInvoiceType.ValueMember = "DOC_TYPE_TEXT"
                cmbInvoiceType.DisplayMember = "DOC_TYPE_TEXT"
                cmbInvoiceType.SelectedIndex = cmbInvoiceType.FindStringExact(OriginalForBuyer)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub ConfigureGridColumnns()
        Try
            dgvInvoiceDetail.Columns.Clear()
            Dim objChkBox As New DataGridViewCheckBoxColumn
            objChkBox.Name = "Selection"
            objChkBox.HeaderText = " "

            Dim objButton As New DataGridViewButtonColumn
            objButton.Name = "View"
            objButton.HeaderText = " "

            dgvInvoiceDetail.Columns.Add(objChkBox)
            dgvInvoiceDetail.Columns.Add(objButton)
            dgvInvoiceDetail.Columns.Add("DocNo", "Doc. No.")
            dgvInvoiceDetail.Columns.Add("DocDate", "Doc. Date")
            dgvInvoiceDetail.Columns.Add("DocType", "Doc. Type")
            dgvInvoiceDetail.Columns.Add("DocStatus", "Status")
            dgvInvoiceDetail.Columns.Add("AccountCode", "Cust. Code")
            dgvInvoiceDetail.Columns.Add("AccountName", "Cust. Name")
            dgvInvoiceDetail.Columns.Add("InvoiceType", "Inv. Type")
            dgvInvoiceDetail.Columns.Add("ASNNo", "ASN No.")

            dgvInvoiceDetail.Columns(GridInvoiceDetails.Selection).Width = 35
            dgvInvoiceDetail.Columns(GridInvoiceDetails.View).Width = 50
            dgvInvoiceDetail.Columns(GridInvoiceDetails.DocNo).Width = 75
            dgvInvoiceDetail.Columns(GridInvoiceDetails.DocDate).Width = 80
            dgvInvoiceDetail.Columns(GridInvoiceDetails.DocType).Width = 180
            dgvInvoiceDetail.Columns(GridInvoiceDetails.DocStatus).Width = 80
            dgvInvoiceDetail.Columns(GridInvoiceDetails.AccountCode).Width = 90
            dgvInvoiceDetail.Columns(GridInvoiceDetails.AccountName).Width = 200
            dgvInvoiceDetail.Columns(GridInvoiceDetails.InvoiceType).Width = 100
            dgvInvoiceDetail.Columns(GridInvoiceDetails.ASNNo).Width = 100

            dgvInvoiceDetail.Columns(GridInvoiceDetails.Selection).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvInvoiceDetail.Columns(GridInvoiceDetails.View).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvInvoiceDetail.Columns(GridInvoiceDetails.DocNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvInvoiceDetail.Columns(GridInvoiceDetails.DocDate).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvInvoiceDetail.Columns(GridInvoiceDetails.DocType).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvInvoiceDetail.Columns(GridInvoiceDetails.DocStatus).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvInvoiceDetail.Columns(GridInvoiceDetails.AccountCode).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvInvoiceDetail.Columns(GridInvoiceDetails.AccountName).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvInvoiceDetail.Columns(GridInvoiceDetails.InvoiceType).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvInvoiceDetail.Columns(GridInvoiceDetails.ASNNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            dgvInvoiceDetail.Columns(GridInvoiceDetails.DocNo).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetails.DocDate).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetails.DocType).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetails.DocStatus).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetails.AccountCode).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetails.AccountName).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetails.InvoiceType).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetails.ASNNo).ReadOnly = True
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub CmdCustCodeHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdCustCodeHelp.Click
        Dim strHelp() As String = Nothing
        Try
            Dim strSql As String = String.Empty
            If dtpDateFrom.Value > dtpDateTo.Value Then
                MessageBox.Show("[From Date] should be less than or equal to [To Date].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                dtpDateFrom.Focus()
                Exit Sub
            End If
            If cmbInvoiceType.SelectedIndex = -1 Then
                MessageBox.Show("Please select [Doc. Type].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                cmbInvoiceType.Focus()
                Exit Sub
            End If
            strSql = "SELECT DISTINCT S.ACCOUNT_CODE , C.CUST_NAME FROM SALESCHALLAN_SIGNED_PDFS S (NOLOCK) INNER JOIN CUSTOMER_MST C (NOLOCK) ON S.UNIT_CODE=C.UNIT_CODE AND S.ACCOUNT_CODE=C.CUSTOMER_CODE INNER JOIN MAHINDRA_EINVOICING_CONFIG M (NOLOCK) ON S.UNIT_CODE=M.UNIT_CODE AND S.ACCOUNT_CODE=M.CUSTOMER_CODE WHERE S.UNIT_CODE='" & gstrUNITID & "' ORDER BY S.ACCOUNT_CODE"
            strHelp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "Customer Help")
            If Not (UBound(strHelp) <= 0) Then
                If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                    MessageBox.Show("No record To Display", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    txtCustomerCode.Text = String.Empty
                    lblCustCodeDes.Text = String.Empty
                    Exit Sub
                Else
                    txtCustomerCode.Text = strHelp(0).Trim
                    lblCustCodeDes.Text = strHelp(1).Trim
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtCustomerCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Try
            If e.KeyCode = Keys.F1 Then
                CmdCustCodeHelp_Click(sender, e)
            ElseIf e.KeyCode = Keys.Delete Then
                txtCustomerCode.Text = String.Empty
                lblCustCodeDes.Text = String.Empty
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtCustomerCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustomerCode.TextChanged
        Try
            txtInvoiceFrom.Text = String.Empty
            txtInvoiceTo.Text = String.Empty
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub cmbInvoiceType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbInvoiceType.SelectedIndexChanged
        Try
            dgvInvoiceDetail.Rows.Clear()
            If cmbInvoiceType.SelectedIndex = cmbInvoiceType.FindStringExact(OriginalForBuyer) Then
                chkAllPendinginvoicesforEDIMapping.Enabled = True
                btnSendToEDI.Enabled = True
            Else
                chkAllPendinginvoicesforEDIMapping.Checked = False
                chkAllPendinginvoicesforEDIMapping.Enabled = False
                btnSendToEDI.Enabled = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub cmdInvoiceFrom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInvoiceFrom.Click
        Dim strHelp() As String = Nothing
        Try
            Dim strSql As String = String.Empty
            If dtpDateFrom.Value > dtpDateTo.Value Then
                MessageBox.Show("[From Date] should be less than or equal to [To Date].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                dtpDateFrom.Focus()
                Exit Sub
            End If
            If cmbInvoiceType.SelectedIndex = -1 Then
                MessageBox.Show("Please select [Doc. Type].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                cmbInvoiceType.Focus()
                Exit Sub
            End If
            If String.IsNullOrEmpty(txtCustomerCode.Text) Then
                MessageBox.Show("Please select [Customer Code].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtCustomerCode.Focus()
                Exit Sub
            End If
            strSql = "SELECT CAST(S.Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),S.Invoice_Date,103) as Invoice_Date FROM SALESCHALLAN_DTL S (NOLOCK) INNER JOIN SALESCHALLAN_SIGNED_PDFS P (NOLOCK) ON S.UNIT_CODE=P.UNIT_CODE AND S.DOC_NO=P.DOC_NO "
            strSql += " WHERE S.UNIT_CODE='" & gstrUNITID & "' AND S.BILL_FLAG=1 AND S.CANCEL_FLAG=0 AND S.ACCOUNT_CODE='" & Trim(txtCustomerCode.Text) & "' AND P.DOC_TYPE_TEXT='" & Trim(cmbInvoiceType.SelectedValue) & "' "
            strSql += " AND (S.INVOICE_DATE BETWEEN '" & VB6.Format(Me.dtpDateFrom.Value, "dd/MMM/yyyy") & "' AND '" & VB6.Format(Me.dtpDateTo.Value, "dd/MMM/yyyy") & "') ORDER BY S.DOC_NO "
            strHelp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "Invoice No. Help")
            If Not (UBound(strHelp) <= 0) Then
                If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                    MessageBox.Show("No record To Display", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    txtInvoiceFrom.Text = String.Empty
                    Exit Sub
                Else
                    txtInvoiceFrom.Text = strHelp(0).Trim
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtInvoiceFrom_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtInvoiceFrom.KeyDown
        Try
            If e.KeyCode = Keys.F1 Then
                cmdInvoiceFrom_Click(sender, e)
            ElseIf e.KeyCode = Keys.Delete Then
                txtInvoiceFrom.Text = String.Empty
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtInvoiceFrom_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtInvoiceFrom.TextChanged
        Try
            txtInvoiceTo.Text = String.Empty
            dgvInvoiceDetail.Rows.Clear()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub cmdInvoiceTo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInvoiceTo.Click
        Dim strHelp() As String = Nothing
        Try
            Dim strSql As String = String.Empty
            If dtpDateFrom.Value > dtpDateTo.Value Then
                MessageBox.Show("[From Date] should be less than or equal to [To Date].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                dtpDateFrom.Focus()
                Exit Sub
            End If
            If cmbInvoiceType.SelectedIndex = -1 Then
                MessageBox.Show("Please select [Doc. Type].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                cmbInvoiceType.Focus()
                Exit Sub
            End If
            If String.IsNullOrEmpty(txtCustomerCode.Text) Then
                MessageBox.Show("Please select [Customer Code].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtCustomerCode.Focus()
                Exit Sub
            End If
            If String.IsNullOrEmpty(txtInvoiceFrom.Text) Then
                MessageBox.Show("Please select [From Invoice].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtInvoiceFrom.Focus()
                Exit Sub
            End If
            strSql = "SELECT CAST(S.Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),S.Invoice_Date,103) as Invoice_Date FROM SALESCHALLAN_DTL S (NOLOCK) INNER JOIN SALESCHALLAN_SIGNED_PDFS P (NOLOCK) ON S.UNIT_CODE=P.UNIT_CODE AND S.DOC_NO=P.DOC_NO "
            strSql += " WHERE S.UNIT_CODE='" & gstrUNITID & "' AND S.BILL_FLAG=1 AND S.CANCEL_FLAG=0 AND S.ACCOUNT_CODE='" & Trim(txtCustomerCode.Text) & "' AND P.DOC_TYPE_TEXT='" & Trim(cmbInvoiceType.SelectedValue) & "' "
            strSql += " AND (S.INVOICE_DATE BETWEEN '" & VB6.Format(Me.dtpDateFrom.Value, "dd/MMM/yyyy") & "' AND '" & VB6.Format(Me.dtpDateTo.Value, "dd/MMM/yyyy") & "') ORDER BY S.DOC_NO "
            strHelp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "Invoice No. Help")
            If Not (UBound(strHelp) <= 0) Then
                If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                    MessageBox.Show("No record To Display", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    txtInvoiceTo.Text = String.Empty
                    Exit Sub
                Else
                    txtInvoiceTo.Text = strHelp(0).Trim
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtInvoiceTo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtInvoiceTo.KeyDown
        Try
            If e.KeyCode = Keys.F1 Then
                cmdInvoiceTo_Click(sender, e)
            ElseIf e.KeyCode = Keys.Delete Then
                txtInvoiceTo.Text = String.Empty
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtInvoiceTo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtInvoiceTo.TextChanged
        Try
            dgvInvoiceDetail.Rows.Clear()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub CheckUncheckInvoicesAll()
        If dgvInvoiceDetail Is Nothing OrElse dgvInvoiceDetail.Rows.Count = 0 Then
            rdbInvoiceCheckAll.Checked = False
            rdbInvoiceUncheckAll.Checked = False
            Exit Sub
        End If
        With dgvInvoiceDetail
            For i As Integer = 0 To .Rows.Count - 1
                .Rows(i).Cells(GridInvoiceDetails.Selection).Value = rdbInvoiceCheckAll.Checked
            Next
        End With
    End Sub

    Private Sub rdbInvoiceCheckAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbInvoiceCheckAll.CheckedChanged
        Try
            CheckUncheckInvoicesAll()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            dgvInvoiceDetail.Rows.Clear()
            Dim currentDate As Date = GetServerDate()
            dtpDateFrom.Value = DateAdd(DateInterval.Day, -1, currentDate)
            dtpDateTo.Value = currentDate
            cmbInvoiceType.SelectedIndex = cmbInvoiceType.FindStringExact(OriginalForBuyer)
            txtCustomerCode.Text = String.Empty
            lblCustCodeDes.Text = String.Empty
            txtInvoiceFrom.Text = String.Empty
            txtInvoiceTo.Text = String.Empty
            rdbInvoiceCheckAll.Checked = False
            rdbInvoiceUncheckAll.Checked = False
            chkAllPendinginvoicesforEDIMapping.Checked = False
            chkAllPendinginvoicesforEDIMapping.Enabled = True
            dtpDateFrom.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnSendToEDI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSendToEDI.Click
        Try
            If String.IsNullOrEmpty(MahindraEDIPathForDigitalSignedInvoicePDF) Then
                MsgBox("Please define EDI path in [MAHINDRA_ASN_CONFIGURATION].", MsgBoxStyle.Exclamation, ResolveResString(100))
                Exit Sub
            End If
            If dgvInvoiceDetail Is Nothing OrElse dgvInvoiceDetail.Rows.Count = 0 Then
                MsgBox("No invoice found to send to EDI.", MsgBoxStyle.Exclamation, ResolveResString(100))
                Exit Sub
            End If
            Dim flag As Boolean = False

            For i As Integer = 0 To dgvInvoiceDetail.Rows.Count - 1
                If Convert.ToBoolean(dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.Selection).Value) Then
                    flag = True
                    Exit For
                End If
            Next
            If Not flag Then
                MsgBox("Please select atleast one invoice to send to EDI.", MsgBoxStyle.Exclamation, ResolveResString(100))
                dgvInvoiceDetail.Focus()
                dgvInvoiceDetail.CurrentCell = dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.Selection)
                Exit Sub
            End If

            Dim docNo As String = String.Empty
            Dim docDate As String = String.Empty
            Dim docTypeText As String = String.Empty
            Dim asnNo As String = String.Empty
            For i As Integer = 0 To dgvInvoiceDetail.Rows.Count - 1
                If Convert.ToBoolean(dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.Selection).Value) Then
                    docNo = Convert.ToString(dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.DocNo).Value).Trim()
                    docTypeText = Convert.ToString(dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.DocType).Value).Trim()
                    asnNo = Convert.ToString(dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.ASNNo).Value).Trim()
                    If String.IsNullOrEmpty(docNo) Then
                        MsgBox("Doc No. should not be blank.", MsgBoxStyle.Exclamation, ResolveResString(100))
                        dgvInvoiceDetail.Focus()
                        dgvInvoiceDetail.CurrentCell = dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.Selection)
                        Exit Sub
                    End If
                    If String.IsNullOrEmpty(docTypeText) Then
                        MsgBox("Doc. Type should not be blank.", MsgBoxStyle.Exclamation, ResolveResString(100))
                        dgvInvoiceDetail.Focus()
                        dgvInvoiceDetail.CurrentCell = dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.Selection)
                        Exit Sub
                    End If
                    If String.IsNullOrEmpty(asnNo) Then
                        MsgBox("ASN No. should not be blank.", MsgBoxStyle.Exclamation, ResolveResString(100))
                        dgvInvoiceDetail.Focus()
                        dgvInvoiceDetail.CurrentCell = dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.Selection)
                        Exit Sub
                    End If
                End If
            Next
            If MessageBox.Show("Are you Sure To Send to EDI?", "Confirmation", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
            docNo = String.Empty
            docDate = String.Empty
            asnNo = String.Empty
            Dim strFilePath, strQry As String
            Dim buffer As Byte()
            Dim destFilePath As String = String.Empty
            Dim fileName As String = String.Empty
            For i As Integer = 0 To dgvInvoiceDetail.Rows.Count - 1
                destFilePath = MahindraEDIPathForDigitalSignedInvoicePDF
                If Convert.ToBoolean(dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.Selection).Value) Then
                    docNo = Convert.ToString(dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.DocNo).Value)
                    docDate = Convert.ToDateTime(dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.DocDate).Value).ToString("yyyyMMdd")
                    docTypeText = Convert.ToString(dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.DocType).Value)
                    asnNo = Convert.ToString(dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.ASNNo).Value)

                    strQry = "Select top 1 PDF_BINARY from SALESCHALLAN_SIGNED_PDFS (NOLOCK) where UNIT_CODE = '" + gstrUNITID + "' and DOC_NO = '" + docNo + "' and doc_type_text = '" + docTypeText + "'"
                    buffer = SqlConnectionclass.ExecuteScalar(strQry)
                    fileName = asnNo + "_" + docNo + "_" + docDate + ".pdf"
                    strFilePath = System.IO.Path.GetTempPath() + fileName
                    destFilePath = destFilePath + fileName
                    Try
                        If System.IO.File.Exists(strFilePath) Then
                            System.IO.File.Delete(strFilePath)
                        End If
                    Catch ex As Exception

                    End Try
                    System.IO.File.WriteAllBytes(strFilePath, buffer)
                    System.IO.File.Copy(strFilePath, destFilePath)

                    strQry = "UPDATE MAHINDRA_ASN_BARCODE_ACKNOWLEDGMENT SET ASN_DIGITAL_SIGN=1,ASN_DIGITAL_SIGN_DT=GETDATE(),ASN_DIGITAL_SIGN_BY='" & mP_User & "' WHERE UNIT_CODE='" & gstrUNITID & "' AND INVOICE_NO=" & docNo & " AND ISNULL(ASN_DIGITAL_SIGN,0)=0"
                    SqlConnectionclass.ExecuteNonQuery(strQry)
                    Try
                        If System.IO.File.Exists(strFilePath) Then
                            System.IO.File.Delete(strFilePath)
                        End If
                    Catch ex As Exception

                    End Try
                End If
            Next
            MsgBox("All selected invoices have been send to EDI successfully.", MsgBoxStyle.Information, ResolveResString(100))
            dgvInvoiceDetail.Rows.Clear()
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub btnShowInvoices_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowInvoices.Click
        Dim dtInvoices As New DataTable
        Try
            Dim strSql As String = String.Empty
            dgvInvoiceDetail.Rows.Clear()
            If (Not chkAllPendinginvoicesforEDIMapping.Checked) Then
                If dtpDateFrom.Value > dtpDateTo.Value Then
                    MessageBox.Show("[From Date] should be less than or equal to [To Date].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    dtpDateFrom.Focus()
                    Exit Sub
                End If
                If cmbInvoiceType.SelectedIndex = -1 Then
                    MessageBox.Show("Please select [Doc. Type].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    cmbInvoiceType.Focus()
                    Exit Sub
                End If
                If String.IsNullOrEmpty(txtCustomerCode.Text) Then
                    MessageBox.Show("Please select [Customer Code].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    txtCustomerCode.Focus()
                    Exit Sub
                End If
                If String.IsNullOrEmpty(txtInvoiceFrom.Text) Then
                    MessageBox.Show("Please select [From Invoice].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    txtInvoiceFrom.Focus()
                    Exit Sub
                End If
                If String.IsNullOrEmpty(txtInvoiceTo.Text) Then
                    MessageBox.Show("Please select [To Invoice].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    txtInvoiceTo.Focus()
                    Exit Sub
                End If
                If Not String.IsNullOrEmpty(txtInvoiceTo.Text) And Not String.IsNullOrEmpty(txtInvoiceFrom.Text) Then
                    If Val(txtInvoiceFrom.Text) > Val(txtInvoiceTo.Text) Then
                        MessageBox.Show("[From Invoice] should be less than or equal to [To Invoice].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        txtInvoiceTo.Focus()
                        Exit Sub
                    End If
                End If

                strSql = "SELECT DISTINCT S.DOC_NO,CONVERT(VARCHAR(11),S.INVOICE_DATE,103) as INVOICE_DATE,P.DOC_TYPE_TEXT,ISNULL(M.ASN_DIGITAL_SIGN,0) AS DOC_STATUS,S.ACCOUNT_CODE,S.CUST_NAME,S.INVOICE_TYPE,M.ASN_ACKNO "
                strSql += " FROM SALESCHALLAN_DTL S (NOLOCK) INNER JOIN SALESCHALLAN_SIGNED_PDFS P (NOLOCK) ON S.UNIT_CODE=P.UNIT_CODE AND S.DOC_NO=P.DOC_NO "
                strSql += " LEFT JOIN MAHINDRA_ASN_BARCODE_ACKNOWLEDGMENT M (NOLOCK) ON M.UNIT_CODE=S.UNIT_CODE AND M.INVOICE_NO=S.DOC_NO "
                strSql += " WHERE S.UNIT_CODE='" & gstrUNITID & "' AND S.BILL_FLAG=1 AND S.CANCEL_FLAG=0 AND S.ACCOUNT_CODE='" & Trim(txtCustomerCode.Text) & "' AND P.DOC_TYPE_TEXT='" & Trim(cmbInvoiceType.SelectedValue) & "' "
                strSql += " AND (S.INVOICE_DATE BETWEEN '" & VB6.Format(Me.dtpDateFrom.Value, "dd/MMM/yyyy") & "' AND '" & VB6.Format(Me.dtpDateTo.Value, "dd/MMM/yyyy") & "') "
                strSql += " AND (S.DOC_NO BETWEEN " & Trim(txtInvoiceFrom.Text) & " AND " & Trim(txtInvoiceTo.Text) & ")"
                strSql += " ORDER BY S.DOC_NO  "
            Else
                If cmbInvoiceType.SelectedIndex = -1 Then
                    MessageBox.Show("Please select [Doc. Type].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    cmbInvoiceType.Focus()
                    Exit Sub
                End If
                If cmbInvoiceType.SelectedValue.ToString().ToUpper() <> OriginalForBuyer Then
                    MessageBox.Show("Please select [ORIGINAL FOR BUYER] in [Doc. Type].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    cmbInvoiceType.Focus()
                    Exit Sub
                End If

                strSql = "SELECT DISTINCT S.DOC_NO,CONVERT(VARCHAR(11),S.INVOICE_DATE,103) as INVOICE_DATE,P.DOC_TYPE_TEXT,ISNULL(M.ASN_DIGITAL_SIGN,0) AS DOC_STATUS,S.ACCOUNT_CODE,S.CUST_NAME,S.INVOICE_TYPE,M.ASN_ACKNO "
                strSql += " FROM SALESCHALLAN_DTL S (NOLOCK) INNER JOIN SALESCHALLAN_SIGNED_PDFS P (NOLOCK) ON S.UNIT_CODE=P.UNIT_CODE AND S.DOC_NO=P.DOC_NO "
                strSql += " LEFT JOIN MAHINDRA_ASN_BARCODE_ACKNOWLEDGMENT M (NOLOCK) ON M.UNIT_CODE=S.UNIT_CODE AND M.INVOICE_NO=S.DOC_NO "
                strSql += " WHERE S.UNIT_CODE='" & gstrUNITID & "' AND S.BILL_FLAG=1 AND S.CANCEL_FLAG=0 AND P.DOC_TYPE_TEXT='" & Trim(cmbInvoiceType.SelectedValue) & "' AND ISNULL(M.ASN_DIGITAL_SIGN,0)=0 "
                strSql += " ORDER BY S.DOC_NO  "
            End If
            dtInvoices = SqlConnectionclass.GetDataTable(strSql)
            If dtInvoices IsNot Nothing AndAlso dtInvoices.Rows.Count > 0 Then
                Dim i As Integer = 0
                dgvInvoiceDetail.Rows.Add(dtInvoices.Rows.Count)
                For Each dr As DataRow In dtInvoices.Rows
                    dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.Selection).Value = False
                    dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.View).Value = "View"
                    dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.DocNo).Value = dr("DOC_NO")
                    dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.DocDate).Value = dr("INVOICE_DATE")
                    dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.DocType).Value = Convert.ToString(dr("DOC_TYPE_TEXT"))
                    dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.DocStatus).Value = IIf(Convert.ToBoolean(dr("DOC_STATUS")) = True, "Send", "Pending")
                    dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.AccountCode).Value = Convert.ToString(dr("ACCOUNT_CODE"))
                    dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.AccountName).Value = Convert.ToString(dr("CUST_NAME"))
                    dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.InvoiceType).Value = Convert.ToString(dr("INVOICE_TYPE"))
                    dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.ASNNo).Value = Convert.ToString(dr("ASN_ACKNO"))
                    i += 1
                Next
                dgvInvoiceDetail.Focus()
                dgvInvoiceDetail.CurrentCell = dgvInvoiceDetail.Rows(0).Cells(GridInvoiceDetails.Selection)
            Else
                MessageBox.Show("No Record(s) found.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                btnShowInvoices.Focus()
                Exit Sub
            End If
            rdbInvoiceCheckAll.Checked = False
            rdbInvoiceUncheckAll.Checked = False
        Catch ex As Exception
            RaiseException(ex)
        Finally
            If dtInvoices IsNot Nothing Then
                dtInvoices.Dispose()
            End If
        End Try
    End Sub

    Private Sub dgvInvoiceDetail_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvInvoiceDetail.CellClick
        Try
            If e.RowIndex < 0 Then Exit Sub
            If e.ColumnIndex = GridInvoiceDetails.View Then
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
                Dim docNo As String = Convert.ToString(dgvInvoiceDetail.Rows(e.RowIndex).Cells(GridInvoiceDetails.DocNo).Value)
                Dim docTypeText As String = Convert.ToString(dgvInvoiceDetail.Rows(e.RowIndex).Cells(GridInvoiceDetails.DocType).Value)
                DownloadPDFs(docNo, docTypeText)
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub DownloadPDFs(ByVal strDocNo As String, ByVal strDocTypeText As String)
        Try
            Dim strFilePath, strQry As String
            Dim buffer As Byte()

            strQry = "Select top 1 PDF_BINARY from SALESCHALLAN_SIGNED_PDFS (NOLOCK) where UNIT_CODE = '" + gstrUNITID + "' and DOC_NO = '" + strDocNo + "' and doc_type_text = '" + strDocTypeText + "'"
            buffer = SqlConnectionclass.ExecuteScalar(strQry)

            If strDocTypeText.ToUpper() = OriginalForBuyer Then
                strFilePath = System.IO.Path.GetTempPath() + strDocNo + "_org.pdf"
            ElseIf strDocTypeText.ToUpper() = DuplicateForTransporter Then
                strFilePath = System.IO.Path.GetTempPath() + strDocNo + "_dup.pdf"
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

                End Try
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
#Region "Bulk Download Pdf"

    'Modify By Anupam Kumar On the Date of 26 Sep 23.
    Dim FolderPath As String = String.Empty
    Dim dset As New DataSet

    Private Function getSignedInvoicesBulk() As Boolean
        Dim ValidField As Boolean = False
        Dim command As New SqlCommand()
        Try
            Dim strSql As String = String.Empty
            dset = New DataSet()

            If dtpDateFrom.Value > dtpDateTo.Value Then
                MessageBox.Show("[From Date] should be less than or equal to [To Date].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                dtpDateFrom.Focus()
                Return False
            End If
            If cmbInvoiceType.SelectedIndex = -1 Then
                MessageBox.Show("Please select [Doc. Type].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                cmbInvoiceType.Focus()
                Return False
            End If
            If String.IsNullOrEmpty(txtCustomerCode.Text) Then
                MessageBox.Show("Please select [Customer Code].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtCustomerCode.Focus()
                Return False
            End If

            strSql = "SELECT DISTINCT S.DOC_NO,CONVERT(VARCHAR(11),S.INVOICE_DATE,103) as INVOICE_DATE,P.DOC_TYPE_TEXT,ISNULL(M.ASN_DIGITAL_SIGN,0) AS DOC_STATUS,S.ACCOUNT_CODE,S.CUST_NAME,S.INVOICE_TYPE,M.ASN_ACKNO "
            strSql += " FROM SALESCHALLAN_DTL S (NOLOCK) INNER JOIN SALESCHALLAN_SIGNED_PDFS P (NOLOCK) ON S.UNIT_CODE=P.UNIT_CODE AND S.DOC_NO=P.DOC_NO "
            strSql += " LEFT JOIN MAHINDRA_ASN_BARCODE_ACKNOWLEDGMENT M (NOLOCK) ON M.UNIT_CODE=S.UNIT_CODE AND M.INVOICE_NO=S.DOC_NO "
            strSql += " WHERE S.UNIT_CODE='" & gstrUNITID & "' AND S.BILL_FLAG=1 AND S.CANCEL_FLAG=0 AND S.ACCOUNT_CODE='" & Trim(txtCustomerCode.Text) & "' AND P.DOC_TYPE_TEXT='" & Trim(cmbInvoiceType.SelectedValue) & "' "
            strSql += " AND (S.INVOICE_DATE BETWEEN '" & VB6.Format(Me.dtpDateFrom.Value, "dd/MMM/yyyy") & "' AND '" & VB6.Format(Me.dtpDateTo.Value, "dd/MMM/yyyy") & "') "
            ' strSql += " AND (S.DOC_NO BETWEEN " & Trim(txtInvoiceFrom.Text) & " AND " & Trim(txtInvoiceTo.Text) & ")"
            strSql += " ORDER BY S.DOC_NO ; "
            strSql += " SELECT ISNULL(CMST.CUST_VENDOR_CODE,'') AS VENDOR_CODE	FROM  CUSTOMER_MST CMST(NOLOCK) WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustomerCode.Text.Trim() & "' ; "

            dset = New DataSet()
            command = New SqlCommand()
            command.CommandText = strSql
            command.CommandTimeout = 0
            command.CommandType = CommandType.Text


            dset = SqlConnectionclass.GetDataSet(command)

            If dset IsNot Nothing AndAlso dset.Tables(0).Rows.Count > 0 Then
                ValidField = True
            Else
                MessageBox.Show("No Record(s) found.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                ValidField = False
            End If
            Return ValidField
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function

    Private Sub DownloadBulkPDFs(ByVal VendorCode As String, ByVal strFilePath As String, ByVal strDocNo As String, ByVal InvoiceDate As String, ByVal strDocTypeText As String)
        Try
            Dim strQry As String
            Dim FileType As String = String.Empty
            Dim buffer As Byte()

            strQry = "Select top 1 PDF_BINARY from SALESCHALLAN_SIGNED_PDFS (NOLOCK) where UNIT_CODE = '" + gstrUNITID + "' and DOC_NO = '" + strDocNo + "' and doc_type_text = '" + strDocTypeText + "'"
            buffer = SqlConnectionclass.ExecuteScalar(strQry)

            'If strDocTypeText.ToUpper() = OriginalForBuyer Then
            '    FileType = "_org.pdf"
            'ElseIf strDocTypeText.ToUpper() = DuplicateForTransporter Then
            '    FileType = "_dup.pdf"
            'Else
            '    FileType = ".pdf"
            'End If
            If String.IsNullOrEmpty(VendorCode) Then
                strFilePath = Path.Combine(strFilePath, strDocNo & ".pdf")
            Else
                strFilePath = Path.Combine(strFilePath, VendorCode & "_" & strDocNo & "_" & InvoiceDate & ".pdf")
            End If


            Try
                If System.IO.File.Exists(strFilePath) Then
                    System.IO.File.Delete(strFilePath)
                End If
                System.IO.File.WriteAllBytes(strFilePath, buffer)
            Catch ex As Exception

            End Try

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub Btn_DownloadAllInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_DownloadAllInvoice.Click
        Dim IsRecordExists As Boolean = False
        Dim VendorCode As String = String.Empty
        If getSignedInvoicesBulk() Then
            If (SelectFolderBrowserDialog.ShowDialog() = DialogResult.OK) Then
                FolderPath = SelectFolderBrowserDialog.SelectedPath
                If (dset IsNot Nothing AndAlso dset.Tables(1).Rows.Count > 0) Then
                    VendorCode = dset.Tables(1).Rows(0)(0).ToString()
                End If

                For Each rw As DataRow In dset.Tables(0).Rows
                    Dim InvoiceNo = rw.Item(0).ToString()
                    Dim InvoiceDate = rw.Item(1).ToString().Replace("/", "")
                    Dim strDocTypeText = rw.Item(2).ToString()
                    DownloadBulkPDFs(VendorCode, FolderPath, InvoiceNo, InvoiceDate, strDocTypeText)
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