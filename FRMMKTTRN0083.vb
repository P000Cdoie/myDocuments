Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Configuration

Public Class FRMMKTTRN0083
    Inherits System.Windows.Forms.Form

#Region "REVISION HISTORY"
    'Copyright          :   MIND Ltd.
    'Form Name          :   FRMMKTTRN0083
    'Created By         :   Ekta Uniyal
    'Created on         :   12 Jun 2014
    'Description        :   ASN File Generation 
    'Issue ID           :   10613846 — eMPro- CSV file generation for ASN 
    '********************************************************************************************************************
    'Issue ID           :   10613846 — eMPro- CSV File Generation: for more than one invoices, single file will generate now
    'Revised by         :   Prashant Rajpal
    '********************************************************************************************************************
    'Modified By        :   ASHISH SHARMA
    'Modified on        :   07 NOV 2017
    'Issue ID           :   101386343 - ASN for Normal Component
    '********************************************************************************************************************
    'Modified By        :   Shubhra Verma
    'Modified on        :   11 Sep 2019
    'Description        :   To Regenerate MAE Export Invoice ASNs
    '********************************************************************************************************************
#End Region

    Dim mintIndex As Short
    Dim mintFormIndex As Short
    Dim strSql As String = String.Empty
    Dim prevcustomercode As String = String.Empty
    Dim ButtonFlag As Boolean = False
    Dim strFullSql As String = String.Empty
#Region "FORM EVENTS"

    Private Sub FRMMKTTRN0083_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            mdifrmMain.CheckFormName = mintFormIndex
            System.Windows.Forms.Application.DoEvents()
            frmModules.NodeFontBold(Me.Tag) = True
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FRMMKTTRN0083_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader.Tag)
            Call FitToClient(Me, FrmMain, ctlFormHeader, CmdGrpBox, 100)

            dtToDate.Format = DateTimePickerFormat.Custom
            dtToDate.CustomFormat = gstrDateFormat
            dtToDate.Value = GetServerDate()

            dtFromDate.Format = DateTimePickerFormat.Custom
            dtFromDate.CustomFormat = gstrDateFormat
            dtFromDate.Value = GetServerDate()
            AddInvoiceTypeInCombo()
            txtCustomerCode.Text = ""
            lblCustomerDesc.Text = ""
            txtFromInvoice.Text = ""
            txtToInvoice.Text = ""

            txtCustomerCode.TabIndex = 0
            txtCustomerCode.Focus()
            optFileType.Checked = True
            btnGenerateFile.Enabled = True
            btnSendASNViaAPI.Visible = False
            TML_ASN_GENERATION_Via_API_Button_Flag()
            RichTextBox1.Visible = False
            btnexception.Visible = False
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub FRMMKTTRN0083_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Try
            Me.Dispose()
            Exit Sub
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub FRMMKTTRN0083_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        Try
            frmModules.NodeFontBold(Me.Tag) = False
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub TML_ASN_GENERATION_Via_API_Button_Flag()
        Try
            ButtonFlag = SqlConnectionclass.ExecuteScalar("Select IsNull(TML_ASN_GENERATION_API,0) TML_ASN_GENERATION_API from SALES_PARAMETER Where Company_Code = '" & gstrUNITID & "'")
        Catch ex As Exception
            RaiseException(ex)
        End Try


        If ButtonFlag Then
            btnSendASNViaAPI.Visible = True
        Else
            btnSendASNViaAPI.Visible = False
        End If
    End Sub
    Private Sub FRMMKTTRN0083_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            Dim KeyCode As Short = e.KeyCode
            Dim Shift As Short = e.KeyData \ &H10000
            If Shift <> 0 Then Exit Sub
            If KeyCode = System.Windows.Forms.Keys.F4 Then Call ctlFormHeader_Click(ctlFormHeader, New System.EventArgs()) : Exit Sub
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

#End Region

#Region "FORM CONTROL EVENTS"

    Private Sub cmdHlpCustomerCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHlpCustomerCode.Click

        Dim strHelp() As String = Nothing

        Try
            If optFileType.Checked = True Then
                strSql = " SELECT Cast(Customer_Code AS Varchar(50)) AS Customer_Code,Cust_Name AS Customer_Name" & _
                     " FROM Customer_Mst X " & _
                     " WHERE    UNIT_CODE='" & gstrUNITID & "' " & _
                     "          And Exists(Select Top 1 1 from FORMATE_CUSTOMER_LINKAGE Y Where X.unit_Code=Y.Unit_Code And X.customer_Code=Y.Customer_Code) " & _
                     " ORDER BY Customer_Code"
            Else
                strSql = " SELECT Cast(Customer_Code AS Varchar(50)) AS Customer_Code,Cust_Name AS Customer_Name" & _
                     " FROM Customer_Mst X " & _
                     " WHERE    UNIT_CODE='" & gstrUNITID & "' and MAHINDRA_ASN_ENABLED = 1 " & _
                     " ORDER BY Customer_Code"
            End If
            
            strHelp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "Customer Help")
            If Not (UBound(strHelp) <= 0) Then
                If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                    MessageBox.Show("No record To Display", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    txtCustomerCode.Text = ""
                    lblCustomerDesc.Text = ""
                    Exit Sub
                Else
                    txtCustomerCode.Text = strHelp(0).Trim
                    lblCustomerDesc.Text = strHelp(1).Trim
                    txtCustomerCode_Validating(txtCustomerCode, New System.ComponentModel.CancelEventArgs(False))
                End If
            End If
            If SqlConnectionclass.ExecuteScalar("SELECT TOP 1 1 FROM Customer_mst where unit_code='" & gstrUNITID & "' and customer_code='" & txtCustomerCode.Text.Trim & "' and TML_ASN_ENABLED=1") Then
                btnexception.Visible = True
            Else
                btnexception.Visible = False
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtCustomerCode_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.GotFocus
        Try
            With txtCustomerCode
                .SelectionStart = 0
                .SelectionLength = txtCustomerCode.Text.Trim.Length
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtCustomerCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown

        Try
            If e.KeyCode = Keys.F1 Then
                Call cmdHlpCustomerCode_Click(cmdHlpCustomerCode, New System.EventArgs)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtCustomerCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustomerCode.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Try
            Select Case KeyAscii
                Case 39, 34, 96
                    Beep()
                    e.Handled = True
                Case 13
                    If txtCustomerCode.Text.Trim.Length > 0 Then
                        dtFromDate.Focus()
                    End If
            End Select
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtCustomerCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.TextChanged

        Try
            If txtCustomerCode.Text.Length = 0 Then
                txtCustomerCode.Text = ""
                lblCustomerDesc.Text = ""
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtCustomerCode_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating

        Dim oSqlDr As SqlDataReader = Nothing

        Try
            If prevcustomercode = txtCustomerCode.Text.Trim Then
                Exit Sub
            End If
            If txtCustomerCode.Text.Length > 0 Then

                txtCustomerCode.Text = Replace(txtCustomerCode.Text, "'", "")
                prevcustomercode = txtCustomerCode.Text.Trim

                txtFromInvoice.Text = String.Empty
                txtToInvoice.Text = String.Empty


                strSql = " SELECT Customer_Code as Customer_Code,Cust_Name " & _
                         " FROM Customer_Mst X " & _
                         " WHERE    UNIT_CODE='" & gstrUNITID & "' AND Customer_Code='" & txtCustomerCode.Text.Trim & "' " & _
                         "          And Exists(Select Top 1 1 from FORMATE_CUSTOMER_LINKAGE Y Where X.unit_Code=Y.Unit_Code And X.customer_Code=Y.Customer_Code) " & _
                         " ORDER BY Customer_Code"
                oSqlDr = SqlConnectionclass.ExecuteReader(strSql)
                If oSqlDr.HasRows = True Then
                    oSqlDr.Read()
                    txtCustomerCode.Text = oSqlDr("Customer_Code").ToString.Trim
                    lblCustomerDesc.Text = oSqlDr("Cust_Name").ToString.Trim
                Else
                    MessageBox.Show("Customer Code does not exists !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    txtCustomerCode.Text = ""
                    txtCustomerCode.Focus()
                    Exit Sub
                End If

            End If
            If Not oSqlDr Is Nothing AndAlso oSqlDr.IsClosed = False Then oSqlDr.Close()

        Catch ex As Exception
            RaiseException(ex)
        Finally
            If Not oSqlDr Is Nothing AndAlso oSqlDr.IsClosed = False Then oSqlDr.Close()
        End Try

    End Sub

    Private Sub cmdFrmInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFrmInvoice.Click

        Dim strHelp() As String = Nothing

        Try
            If Len(Trim(txtCustomerCode.Text)) = 0 Then
                MessageBox.Show("Select Customer Code First.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtCustomerCode.Focus()
                Exit Sub
            ElseIf dtFromDate.Value > dtToDate.Value Then
                MessageBox.Show("[From date] should be less than or equal to [To date].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                dtFromDate.Focus()
                Exit Sub
            ElseIf String.IsNullOrEmpty(cmbInvoiceType.Text) Then
                MessageBox.Show("Select Invoice Type First.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                cmbInvoiceType.Focus()
                Exit Sub
            Else
                If optFileType.Checked = True Then
                    strSql = " SELECT CAST(Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),Invoice_Date,103) as Invoice_Date" & _
                         " FROM SALESCHALLAN_DTL " & _
                         " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                         " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 AND Invoice_Type in ('INV','EXP') AND Sub_Category='" & Convert.ToString(cmbInvoiceType.SelectedValue) & "'" & _
                         " AND UNIT_CODE = '" & gstrUNITID & "' And Account_Code='" & txtCustomerCode.Text.Trim & "' " & _
                         " AND EWAY_IRN_REQUIRED='N' " & _
                         " UNION SELECT CAST(S.Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),S.Invoice_Date,103) as Invoice_Date " & _
                         " FROM SALESCHALLAN_DTL S" & _
                         " LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO " & _
                         " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                         " AND S.Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 AND S.Invoice_Type in ('INV','EXP') AND Sub_Category='" & Convert.ToString(cmbInvoiceType.SelectedValue) & "'" & _
                         " AND S.UNIT_CODE = '" & gstrUNITID & "' And S.Account_Code='" & txtCustomerCode.Text.Trim & "' " & _
                         " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) "

                    If SqlConnectionclass.ExecuteScalar("SELECT TOP 1 1 FROM Customer_mst where unit_code='" & gstrUNITID & "' and customer_code='" & txtCustomerCode.Text.Trim & "' and TML_ASN_ENABLED=1") Then
                        strFullSql = " AND NOT EXISTS ( SELECT TOP  1 1 FROM TATA_CYGNET_ASN_ACKNOWLEDGMENT TML " & _
                        " WHERE TML.UNIT_CODE = S.UNIT_CODE  AND TML.INVOICE_NO=S.DOC_NO AND ASN_STATUS <>'Error' ) "
                        strSql = strSql + strFullSql + " ORDER BY Invoice_No"
                    Else
                        strSql = strSql + " ORDER BY Invoice_No"
                    End If


                Else
                    strSql = " SELECT CAST(Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),Invoice_Date,103) as Invoice_Date" & _
                         " FROM SALESCHALLAN_DTL " & _
                         " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                         " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 AND Invoice_Type in ('INV','EXP') AND Sub_Category='" & Convert.ToString(cmbInvoiceType.SelectedValue) & "'" & _
                         " AND UNIT_CODE = '" & gstrUNITID & "' And Account_Code='" & txtCustomerCode.Text.Trim & "' " & _
                         " AND EWAY_IRN_REQUIRED='N' " & _
                         " AND NOT EXISTS " & _
                         " ( SELECT TOP 1 1 FROM MAHINDRA_ASN_BARCODE_ACKNOWLEDGMENT MM WHERE MM.UNIT_CODE=SALESCHALLAN_DTL.UNIT_CODE AND MM.INVOICE_NO=SALESCHALLAN_DTL.DOC_NO )" & _
                         " UNION SELECT CAST(S.Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),S.Invoice_Date,103) as Invoice_Date " & _
                         " FROM SALESCHALLAN_DTL S" & _
                         " LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO " & _
                         " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                         " AND S.Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 AND S.Invoice_Type in ('INV','EXP') AND Sub_Category='" & Convert.ToString(cmbInvoiceType.SelectedValue) & "'" & _
                         " AND S.UNIT_CODE = '" & gstrUNITID & "' And S.Account_Code='" & txtCustomerCode.Text.Trim & "' " & _
                         " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) " & _
                         " AND NOT EXISTS " & _
                         " ( SELECT TOP 1 1 FROM MAHINDRA_ASN_BARCODE_ACKNOWLEDGMENT MM WHERE MM.UNIT_CODE=S.UNIT_CODE AND MM.INVOICE_NO=S.DOC_NO )" & _
                         " ORDER BY Invoice_No"
                End If

                strHelp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "Invoice No. Help")
                If Not (UBound(strHelp) <= 0) Then
                    If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                        MessageBox.Show("No record To Display", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        txtFromInvoice.Text = ""
                        Exit Sub
                    Else
                        txtFromInvoice.Text = strHelp(0).Trim
                        txtFromInvoice_Validating(txtFromInvoice, New System.ComponentModel.CancelEventArgs(False))
                    End If
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub cmdToInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdToInvoice.Click

        Dim strHelp() As String = Nothing

        Try
            If Len(Trim(txtCustomerCode.Text)) = 0 Then
                MessageBox.Show("Select Customer Code First.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtCustomerCode.Focus()
                Exit Sub
            ElseIf dtFromDate.Value > dtToDate.Value Then
                MessageBox.Show("[From date] should be less than or equal to [To date].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                dtFromDate.Focus()
                Exit Sub
            ElseIf String.IsNullOrEmpty(cmbInvoiceType.Text) Then
                MessageBox.Show("Select Invoice Type First.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                cmbInvoiceType.Focus()
                Exit Sub
            Else
                strSql = String.Empty
                If optFileType.Checked = True Then
                    strSql = " SELECT CAST(Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),Invoice_Date,103) as Invoice_Date" & _
                         " FROM SALESCHALLAN_DTL " & _
                         " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                         " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 AND Invoice_Type in ('INV','EXP') AND Sub_Category='" & Convert.ToString(cmbInvoiceType.SelectedValue) & "'" & _
                         " AND UNIT_CODE = '" & gstrUNITID & "' And Account_Code='" & txtCustomerCode.Text.Trim & "' " & _
                         " AND EWAY_IRN_REQUIRED='N' " & _
                         " UNION SELECT CAST(S.Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),S.Invoice_Date,103) as Invoice_Date " & _
                         " FROM SALESCHALLAN_DTL S" & _
                         " LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO " & _
                         " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                         " AND S.Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 AND S.Invoice_Type in ('INV','EXP') AND Sub_Category='" & Convert.ToString(cmbInvoiceType.SelectedValue) & "'" & _
                         " AND S.UNIT_CODE = '" & gstrUNITID & "' And Account_Code='" & txtCustomerCode.Text.Trim & "' " & _
                         " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) "

                    If SqlConnectionclass.ExecuteScalar("SELECT TOP 1 1 FROM Customer_mst where unit_code='" & gstrUNITID & "' and customer_code='" & txtCustomerCode.Text.Trim & "' and TML_ASN_ENABLED=1") Then
                        strFullSql = " AND NOT EXISTS ( SELECT TOP  1 1 FROM TATA_CYGNET_ASN_ACKNOWLEDGMENT TML " & _
                        " WHERE TML.UNIT_CODE = S.UNIT_CODE  AND TML.INVOICE_NO=S.DOC_NO AND  ASN_STATUS <>'Error' ) "
                        strSql = strSql + strFullSql + " ORDER BY Invoice_No"
                    Else
                        strSql = strSql + " ORDER BY Invoice_No"
                    End If

                    ' " ORDER BY Invoice_No"
                Else
                    strSql = " SELECT CAST(Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),Invoice_Date,103) as Invoice_Date" & _
                         " FROM SALESCHALLAN_DTL " & _
                         " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                         " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 AND Invoice_Type in ('INV','EXP') AND Sub_Category='" & Convert.ToString(cmbInvoiceType.SelectedValue) & "'" & _
                         " AND UNIT_CODE = '" & gstrUNITID & "' And Account_Code='" & txtCustomerCode.Text.Trim & "' " & _
                         " AND EWAY_IRN_REQUIRED='N' " & _
                         " AND NOT EXISTS " & _
                         " ( SELECT TOP 1 1 FROM MAHINDRA_ASN_BARCODE_ACKNOWLEDGMENT MM WHERE MM.UNIT_CODE=SALESCHALLAN_DTL.UNIT_CODE AND MM.INVOICE_NO=SALESCHALLAN_DTL.DOC_NO )" & _
                          " UNION SELECT CAST(S.Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),S.Invoice_Date,103) as Invoice_Date " & _
                         " FROM SALESCHALLAN_DTL S" & _
                         " LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO " & _
                         " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                         " AND S.Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 AND S.Invoice_Type in ('INV','EXP') AND Sub_Category='" & Convert.ToString(cmbInvoiceType.SelectedValue) & "'" & _
                         " AND S.UNIT_CODE = '" & gstrUNITID & "' And Account_Code='" & txtCustomerCode.Text.Trim & "' " & _
                         " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) " & _
                         " AND NOT EXISTS " & _
                          " ( SELECT TOP 1 1 FROM MAHINDRA_ASN_BARCODE_ACKNOWLEDGMENT MM WHERE MM.UNIT_CODE=S.UNIT_CODE AND MM.INVOICE_NO=S.DOC_NO )" & _
                         " ORDER BY Invoice_No"

                End If
                strHelp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "Invoice No. Help")
                If Not (UBound(strHelp) <= 0) Then
                    If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                        MessageBox.Show("No record To Display", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        txtFromInvoice.Text = ""
                        Exit Sub
                    Else
                        txtToInvoice.Text = strHelp(0).Trim
                        txtToInvoice_Validating(txtToInvoice, New System.ComponentModel.CancelEventArgs(False))
                    End If
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtToInvoice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtToInvoice.GotFocus
        Try
            With txtToInvoice
                .SelectionStart = 0
                .SelectionLength = txtToInvoice.Text.Trim.Length
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtToInvoice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtToInvoice.KeyDown

        Try
            If e.KeyCode = Keys.F1 Then
                Call cmdToInvoice_Click(cmdToInvoice, New System.EventArgs)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtToInvoice_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtToInvoice.Validating

        Try

            txtToInvoice.Text = Replace(txtToInvoice.Text, "'", "")
            If txtToInvoice.Text.Trim.Length = 0 Then Return

            strSql = " SELECT Doc_No As Invoice_No " & _
                     " FROM SALESCHALLAN_DTL " & _
                      " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                     " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 AND Invoice_Type in ('INV','EXP')" & _
                     " AND Sub_Category='" & Convert.ToString(cmbInvoiceType.SelectedValue) & "' AND UNIT_CODE = '" & gstrUNITID & "'" & _
                     " And Account_Code='" & txtCustomerCode.Text.Trim & "' " & _
                     " AND Doc_No='" & txtToInvoice.Text.Trim & "' AND EWAY_IRN_REQUIRED='N'" & _
                     " UNION SELECT CAST(S.Doc_No AS VARCHAR(20)) As Invoice_No " & _
                    " FROM SALESCHALLAN_DTL S" & _
                    " LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO " & _
                    " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                    " AND S.Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 AND S.Invoice_Type in ('INV','EXP') AND Sub_Category='" & Convert.ToString(cmbInvoiceType.SelectedValue) & "'" & _
                    " AND S.UNIT_CODE = '" & gstrUNITID & "' And S.Account_Code='" & txtCustomerCode.Text.Trim & "' " & _
                    " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) " & _
                    " AND s.Doc_No='" & txtToInvoice.Text.Trim & "' "

            If SqlConnectionclass.ExecuteScalar("SELECT TOP 1 1 FROM Customer_mst where unit_code='" & gstrUNITID & "' and customer_code='" & txtCustomerCode.Text.Trim & "' and TML_ASN_ENABLED=1") Then
                strFullSql = " AND NOT EXISTS ( SELECT TOP  1 1 FROM TATA_CYGNET_ASN_ACKNOWLEDGMENT TML " & _
                " WHERE TML.UNIT_CODE = S.UNIT_CODE  AND TML.INVOICE_NO=S.DOC_NO AND ASN_STATUS <>'Error' ) "
                strSql = strSql + strFullSql
            End If

            If IsRecordExists(strSql) = False Then
                MessageBox.Show("Selected Invoice No. does not exists.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtToInvoice.Text = ""
                txtToInvoice.Focus()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtFromInvoice_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFromInvoice.GotFocus
        Try
            With txtFromInvoice
                .SelectionStart = 0
                .SelectionLength = txtFromInvoice.Text.Trim.Length
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtFromInvoice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFromInvoice.KeyDown

        Try

            If e.KeyCode = Keys.F1 Then
                Call cmdFrmInvoice_Click(cmdFrmInvoice, New System.EventArgs)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtFromInvoice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFromInvoice.KeyPress

        Dim KeyAscii As Short = Asc(e.KeyChar)

        Try
            Select Case KeyAscii
                Case 13
                    If txtCustomerCode.Text.Trim.Length > 0 Then
                        txtToInvoice.Focus()
                    End If
            End Select

            KeyAscii = validateKey(txtFromInvoice.Text, Len(Me.txtFromInvoice.Text), KeyAscii, 12, 0)
            e.KeyChar = Chr(KeyAscii)

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub txtFromInvoice_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFromInvoice.Validating

        Try

            txtFromInvoice.Text = Replace(txtFromInvoice.Text, "'", "")
            If txtFromInvoice.Text.Trim.Length = 0 Then Return

            strSql = " SELECT Doc_No As Invoice_No " & _
                     " FROM SALESCHALLAN_DTL " & _
                      " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                     " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 AND Invoice_Type in ('INV','EXP')" & _
                     " AND Sub_Category='" & Convert.ToString(cmbInvoiceType.SelectedValue) & "' AND UNIT_CODE = '" & gstrUNITID & "'" & _
                     " And Account_Code='" & txtCustomerCode.Text.Trim & "' " & _
                     " AND Doc_No='" & txtFromInvoice.Text.Trim & "' AND EWAY_IRN_REQUIRED='N'" & _
                    " UNION SELECT CAST(S.Doc_No AS VARCHAR(20)) As Invoice_No " & _
                    " FROM SALESCHALLAN_DTL S" & _
                    " LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO " & _
                    " WHERE Convert(Date,S.Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                    " AND S.Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 AND S.Invoice_Type in ('INV','EXP') AND Sub_Category='" & Convert.ToString(cmbInvoiceType.SelectedValue) & "'" & _
                    " AND S.UNIT_CODE = '" & gstrUNITID & "' And S.Account_Code='" & txtCustomerCode.Text.Trim & "' " & _
                    " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>''))  AND s.Doc_No='" & txtFromInvoice.Text.Trim & "' "

            If IsRecordExists(strSql) = False Then
                MessageBox.Show("Selected Invoice No. does not exists.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtFromInvoice.Text = ""
                txtFromInvoice.Focus()
                Exit Sub
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtFromInvoice_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFromInvoice.TextChanged

        Try
            If txtFromInvoice.Text.Length = 0 Then
                txtFromInvoice.Text = ""
                txtToInvoice.Text = ""
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtToInvoice_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtToInvoice.KeyPress

        Dim KeyAscii As Short = Asc(e.KeyChar)

        Try
            Select Case KeyAscii
                Case 13
                    If txtCustomerCode.Text.Trim.Length > 0 Then
                        btnGenerateFile.Focus()
                    End If
            End Select

            KeyAscii = validateKey(txtFromInvoice.Text, Len(Me.txtFromInvoice.Text), KeyAscii, 12, 0)
            e.KeyChar = Chr(KeyAscii)



        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub txtToInvoice_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtToInvoice.TextChanged

        Try
            If txtToInvoice.Text.Length = 0 Then
                txtFromInvoice.Text = ""
                txtToInvoice.Text = ""
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub ctlFormHeader_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader.Click

        Try
            Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub dtFromDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dtFromDate.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Try
            Select Case KeyAscii
                Case 13
                    dtToDate.Focus()
            End Select
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub dtToDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dtToDate.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Try
            Select Case KeyAscii
                Case 13
                    txtFromInvoice.Focus()
            End Select
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

#End Region

#Region "ROUTINES"

    Private Sub GenerateASNFile()

        Dim oSqlDr As SqlDataReader = Nothing
        Dim Dt As New DataTable
        Dim sw As StreamWriter = Nothing
        Dim iColCount As Integer = 0
        Dim intCol As Integer = 0
        Dim intLoopCounter As Integer = 0
        Dim strFinalQry As String = String.Empty
        Dim strValDocNo As String = String.Empty
        Dim strQuery As String = String.Empty
        Dim strFileLocation As String = String.Empty
        Dim strASNFilePath As String = String.Empty
        Dim fs As FileStream = Nothing
        Dim Obj_FSO As Scripting.FileSystemObject = Nothing

        Try
            If Validate_Data() = True Then
                strSql = " SELECT Query " & _
                         " FROM Formate_Mst A " & _
                         " INNER JOIN Formate_Customer_Linkage B On B.Formate_ID =A.Formate_ID " & _
                         " AND B.Unit_Code =A.Unit_Code " & _
                         " WHERE Customer_Code='" & txtCustomerCode.Text.Trim & "' And A.IsActive=1 " & _
                         " AND A.UNIT_CODE='" & gstrUNITID & "'"
                strQuery = SqlConnectionclass.ExecuteScalar(strSql)
                If strQuery <> Nothing Or strQuery <> String.Empty Then
                    '10613846 
                    strSql = " SELECT Doc_No as Invoice_No " & _
                             " FROM SALESCHALLAN_DTL " & _
                             " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                             " AND Doc_No BETWEEN '" & txtFromInvoice.Text & "' AND '" & txtToInvoice.Text & "' " & _
                             " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 AND Invoice_Type in ('INV','EXP') AND Sub_Category='" & Convert.ToString(cmbInvoiceType.SelectedValue) & "'" & _
                             " AND Account_Code='" & txtCustomerCode.Text.Trim & "' " & _
                             " AND UNIT_CODE = '" & gstrUNITID & "' " & _
                             " "
                    strFinalQry = strQuery + " WHERE X.Doc_No in  (" & strSql & ") "

                    'oSqlDr = SqlConnectionclass.ExecuteReader(strSql)
                    'If oSqlDr.HasRows = True Then
                    '    While oSqlDr.Read
                    '        strValDocNo = oSqlDr("Invoice_No").ToString
                    '        strFinalQry = strQuery + " WHERE X.Doc_No=" & strValDocNo & " "
                    '    End While

                    '10613846 
                    If Dt.Rows.Count > 0 Then Dt.Clear()
                    Dt = SqlConnectionclass.GetDataTable(strFinalQry)

                    If Dt.Rows.Count = 0 Then
                        MessageBox.Show("No Data Found For File Generation.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Exit Sub
                    End If

                    strFileLocation = FN_Get_Folder_Path()
                    If strFileLocation.Trim = "" Then
                        MessageBox.Show("ASN File path not defined in Sales Parameter.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Exit Sub
                    Else
                        Obj_FSO = New Scripting.FileSystemObject
                        If Not Obj_FSO.FolderExists(strFileLocation) Then
                            Obj_FSO.CreateFolder(strFileLocation)
                        End If

                        If Mid(Trim(strFileLocation), Len(Trim(strFileLocation))) <> "\" Then
                            strFileLocation = strFileLocation & "\"
                        End If

                        strASNFilePath = strFileLocation & "ASN_" & txtCustomerCode.Text.Trim() & "_from_" & txtFromInvoice.Text.Trim() & "_To_" & txtToInvoice.Text.Trim() & "_" & Now.Day & Now.Month & Now.Year & Now.Hour & Now.Minute & Now.Second & ".csv"
                        fs = File.Create(strASNFilePath)
                        sw = New StreamWriter(fs)

                        'Setting Header First
                        iColCount = Dt.Columns.Count
                        For intCol = 0 To iColCount - 1
                            sw.Write(Dt.Columns(intCol).ToString)
                            If intCol < iColCount - 1 Then
                                sw.Write(",")
                            End If
                        Next
                        sw.Write(sw.NewLine)

                        'Seeting Row Values
                        For Each row As DataRow In Dt.Rows
                            For intCol = 0 To iColCount - 1
                                If IsDBNull(row(intCol)) = False Then
                                    sw.Write(row(intCol).ToString)
                                End If
                                If intCol < iColCount - 1 Then
                                    sw.Write(",")
                                End If
                            Next
                            sw.Write(System.Environment.NewLine)
                        Next
                    End If
                    If Not sw Is Nothing Then sw.Close()
                    If Not fs Is Nothing Then fs.Close()
                    If Not Obj_FSO Is Nothing Then Obj_FSO = Nothing

                    MessageBox.Show("ASN Files are Generated Succesfully. " & vbCrLf & "Files generated on the following Path - " & strFileLocation, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                End If
            Else
                Exit Sub
            End If

            If Not oSqlDr Is Nothing AndAlso oSqlDr.IsClosed = False Then oSqlDr.Close()
            If Dt.Rows.Count > 0 Then Dt.Dispose()

        Catch ex As Exception
            RaiseException(ex)
        Finally
            If Not sw Is Nothing Then sw.Close()
            If Not fs Is Nothing Then fs.Close()
            If Not Obj_FSO Is Nothing Then Obj_FSO = Nothing
            If Not oSqlDr Is Nothing AndAlso oSqlDr.IsClosed = False Then oSqlDr.Close()
            If Dt.Rows.Count > 0 Then Dt.Dispose()
        End Try

    End Sub
    Private Sub GenerateASNFile_TML()

        Dim oSqlDr As SqlDataReader = Nothing
        Dim Dt As New DataTable
        Dim sw As StreamWriter = Nothing
        Dim iColCount As Integer = 0
        Dim intCol As Integer = 0
        Dim intLoopCounter As Integer = 0
        Dim strFinalQry As String = String.Empty
        Dim strValDocNo As String = String.Empty
        Dim strQuery As String = String.Empty
        Dim strFileLocation As String = String.Empty
        Dim strASNFilePath As String = String.Empty
        Dim fs As FileStream = Nothing
        Dim Obj_FSO As Scripting.FileSystemObject = Nothing
        Dim strTATAACKN As String = String.Empty


        Try
            If Validate_Data() = True Then
                strSql = " SELECT Query " & _
                         " FROM Formate_Mst A " & _
                         " INNER JOIN Formate_Customer_Linkage B On B.Formate_ID =A.Formate_ID " & _
                         " AND B.Unit_Code =A.Unit_Code " & _
                         " WHERE Customer_Code='" & txtCustomerCode.Text.Trim & "' And A.IsActive=1 " & _
                         " AND A.UNIT_CODE='" & gstrUNITID & "'"
                strQuery = SqlConnectionclass.ExecuteScalar(strSql)
                If strQuery <> Nothing Or strQuery <> String.Empty Then
                    '10613846 
                    strSql = " SELECT Doc_No as Invoice_No " & _
                             " FROM SALESCHALLAN_DTL " & _
                             " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                             " AND Doc_No BETWEEN '" & txtFromInvoice.Text & "' AND '" & txtToInvoice.Text & "' " & _
                             " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 AND Invoice_Type in ('INV','EXP') AND Sub_Category='" & Convert.ToString(cmbInvoiceType.SelectedValue) & "'" & _
                             " AND Account_Code='" & txtCustomerCode.Text.Trim & "' " & _
                             " AND UNIT_CODE = '" & gstrUNITID & "' " & _
                             " "
                    

                    If SqlConnectionclass.ExecuteScalar("SELECT TOP 1 1 FROM Customer_mst where unit_code='" & gstrUNITID & "' and customer_code='" & txtCustomerCode.Text.Trim & "' and TML_ASN_ENABLED=1") Then
                        strFinalQry = strQuery + " WHERE X.Doc_No in  (" & strSql & ") " + " AND NOT EXISTS ( SELECT TOP  1 1 FROM TATA_CYGNET_ASN_ACKNOWLEDGMENT TML " & _
                            " WHERE TML.UNIT_CODE = X.UNIT_CODE  AND TML.INVOICE_NO=X.DOC_NO AND ASN_STATUS <>'Error' ) "
                    Else
                        strFinalQry = strQuery + " WHERE X.Doc_No in  (" & strSql & ") "
                    End If
                    'oSqlDr = SqlConnectionclass.ExecuteReader(strSql)
                    'If oSqlDr.HasRows = True Then
                    '    While oSqlDr.Read
                    '        strValDocNo = oSqlDr("Invoice_No").ToString
                    '        strFinalQry = strQuery + " WHERE X.Doc_No=" & strValDocNo & " "
                    '    End While

                    '10613846 
                    If Dt.Rows.Count > 0 Then Dt.Clear()
                    Dt = SqlConnectionclass.GetDataTable(strFinalQry)

                    If Dt.Rows.Count = 0 Then
                        MessageBox.Show("No Data Found For File Generation.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Exit Sub
                    End If

                    strFileLocation = FN_Get_CYGNET_Folder_Path()
                    If strFileLocation.Trim = "" Then
                        MessageBox.Show("ASN File path not defined in Sales Parameter.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Exit Sub
                    Else
                        Obj_FSO = New Scripting.FileSystemObject
                        If Not Obj_FSO.FolderExists(strFileLocation) Then
                            Obj_FSO.CreateFolder(strFileLocation)
                        End If

                        If Mid(Trim(strFileLocation), Len(Trim(strFileLocation))) <> "\" Then
                            strFileLocation = strFileLocation & "\"
                        End If
                        For Each row As DataRow In Dt.Rows

                            If gstrUNITID = "MPU" Then
                                'strASNFilePath = strFileLocation & "MATEPune_13_9_INVOICENO" & Now.Day & Now.Month & Now.Year & Now.Hour & Now.Minute & Now.Second & ".csv"
                                strASNFilePath = strFileLocation & "MATEPune_13_9_INVOICENO" & Convert.ToString(row("InvoiceNumber")) & "_" & Now.Day & Now.Month & Now.Year & Now.Hour & Now.Minute & Now.Second & ".csv"
                            ElseIf gstrUNITID = "SMP" Or gstrUNITID = "SMC" Then
                                strASNFilePath = strFileLocation & "SMRC_13_9_INVOICENO" & Convert.ToString(row("InvoiceNumber")) & "_" & Now.Day & Now.Month & Now.Year & Now.Hour & Now.Minute & Now.Second & ".csv"
                            Else
                                strASNFilePath = strFileLocation & "ASN_" & txtCustomerCode.Text.Trim() & "_from_" & txtFromInvoice.Text.Trim() & "_To_" & txtToInvoice.Text.Trim() & "_" & Now.Day & Now.Month & Now.Year & Now.Hour & Now.Minute & Now.Second & ".csv"
                            End If

                            strTATAACKN = "insert into TATA_CYGNET_ASN_ACKNOWLEDGMENT(UNIT_CODE,CUSTOMER_CODE,INVOICE_NO,IS_ASN_ACK)"
                            strTATAACKN += " select '" + gstrUNITID + "','" + txtCustomerCode.Text & "','" + Convert.ToString(row("InvoiceNumber")) + "',0"
                            SqlConnectionclass.ExecuteNonQuery(strTATAACKN)

                            fs = File.Create(strASNFilePath)
                            sw = New StreamWriter(fs)

                            'Setting Header First
                            iColCount = Dt.Columns.Count
                            For intCol = 0 To iColCount - 1
                                sw.Write(Dt.Columns(intCol).ToString)
                                If intCol < iColCount - 1 Then
                                    sw.Write(",")
                                End If
                            Next
                            sw.Write(sw.NewLine)

                            'Seeting Row Values
                            '
                            For intCol = 0 To iColCount - 1
                                If IsDBNull(row(intCol)) = False Then
                                    sw.Write(row(intCol).ToString)
                                End If
                                'If intCol = 1 Then
                                'strASNFilePath = strASNFilePath & row(intCol).ToString & Now.Day & Now.Month & Now.Year & Now.Hour & Now.Minute & Now.Second & ".csv"
                                'End If
                                If intCol < iColCount - 1 Then
                                    sw.Write(",")
                                End If
                            Next
                            sw.Write(System.Environment.NewLine)


                            If Not sw Is Nothing Then sw.Close()
                            If Not fs Is Nothing Then fs.Close()
                            If Not Obj_FSO Is Nothing Then Obj_FSO = Nothing
                        Next
                    End If
                    MessageBox.Show("ASN Files are Generated Succesfully. " & vbCrLf & "Files generated on the following Path - " & strFileLocation, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                End If
            Else
                Exit Sub
            End If

            If Not oSqlDr Is Nothing AndAlso oSqlDr.IsClosed = False Then oSqlDr.Close()
            If Dt.Rows.Count > 0 Then Dt.Dispose()

        Catch ex As Exception
            RaiseException(ex)
        Finally
            If Not sw Is Nothing Then sw.Close()
            If Not fs Is Nothing Then fs.Close()
            If Not Obj_FSO Is Nothing Then Obj_FSO = Nothing
            If Not oSqlDr Is Nothing AndAlso oSqlDr.IsClosed = False Then oSqlDr.Close()
            If Dt.Rows.Count > 0 Then Dt.Dispose()
        End Try

    End Sub
    Private Sub GenerateTextFile_Mahindra()
        Dim strsql As String
        Dim strStartDate As String = String.Empty
        Dim mFILE_NAME As String
        Dim Dt As DataTable = Nothing
        Dim Dt1 As DataTable = Nothing
        Dim INTCOUNTER As Integer
        Dim INTCOUNTER1 As Integer
        Dim obj_FSO As Scripting.FileSystemObject
        Dim strRecords As String = String.Empty
        Dim strAsnPath As String = String.Empty
        Dim strPath As String = String.Empty
        Dim strmahindraackn As String = String.Empty

        'objwrite.AutoFlush = True

        strStartDate = CDate(SqlConnectionclass.ExecuteScalar("SELECT MAHINDRA_ASN_STARTDATE from sales_parameter where unit_code='" & gstrUNITID & "'"))
        strAsnPath = SqlConnectionclass.ExecuteScalar("SELECT PATH from MAHINDRA_ASN_CONFIGURATION where unit_code='" & gstrUNITID & "' and TYPE='TXT'")

        Try
            If Directory.Exists(strAsnPath) = False Then
                Directory.CreateDirectory(strAsnPath)
            End If

            If txtFromInvoice.Text.Length > 0 And txtToInvoice.Text.Length > 0 Then
                strsql = "SET DATEFORMAT 'DMY' SELECT SC.DOC_NO FROM SALESCHALLAN_DTL SC WHERE SC.UNIT_CODE='" & gstrUNITID & "' AND SC.ACCOUNT_CODE='" & txtCustomerCode.Text & "'" & _
                    " AND SC.DOC_NO BETWEEN " & txtFromInvoice.Text & " And " & txtToInvoice.Text & "and SC.invoice_date >='" & strStartDate & "'" & _
                    " AND SC.BILL_FLAG=1 AND SC.CANCEL_FLAG=0 " & _
                    " AND NOT EXISTS " & _
                    " ( SELECT TOP 1 1 FROM MAHINDRA_ASN_BARCODE_ACKNOWLEDGMENT MM WHERE MM.UNIT_CODE=SC.UNIT_CODE AND MM.INVOICE_NO=SC.DOC_NO )"
                Dt = SqlConnectionclass.GetDataTable(strsql)

            End If

            If Dt.Rows.Count > 0 Then
                obj_FSO = New Scripting.FileSystemObject

                For INTCOUNTER = 0 To Dt.Rows.Count - 1
                    strPath = ""
                    If gstrUNITID = "SMA" Or gstrUNITID = "SMC" Or gstrUNITID = "SMK" Or gstrUNITID = "SMP" Then
                        strPath = strAsnPath + "INVOIC_SMRC_MAHI_" + Dt.Rows(INTCOUNTER).Item("DOC_NO").ToString + ".txt"
                    Else
                        strPath = strAsnPath + "INVOIC_MATE_MAHI_" + Dt.Rows(INTCOUNTER).Item("DOC_NO").ToString + ".txt"
                    End If

                    FileOpen(1, strPath, OpenMode.Append)

                    strsql = "SELECT * FROM VW_MAHINDRA_ASNDATA WHERE unit_code='" & gstrUNITID & "' and  INVOICENUMBER = " & Dt.Rows(INTCOUNTER).Item("DOC_NO") & ""
                    Dt1 = SqlConnectionclass.GetDataTable(strsql)

                    If Dt1.Rows.Count > 0 Then
                        For INTCOUNTER1 = 0 To Dt1.Rows.Count - 1
                            strRecords = ""
                            If INTCOUNTER1 = 0 Then

                                strRecords = Dt1.Rows(INTCOUNTER1).Item("HEADER").ToString.Trim + "," + Dt1.Rows(INTCOUNTER1).Item("CUSTOMERCODE").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("CUSTVENDOR").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("GSTINVOICE").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("INVOICENUMBER").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("INVOICEDATE").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("GSTNO").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("INVOICEDATE").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("PONUMBER").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("PODATE").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("IRN_NO").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("QRCODE").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("ORIGINALINVOICENUMBER").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("OCINVOICETYPE").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("ASNNO").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("GOODSREMOVALDATETIME").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("CONSIGNMENTNO").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("SOTO").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("SOVATREGTNO").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("SOLDGSTNNO").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("BILLTOID").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("PAYEEGSTNNO").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("SUPPLIERCODE").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("SELLERGSTIN").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("CURRENCY").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("MODEOFTRANSPORT").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("TRANSPORTERNAME").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("DESINATION").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("INVOICEAMOUNT").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("TAXABLEAMOUNT").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("TAXABLESUBAMOUNT").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("TOTALLINEITEMAMOUNT").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("TAMOUNTPAYABLE").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("TAAMOUNT").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("CGSTAMOUNT").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("SGSTAMOUNT").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("TCSPERCENTAGE").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("TCSTAXAMOUNT").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("VEHICLENUMBER").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("LRNUMBER").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("LRDATE").ToString.Trim + ","
                                'strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("TCSTAXAMOUNT").ToString.Trim + ","
                                strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("IGSTAMOUNT").ToString.Trim + ","

                                strmahindraackn = "insert into MAHINDRA_ASN_BARCODE_ACKNOWLEDGMENT(UNIT_CODE,CUSTOMER_CODE,INVOICE_NO,IS_ASN_ACK)"
                                strmahindraackn += " select '" + gstrUNITID + "','" + txtCustomerCode.Text & "','" + Dt1.Rows(INTCOUNTER1).Item("INVOICENUMBER").ToString.Trim + "',0"
                                SqlConnectionclass.ExecuteNonQuery(strmahindraackn)

                            End If
                            strRecords = strRecords + vbCrLf
                            strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("DETAIL").ToString.Trim + ","
                            strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("BUYERSITEMNUMBER").ToString.Trim + ","
                            strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("SUPPLIERITEMNUMBER").ToString.Trim + ","
                            strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("INVOICEQUANTITY").ToString.Trim + ","
                            strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("CALCULATIONNET").ToString.Trim + ","
                            strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("CALCULATIONGROSS").ToString.Trim + ","
                            strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("VATPERCENTAGE").ToString.Trim + ","
                            strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("VATAMOUNT").ToString.Trim + ","
                            strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("CGS_PERCENTAGE").ToString.Trim + ","
                            strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("CGS_AMOUNT").ToString.Trim + ","
                            strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("SGS_PERCENTAGE").ToString.Trim + ","
                            strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("SGS_AMOUNT").ToString.Trim + ","
                            strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("EXCISE_PERCENTAGE").ToString.Trim + ","
                            strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("EXCISE_AMOUNT").ToString.Trim + ","
                            strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("NETPRICE").ToString.Trim + ","
                            strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("IGS_PERCENTAGE").ToString.Trim + ","
                            strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("IGS_AMOUNT").ToString.Trim + ","

                            'strRecords = strRecords + Dt1.Rows(INTCOUNTER1).Item("INVOICEAMOUNT").ToString.Trim + ","


                            PrintLine(1, strRecords)
                        Next
                        FileClose(1)
                    End If

                Next
                MsgBox("ASN TEXT File(s) Generated Successfully !! ", MsgBoxStyle.Information, ResolveResString(100))
                Call ClearAllFields()
            Else
                MsgBox("No Pending invoice for ASN TEXT File !! ", MsgBoxStyle.Information, ResolveResString(100))
                ClearAllFields()
            End If

            'objwrite = New System.IO.StreamWriter(mFILE_NAME)

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub ClearAllFields()
        Try
            txtFromInvoice.Text = String.Empty
            txtToInvoice.Text = String.Empty
            txtCustomerCode.Text = String.Empty


        Catch ex As Exception

        End Try
    End Sub

    Private Function Validate_Data() As Boolean

        Dim strValidate As String = String.Empty

        Try

            If txtCustomerCode.Text.Trim.Length <= 0 Then
                strValidate = "Customer Code cannot be Blank."
                txtCustomerCode.Focus()
            End If

            If String.IsNullOrEmpty(cmbInvoiceType.Text) Then
                strValidate = strValidate + Chr(13) & "Invoice Type cannot be blank."
                cmbInvoiceType.Focus()
            End If

            If txtFromInvoice.Text.Trim.Length <= 0 Then
                strValidate = strValidate + Chr(13) & "From Invoice cannot be Blank."
                txtFromInvoice.Focus()
            End If

            If txtToInvoice.Text.Trim.Length <= 0 Then
                strValidate = strValidate + Chr(13) & "To Invoice cannot be Blank."
                txtToInvoice.Focus()
            End If

            If dtFromDate.Value > dtToDate.Value Then
                strValidate = strValidate + Chr(13) & "[From date] should be less than or equal to [To date]."
                dtFromDate.Focus()
            End If

            If Val(txtFromInvoice.Text) > Val(txtToInvoice.Text) Then
                strValidate = strValidate + Chr(13) & "[From Invoice] should be less than or equal to [To Invoice]."
                txtFromInvoice.Focus()
            End If

            If strValidate <> String.Empty Then
                Validate_Data = False
                MsgBox(strValidate, MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            Else
                Validate_Data = True
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Function

    Private Function FN_Get_Folder_Path() As String

        Dim strFilePath As String = String.Empty
        Dim strReturnValue As String = String.Empty

        Try
            strSql = " SELECT ISNULL(ASN_HMIL_FilePath,'')as ASN_HMIL_FilePath" & _
                     " FROM Sales_Parameter (nolock)" & _
                     " WHERE UNIT_CODE = '" & gstrUNITID & "'"
            strFilePath = SqlConnectionclass.ExecuteScalar(strSql)
            If strFilePath <> String.Empty Then
                strReturnValue = strFilePath
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

        FN_Get_Folder_Path = strReturnValue

    End Function
    Private Function FN_Get_CYGNET_Folder_Path() As String

        Dim strFilePath As String = String.Empty
        Dim strReturnValue As String = String.Empty

        Try
            strSql = " SELECT ISNULL(TML_CYGNET_PATH,'')as TML_CYGNET_PATH " & _
                     " FROM Sales_Parameter (nolock)" & _
                     " WHERE UNIT_CODE = '" & gstrUNITID & "'"
            strFilePath = SqlConnectionclass.ExecuteScalar(strSql)
            If strFilePath <> String.Empty Then
                strReturnValue = strFilePath
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

        FN_Get_CYGNET_Folder_Path = strReturnValue

    End Function
    Private Sub AddInvoiceTypeInCombo()
        Dim dtTemp As New DataTable
        Try
            cmbInvoiceType.Items.Clear()
            dtTemp.Columns.Add("Text", GetType(System.String))
            dtTemp.Columns.Add("Value", GetType(System.String))
            Dim dr As DataRow

            dr = dtTemp.NewRow()
            dr("Text") = "NORMAL - FINISHED GOODS"
            dr("Value") = "F"
            dtTemp.Rows.Add(dr)

            dr = dtTemp.NewRow()
            dr("Text") = "NORMAL - COMPONENTS"
            dr("Value") = "C"
            dtTemp.Rows.Add(dr)

            dr = dtTemp.NewRow()
            dr("Text") = "Export"
            dr("Value") = "E"
            dtTemp.Rows.Add(dr)

            cmbInvoiceType.DataSource = dtTemp
            cmbInvoiceType.ValueMember = "Value"
            cmbInvoiceType.DisplayMember = "Text"
            cmbInvoiceType.SelectedIndex = 0

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Function generateASN_MAE()
        Dim objExportInvASN As clsExportInvASN = New clsExportInvASN(gstrUNITID)
        Dim mInvNo As Double
        Dim Dt As New DataTable
        Dim strSql As String

        Try

            strSql = " SELECT Doc_No as Invoice_No " & _
                 " FROM SALESCHALLAN_DTL " & _
                 " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                 " AND Doc_No BETWEEN '" & txtFromInvoice.Text & "' AND '" & txtToInvoice.Text & "' " & _
                 " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 AND Invoice_Type in ('EXP') AND Sub_Category='" & Convert.ToString(cmbInvoiceType.SelectedValue) & "'" & _
                 " AND Account_Code='" & txtCustomerCode.Text.Trim & "' " & _
                 " AND UNIT_CODE = '" & gstrUNITID & "' "


            Dt = SqlConnectionclass.GetDataTable(strSql)

            If Dt.Rows.Count = 0 Then
                MessageBox.Show("No Data Found For File Generation.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Function
            Else
                For Each row As DataRow In Dt.Rows
                    mInvNo = Convert.ToDouble(row.Item(0).ToString())
                    objExportInvASN.WriteXML(mInvNo)
                Next
            End If

            MessageBox.Show("ASN files generated in " + gstrLocalCDrive.ToUpper() & "XML Files\")

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Generate ASN MAE", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Private Function IsTEXTFile_Mahindra(ByVal strcustomercode As String)
        If SqlConnectionclass.ExecuteScalar("SELECT TOP 1 1 FROM Customer_mst where unit_code='" & gstrUNITID & "' and customer_code='" & strcustomercode & "' and MAHINDRA_ASN_ENABLED=1") Then
            Return True
        Else
            Return False
        End If

    End Function
#End Region

#Region "COMMAND BUTTONS"

    Private Sub btnGenerateFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenerateFile.Click

        Try

            If (gstrUNITID = "MAE" Or gstrUNITID = "MA3") And cmbInvoiceType.SelectedValue = "E" Then
                generateASN_MAE()
            ElseIf SqlConnectionclass.ExecuteScalar("SELECT TOP 1 1 FROM Customer_mst where unit_code='" & gstrUNITID & "' and customer_code='" & txtCustomerCode.Text.Trim & "' and TML_ASN_ENABLED=1") Then
                GenerateASNFile_TML()
            Else
                GenerateASNFile()
            End If

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

#End Region

    Private Sub cmbInvoiceType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbInvoiceType.SelectedIndexChanged
        Try
            txtFromInvoice.Text = String.Empty
            txtToInvoice.Text = String.Empty
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub


    Private Sub btnGeneratetxtFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGeneratetxtFile.Click

        Try

            If IsTEXTFile_Mahindra(txtCustomerCode.Text) = True And txtFromInvoice.Text.Length > 0 And txtToInvoice.Text.Length > 0 Then
                GenerateTextFile_Mahindra()
                Exit Sub
            Else
                MsgBox("Auto DX is not associated with this customer code : " + txtCustomerCode.Text, MsgBoxStyle.Information, ResolveResString(100))

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Generate TEXT FILE ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub optFileType_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optFileType.CheckedChanged
        Try
            If optFileType.Checked = True Then
                btnSendASNViaAPI.Enabled = True
                btnGenerateFile.Enabled = True
                btnGeneratetxtFile.Enabled = False
            Else
                btnSendASNViaAPI.Enabled = False
                btnGenerateFile.Enabled = False
                btnGeneratetxtFile.Enabled = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub optTextfileType_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optTextfileType.CheckedChanged
        Try
            If optTextfileType.Checked = True Then
                btnGenerateFile.Enabled = False
                btnGeneratetxtFile.Enabled = True
            Else
                btnGenerateFile.Enabled = True
                btnGeneratetxtFile.Enabled = False
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnSendASNViaAPI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSendASNViaAPI.Click
        Dim ObjFRMMKTTRN0083A As FRMMKTTRN0083A
        Dim dt As DataTable = New DataTable("InvoiceData")
        Dim strSql As String = String.Empty, strQuery As String = String.Empty, strFinalQry As String = String.Empty, StrProcessID As String = String.Empty
        Try

            If Validate_Data() = True Then
                strSql = " SELECT Query " & _
                             " FROM Formate_Mst A " & _
                             " INNER JOIN Formate_Customer_Linkage B On B.Formate_ID =A.Formate_ID " & _
                             " AND B.Unit_Code =A.Unit_Code " & _
                             " WHERE Customer_Code='" & txtCustomerCode.Text.Trim & "' And A.IsActive=1 " & _
                             " AND A.UNIT_CODE='" & gstrUNITID & "'"
                strQuery = SqlConnectionclass.ExecuteScalar(strSql)
                If strQuery <> Nothing Or strQuery <> String.Empty Then
                    strSql = " SELECT Doc_No as Invoice_No " & _
                             " FROM SALESCHALLAN_DTL " & _
                             " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtFromDate.Text & "',103) AND Convert(Date,'" & dtToDate.Text & "',103)" & _
                             " AND Doc_No BETWEEN '" & txtFromInvoice.Text & "' AND '" & txtToInvoice.Text & "' " & _
                             " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 AND Invoice_Type in ('INV','EXP') AND Sub_Category='" & Convert.ToString(cmbInvoiceType.SelectedValue) & "'" & _
                             " AND Account_Code='" & txtCustomerCode.Text.Trim & "' " & _
                             " AND UNIT_CODE = '" & gstrUNITID & "' " & _
                             " "
                    'Code to get ProcessID using system date sql function
                    StrProcessID = SqlConnectionclass.ExecuteScalar("Select REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(SYSDATETIME(),'-',''),':',''),' ',''),'.',''),'  ','') ProcessID")

                    strFinalQry = "Select PROCESSID = '" & StrProcessID & "', InvoiceNo = [Vendor Challan No], Unit_Code = '" + gstrUNITID + "',InvoiceDate =  CHN.Invoice_Date,Customer_Code= CHN.Account_Code , Ent_Dt = Getdate(), Ent_UserId = '" + mP_User + "'" & vbCrLf & "From (" + strQuery & vbCrLf & " WHERE X.Doc_No in  (" & strSql & ") " + " )St " & vbCrLf
                    strFinalQry = strFinalQry & " Inner Join SALESCHALLAN_DTL CHN on CHN.DOC_NO = st.[Vendor Challan No] And CHN.UNIT_CODE = '" + gstrUNITID + "'"
                    strFinalQry = strFinalQry + " Left Join TML_TCS_ASN_WebAPI_Integration_Data TML On TML.InvoiceNo = St.[Vendor Challan No] And TML.Unit_Code = '" + gstrUNITID + "' And Coalesce(TMl.APIStatus,'Success') = 'Success' " & vbCrLf
                    strFinalQry = strFinalQry + " Where TML.InvoiceNo Is Null And Coalesce(TML.IsProcessed,1) = 1 "
                    If dt.Rows.Count > 0 Then dt.Clear()
                    dt = SqlConnectionclass.GetDataTable(strFinalQry)

                    If (dt.Rows.Count > 10) Then
                        MessageBox.Show("Maximum 10 Invoice are allowed in API.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        'ElseIf (dt.Rows.Count <= 0) Then
                        '    MessageBox.Show("No record found.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Else
                        strSql = "insert into TML_TCS_ASN_WebAPI_Integration_Data(PROCESSID,InvoiceNo, Unit_Code, InvoiceDate, Customer_Code, Ent_Dt, Ent_UserId) " + strFinalQry
                        SqlConnectionclass.ExecuteNonQuery(strSql)
                        ASN_Integration_EXE_Calling(StrProcessID)
                        ObjFRMMKTTRN0083A = New FRMMKTTRN0083A()
                        ObjFRMMKTTRN0083A.ShowDialog()
                    End If
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub


    Private Sub ASN_Integration_EXE_Calling(ByVal StrProcessID As String)
        Try
            'Dim Hilex_TML_TCS_ASN_GenerationIntegration_Path As String = ConfigurationManager.AppSettings("Hilex_TML_TCS_ASN_GenerationIntegration_Path")
            Dim Hilex_TML_TCS_ASN_GenerationIntegration_Path As String = Application.StartupPath + "\TML_TCS_ASNGenerationAPI_IntegrationAppilcation\TML_TCS_ASNGenerationAPI_Integration.exe"
            Dim pHelp As New ProcessStartInfo
            pHelp.FileName = Hilex_TML_TCS_ASN_GenerationIntegration_Path
            pHelp.Arguments = """" & gstrUNITID & """" & " """ & StrProcessID & """"
            pHelp.UseShellExecute = True
            pHelp.WindowStyle = ProcessWindowStyle.Hidden
            Dim proc As Process = Process.Start(pHelp)
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub btnexception_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnexception.Click
        Dim oData As New DataSet
        Dim oAdpt As SqlDataAdapter
        Dim cnt As Long
        Dim strsql As String
        Dim fulldata As String = String.Empty
        Try

        
            If txtCustomerCode.Text = "" Then
                MsgBox("PLEASE DEFINE CUSTOMER CODE FIRST  " + txtCustomerCode.Text, MsgBoxStyle.Information, ResolveResString(100))
                RichTextBox1.Text = ""
                Exit Sub
            End If

            strsql = "SELECT CAST(INVOICE_NO AS VARCHAR(10))+ ':'  + ERROR_DESCRIPTION INVOICE_DESC FROM VW_TATACYGNET_EXCEPTION WHERE CUSTOMER_CODE ='" & txtCustomerCode.Text & "'"
            oAdpt = New SqlDataAdapter(strsql, SqlConnectionclass.GetConnection)
            oAdpt.Fill(oData, "tempTable")
            RichTextBox1.Visible = False

            If oData.Tables(0).Rows.Count > 0 Then
                RichTextBox1.Visible = True
                For cnt = 1 To oData.Tables(0).Rows.Count
                    fulldata = fulldata + oData.Tables(0).Rows(cnt - 1).Item("INVOICE_DESC").ToString.Trim() + " " + Chr(13) + Chr(13)
                Next
                RichTextBox1.Text = fulldata
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox1.TextChanged

    End Sub
End Class