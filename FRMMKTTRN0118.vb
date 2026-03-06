Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
'*********************************************************************************************************************
'Copyright(c)       - MIND
'Name of Module     - BSR Packing List Generation
'Name of Form       - FRMMKTTRN0118  , BSR Packing List Generation
'Created by         - Ashish sharma
'Created Date       - 11 OCT 2021
'description        - BSR Packing List Generation and Print
'*********************************************************************************************************************
Public Class FRMMKTTRN0118
    Private Enum GridInvoiceDetails
        btnDelete = 0
        DocNo
        DocDate
    End Enum

    Private Sub FRMMKTTRN0118_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            txtScanInvoice.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FRMMKTTRN0118_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Call FitToClient(Me, GrpMain, ctlHeader, GrpBoxButtons, 500)
            Me.MdiParent = mdifrmMain
            ConfigureGridColumnns()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub ConfigureGridColumnns()
        Try
            dgvInvoiceDetail.Columns.Clear()
            Dim objButton As New DataGridViewButtonColumn
            objButton.Name = "Delete"
            objButton.HeaderText = " "

            dgvInvoiceDetail.Columns.Add(objButton)
            dgvInvoiceDetail.Columns.Add("DocNo", "Doc. No.")
            dgvInvoiceDetail.Columns.Add("DocDate", "Doc. Date")

            dgvInvoiceDetail.Columns(GridInvoiceDetails.btnDelete).Width = 70
            dgvInvoiceDetail.Columns(GridInvoiceDetails.DocNo).Width = 100
            dgvInvoiceDetail.Columns(GridInvoiceDetails.DocDate).Width = 100

            dgvInvoiceDetail.Columns(GridInvoiceDetails.btnDelete).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvInvoiceDetail.Columns(GridInvoiceDetails.DocNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvInvoiceDetail.Columns(GridInvoiceDetails.DocDate).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            dgvInvoiceDetail.Columns(GridInvoiceDetails.DocNo).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetails.DocDate).ReadOnly = True
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

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            dgvInvoiceDetail.Rows.Clear()
            lblTotalScanInvoices.Text = "0"
            txtScanInvoice.Text = String.Empty
            lblCustomerCode.Text = String.Empty
            lblCustCodeDes.Text = String.Empty
            txtScanInvoice.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtScanInvoice_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtScanInvoice.KeyDown
        Dim dtInvoiceDtl As New DataTable
        Try
            If e.KeyCode = Keys.Enter Then
                Dim invoiceNo As String = String.Empty
                invoiceNo = txtScanInvoice.Text.Trim()
                txtScanInvoice.Text = String.Empty

                If (String.IsNullOrEmpty(invoiceNo)) Then
                    MsgBox("Please scan a valid Invoice.", MsgBoxStyle.Exclamation, ResolveResString(100))
                    txtScanInvoice.Focus()
                    Exit Sub
                End If
                If (invoiceNo.Length < 6) Then
                    MsgBox("Please scan a valid Invoice.", MsgBoxStyle.Exclamation, ResolveResString(100))
                    txtScanInvoice.Focus()
                    Exit Sub
                End If
                If (IsNumeric(invoiceNo) = False) Then
                    MsgBox("Please scan a valid Invoice.", MsgBoxStyle.Exclamation, ResolveResString(100))
                    txtScanInvoice.Focus()
                    Exit Sub
                End If

                If (CheckDuplicateInvoice(invoiceNo)) Then
                    txtScanInvoice.Focus()
                    Exit Sub
                End If

                Using sqlcmd As SqlCommand = New SqlCommand
                    With sqlcmd
                        .CommandText = "USP_BSR_PACKING_LIST"
                        .CommandTimeout = 0
                        .CommandType = CommandType.StoredProcedure
                        .Parameters.Clear()
                        .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUnitId
                        .Parameters.Add("@INVOICE_NO", SqlDbType.VarChar, 18).Value = invoiceNo
                        .Parameters.Add("@USER_ID", SqlDbType.VarChar, 50).Value = mP_User
                        .Parameters.Add("@OPERATION", SqlDbType.VarChar, 50).Value = "SCAN_INVOICE"
                        .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                        dtInvoiceDtl = SqlConnectionclass.GetDataTable(sqlcmd)
                        If Convert.ToString(.Parameters("@MESSAGE").Value) = String.Empty Then
                            If dtInvoiceDtl IsNot Nothing AndAlso dtInvoiceDtl.Rows.Count > 0 Then
                                If dgvInvoiceDetail.Rows.Count = 0 Then
                                    lblCustomerCode.Text = Convert.ToString(dtInvoiceDtl.Rows(0)("ACCOUNT_CODE"))
                                    lblCustCodeDes.Text = Convert.ToString(dtInvoiceDtl.Rows(0)("ACCOUNT_NAME"))
                                End If
                                AddRowInGrid(dtInvoiceDtl)
                                txtScanInvoice.Focus()
                            Else
                                MsgBox("Please scan a valid Invoice.", MsgBoxStyle.Exclamation, ResolveResString(100))
                                txtScanInvoice.Focus()
                            End If
                        Else
                            MsgBox(Convert.ToString(.Parameters("@MESSAGE").Value), MsgBoxStyle.Exclamation, ResolveResString(100))
                            txtScanInvoice.Focus()
                        End If
                    End With
                End Using
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            If dtInvoiceDtl IsNot Nothing Then
                dtInvoiceDtl.Dispose()
            End If
        End Try
    End Sub

    Private Function CheckDuplicateInvoice(ByVal invoiceNo As String) As Boolean
        Dim result As Boolean = False
        If dgvInvoiceDetail IsNot Nothing AndAlso dgvInvoiceDetail.Rows IsNot Nothing AndAlso dgvInvoiceDetail.Rows.Count > 0 Then
            For i As Integer = 0 To dgvInvoiceDetail.Rows.Count - 1
                If Convert.ToString(dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.DocNo).Value) = invoiceNo Then
                    MsgBox("Scan invoice is duplicate.", MsgBoxStyle.Exclamation, ResolveResString(100))
                    result = True
                    Exit For
                End If
            Next
        End If
        Return result
    End Function

    Private Sub dgvInvoiceDetail_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvInvoiceDetail.CellClick
        Try
            If e.RowIndex < 0 Then Exit Sub
            If e.ColumnIndex = GridInvoiceDetails.btnDelete Then
                dgvInvoiceDetail.Rows.RemoveAt(e.RowIndex)
                UpdateTotalInvoices("DELETE", 1)
                If dgvInvoiceDetail.Rows.Count = 0 Then
                    txtScanInvoice.Text = String.Empty
                    lblCustomerCode.Text = String.Empty
                    lblCustCodeDes.Text = String.Empty
                    lblTotalScanInvoices.Text = "0"
                    txtScanInvoice.Focus()
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub UpdateTotalInvoices(ByVal mode As String, ByVal qty As Integer)
        If String.IsNullOrEmpty(lblTotalScanInvoices.Text) Then lblTotalScanInvoices.Text = "0"
        If mode = "ADD" Then
            lblTotalScanInvoices.Text = Convert.ToString(Convert.ToInt32(lblTotalScanInvoices.Text) + qty)
        Else
            lblTotalScanInvoices.Text = Convert.ToString(Convert.ToInt32(lblTotalScanInvoices.Text) - qty)
        End If
        If String.IsNullOrEmpty(lblTotalScanInvoices.Text) Then lblTotalScanInvoices.Text = "0"
    End Sub
    Private Sub AddRowInGrid(ByVal dt As DataTable)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            If dgvInvoiceDetail.Rows.Count = 0 Then
                dgvInvoiceDetail.Rows.Add("Delete", Convert.ToString(dt.Rows(0)("DOC_NO")), Convert.ToString(dt.Rows(0)("INVOICE_DATE")))
                UpdateTotalInvoices("ADD", 1)
                Exit Sub
            End If
            If Not CheckCustomer(Convert.ToString(dt.Rows(0)("ACCOUNT_CODE"))) Then
                dgvInvoiceDetail.Rows.Add("Delete", Convert.ToString(dt.Rows(0)("DOC_NO")), Convert.ToString(dt.Rows(0)("INVOICE_DATE")))
                UpdateTotalInvoices("ADD", 1)
            End If
        End If
    End Sub
    Private Function CheckCustomer(ByVal strCustomerCode As String) As Boolean
        Dim result As Boolean = False

        If Convert.ToString(lblCustomerCode.Text.Trim().ToUpper()) <> strCustomerCode.Trim().ToUpper() Then
            MsgBox("Scanned invoice customer should be same with previous scanned invoice customer.", MsgBoxStyle.Exclamation, ResolveResString(100))
            txtScanInvoice.Text = String.Empty
            txtScanInvoice.Focus()
            result = True
        End If
        Return result
    End Function

    Private Sub btnGeneratePackingList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGeneratePackingList.Click
        Dim dtInvoices As New DataTable
        Try
            If String.IsNullOrEmpty(lblCustomerCode.Text) Then
                MsgBox("[Customer Code] can't be blank.Please Scan invoice.", MsgBoxStyle.Exclamation, ResolveResString(100))
                txtScanInvoice.Text = String.Empty
                txtScanInvoice.Focus()
                Exit Sub
            End If
            If dgvInvoiceDetail Is Nothing OrElse dgvInvoiceDetail.Rows.Count = 0 Then
                MsgBox("Please scan invoice.", MsgBoxStyle.Exclamation, ResolveResString(100))
                txtScanInvoice.Text = String.Empty
                txtScanInvoice.Focus()
                Exit Sub
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
            dtInvoices.Columns.Add("INVOICE_NO", GetType(System.Int64))
            Dim drSave As DataRow
            For i As Integer = 0 To dgvInvoiceDetail.Rows.Count - 1
                drSave = dtInvoices.NewRow()
                drSave("INVOICE_NO") = dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetails.DocNo).Value
                dtInvoices.Rows.Add(drSave)
            Next
            Using sqlcmd As SqlCommand = New SqlCommand
                With sqlcmd
                    .CommandText = "USP_BSR_PACKING_LIST"
                    .CommandTimeout = 0
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Clear()
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUnitId
                    .Parameters.Add("@UDT_BSR_PACKING_LIST_INVOICENO", SqlDbType.Structured).Value = dtInvoices
                    .Parameters.Add("@CUSTOMER_CODE", SqlDbType.VarChar, 8).Value = Convert.ToString(lblCustomerCode.Text)
                    .Parameters.Add("@USER_ID", SqlDbType.VarChar, 50).Value = mP_User
                    .Parameters.Add("@OPERATION", SqlDbType.VarChar, 50).Value = "GENERATE_PACKINGLIST"
                    .Parameters.Add("@RET_PACKING_LIST_NO", SqlDbType.VarChar, 18).Direction = ParameterDirection.Output
                    .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                    SqlConnectionclass.ExecuteNonQuery(sqlcmd)
                    If Convert.ToString(.Parameters("@MESSAGE").Value) = String.Empty Then
                        Dim packingListNo As String = Convert.ToString(.Parameters("@RET_PACKING_LIST_NO").Value)
                        If Val(packingListNo) = 0 Then
                            MsgBox("Error to generate packing list no.", MsgBoxStyle.Exclamation, ResolveResString(100))
                            txtScanInvoice.Focus()
                            Exit Sub
                        Else
                            PrintPackingList(packingListNo, "O")
                            MsgBox("Packing List No. : " & packingListNo & " has been successfully generated.", MsgBoxStyle.Information, ResolveResString(100))
                            btnCancel.PerformClick()
                        End If
                    Else
                        MsgBox(Convert.ToString(.Parameters("@MESSAGE").Value), MsgBoxStyle.Exclamation, ResolveResString(100))
                        txtScanInvoice.Focus()
                    End If
                End With
            End Using
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
            If dtInvoices IsNot Nothing Then
                dtInvoices.Dispose()
            End If
        End Try
    End Sub

    Private Sub PrintPackingList(ByVal packingListNo As String, ByVal copyName As String)
        Try
            Dim Address As String = gstr_WRK_ADDRESS1 & gstr_WRK_ADDRESS2
            Dim objRpt As ReportDocument
            Dim frmReportViewer As New eMProCrystalReportViewer
            objRpt = frmReportViewer.GetReportDocument()
            objRpt.Load(My.Application.Info.DirectoryPath & "\Reports\rptBSRPackingList.rpt")
            objRpt.DataDefinition.FormulaFields("CompanyName").Text = "'" + gstrCOMPANY + "'"
            objRpt.DataDefinition.FormulaFields("CompanyAddress").Text = "'" + Address + "'"
            objRpt.DataDefinition.FormulaFields("COPY_NAME").Text = "'" + copyName + "'"
            objRpt.RecordSelectionFormula = "{BSR_PACKING_LIST_DTL.PACKING_LIST_NO}=" & packingListNo & " AND {BSR_PACKING_LIST_DTL.UNIT_CODE}='" & gstrUnitId & "'"
            frmReportViewer.ShowPrintButton = False
            frmReportViewer.ShowTextSearchButton = False
            frmReportViewer.ShowZoomButton = False
            frmReportViewer.SetReportDocument()
            objRpt.PrintToPrinter(1, False, 0, 0)
            'frmReportViewer.Show()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtScanInvoice_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtScanInvoice.KeyPress
        Try
            If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
                e.Handled = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
End Class