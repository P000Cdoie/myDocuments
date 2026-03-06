Imports System.Data
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
'*********************************************************************************************************************
'Copyright(c)       - MIND
'Name of Module     - Bulk Invoice Printing
'Name of Form       - FRMMKTTRN0114  , Bulk Invoice Printing
'Created by         - Ashish sharma
'Created Date       - 08 MAR 2021
'description        - Bulk invoice / RGP / NRGP Printing
'*********************************************************************************************************************
Public Class FRMMKTTRN0114
    Private Const fromLocation As String = "01P1"
    Dim isRGPNRGPMandatoryForInvoicePrint As Boolean = False
    Dim printCountRGPNRGP As Integer = 0
    Private Enum GridInvoiceSummary
        TotalInvoice = 0
        LockedInvoice
        TempInvoice
    End Enum
    Private Enum GridInvoiceDetail
        Selection = 0
        DocNo
        IRNNo
        EwayBillNo
    End Enum
    Private Enum GridRGPDetail
        CloseBoxNo = 0
        DocType
        DocNo
    End Enum
    Private Enum GridRGPNRGPDetail
        Selection = 0
        DocType
        DocNo
    End Enum
    Private Sub FRMMKTTRN0114_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            txtCustomerCode.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FRMMKTTRN0114_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000
        Try
            If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
                Call ctlHeader_Click(ctlHeader, New System.EventArgs())
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub FRMMKTTRN0114_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Call FitToClient(Me, GrpMain, ctlHeader, GrpBoxButtons, 550)
            Me.MdiParent = mdifrmMain
            dtpDateFrom.Value = GetServerDate()
            dtpDateTo.Value = GetServerDate()
            ConfigureGridColumn()
            isRGPNRGPMandatoryForInvoicePrint = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("SELECT ISNULL(IS_RGP_NRGP_MANDATORY_FOR_INV_PRINT,0) AS IS_RGP_NRGP_MANDATORY_FOR_INV_PRINT FROM BSR_APP_CONFIG_MST (NOLOCK) WHERE UNIT_CODE='" & gstrUnitId & "'"))
            printCountRGPNRGP = Convert.ToInt32(SqlConnectionclass.ExecuteScalar("SELECT ISNULL(RGP_NRGP_PRINT_COUNT,0) AS RGP_NRGP_PRINT_COUNT FROM BSR_APP_CONFIG_MST (NOLOCK) WHERE UNIT_CODE='" & gstrUnitId & "'"))
            If printCountRGPNRGP = 0 Then
                printCountRGPNRGP = 1
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub ctlHeader_Click(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles ctlHeader.Click
        Try
            Call ShowHelp("UNDERCONSTRUCTION.HTM")
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub CmdCustCodeHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdCustCodeHelp.Click
        Dim strHelp() As String = Nothing
        Dim strSql As String = String.Empty
        Try
            If dtpDateFrom.Value > dtpDateTo.Value Then
                MessageBox.Show("[Date From] should be less than or equal to [Date To].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                dtpDateFrom.Focus()
                Exit Sub
            Else
                strSql = "SELECT CUST_CODE,CUST_NAME FROM DBO.UDF_BSR_GET_CUSTOMERS_BULK_INVOICE_PRINTING('" & gstrUnitId & "','" & dtpDateFrom.Value.ToString("dd MMM yyyy") & "','" & dtpDateTo.Value.ToString("dd MMM yyyy") & "') ORDER BY CUST_CODE"
                strHelp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "Customer(s) Help")
                If Not (UBound(strHelp) <= 0) Then
                    If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                        MessageBox.Show("No record To Display", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        txtCustomerCode.Text = String.Empty
                        Exit Sub
                    Else
                        txtCustomerCode.Text = strHelp(0).Trim
                        lblCustCodeDes.Text = strHelp(1).Trim
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub cmdVehicleNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdVehicleNo.Click
        Dim strHelp() As String = Nothing
        Dim strSql As String = String.Empty
        Try
            If dtpDateFrom.Value > dtpDateTo.Value Then
                MessageBox.Show("[Date From] should be less than or equal to [Date To].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                dtpDateFrom.Focus()
                Exit Sub
            ElseIf String.IsNullOrEmpty(txtCustomerCode.Text) Then
                MessageBox.Show("Please first select [Customer Code].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtCustomerCode.Text = String.Empty
                lblCustCodeDes.Text = String.Empty
                txtCustomerCode.Focus()
                Exit Sub
            Else
                strSql = "SELECT VEHICLE_NO,TRANSPORTER_ID FROM DBO.UDF_BSR_GET_VEHICLE_BULK_INVOICE_PRINTING('" & gstrUnitId & "','" & Trim(txtCustomerCode.Text) & "','" & dtpDateFrom.Value.ToString("dd MMM yyyy") & "','" & dtpDateTo.Value.ToString("dd MMM yyyy") & "') ORDER BY VEHICLE_NO"
                strHelp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "Vehicle(s) Help")
                If Not (UBound(strHelp) <= 0) Then
                    If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                        MessageBox.Show("No record To Display", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        txtVehicleNo.Text = String.Empty
                        Exit Sub
                    Else
                        txtVehicleNo.Text = strHelp(0).Trim
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub cmdDocNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDocNo.Click
        Dim strHelp() As String = Nothing
        Dim strSql As String = String.Empty
        Try
            If dtpDateFrom.Value > dtpDateTo.Value Then
                MessageBox.Show("[Date From] should be less than or equal to [Date To].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                dtpDateFrom.Focus()
                Exit Sub
            ElseIf String.IsNullOrEmpty(txtCustomerCode.Text) Then
                MessageBox.Show("Please first select [Customer Code].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtCustomerCode.Text = String.Empty
                lblCustCodeDes.Text = String.Empty
                txtCustomerCode.Focus()
                Exit Sub
            ElseIf String.IsNullOrEmpty(txtVehicleNo.Text) Then
                MessageBox.Show("Please first select [Vehicle No].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtVehicleNo.Text = String.Empty
                txtVehicleNo.Focus()
                Exit Sub
            Else
                strSql = "SELECT DISP_LOT_NO,DISP_LOT_DT FROM DBO.UDF_BSR_GET_DISP_LOT_NO_BULK_INVOICE_PRINTING('" & gstrUnitId & "','" & Trim(txtCustomerCode.Text) & "','" & Trim(txtVehicleNo.Text) & "','" & dtpDateFrom.Value.ToString("dd MMM yyyy") & "','" & dtpDateTo.Value.ToString("dd MMM yyyy") & "') ORDER BY DISP_LOT_NO"
                strHelp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "DocNo(s) Help")
                If Not (UBound(strHelp) <= 0) Then
                    If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                        MessageBox.Show("No record To Display", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        txtDocNo.Text = String.Empty
                        Exit Sub
                    Else
                        txtDocNo.Text = strHelp(0).Trim
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Try
            If dtpDateFrom.Value > dtpDateTo.Value Then
                MessageBox.Show("[Date From] should be less than or equal to [Date To].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                dtpDateFrom.Focus()
                Exit Sub
            ElseIf String.IsNullOrEmpty(txtCustomerCode.Text) Then
                MessageBox.Show("Please first select [Customer Code].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtCustomerCode.Text = String.Empty
                lblCustCodeDes.Text = String.Empty
                txtCustomerCode.Focus()
                Exit Sub
            ElseIf String.IsNullOrEmpty(txtVehicleNo.Text) Then
                MessageBox.Show("Please first select [Vehicle No].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtVehicleNo.Text = String.Empty
                txtVehicleNo.Focus()
                Exit Sub
            ElseIf String.IsNullOrEmpty(txtDocNo.Text) Then
                MessageBox.Show("Please first select [Disp. Lot No].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtDocNo.Text = String.Empty
                txtDocNo.Focus()
                Exit Sub
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
            FillInvoices()
            rdbInvoiceCheckAll.Checked = False
            rdbInvoiceUncheckAll.Checked = False
            rdbRGPCheckAll.Checked = False
            rdbRGPUnCheckAll.Checked = False
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub txtCustomerCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Try
            If e.KeyCode = Keys.Delete Then
                txtCustomerCode.Text = String.Empty
                lblCustCodeDes.Text = String.Empty
            ElseIf e.KeyCode = Keys.F1 Then
                CmdCustCodeHelp_Click(sender, e)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtCustomerCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustomerCode.TextChanged
        Try
            txtVehicleNo.Text = String.Empty
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtVehicleNo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtVehicleNo.KeyDown
        Try
            If e.KeyCode = Keys.Delete Then
                txtVehicleNo.Text = String.Empty
            ElseIf e.KeyCode = Keys.F1 Then
                cmdVehicleNo_Click(sender, e)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtVehicleNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVehicleNo.TextChanged
        Try
            txtDocNo.Text = String.Empty
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtDocNo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDocNo.KeyDown
        Try
            If e.KeyCode = Keys.Delete Then
                txtDocNo.Text = String.Empty
            ElseIf e.KeyCode = Keys.F1 Then
                cmdDocNo_Click(sender, e)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtDocNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDocNo.TextChanged
        Try
            dgvInvoiceSummary.Rows.Clear()
            dgvInvoiceDetail.Rows.Clear()
            dgvRGPDetail.Rows.Clear()
            dgvRGPNRGPDetails.Rows.Clear()
            rdbInvoiceCheckAll.Checked = False
            rdbInvoiceUncheckAll.Checked = False
            rdbRGPCheckAll.Checked = False
            rdbRGPUnCheckAll.Checked = False
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub ConfigureGridColumn()
        Try
            dgvInvoiceSummary.Columns.Clear()

            dgvInvoiceSummary.Columns.Add("TotalInvoice", "Total Invoice(s)")
            dgvInvoiceSummary.Columns.Add("LockedInvoice", "Locked Invoice(s)")
            dgvInvoiceSummary.Columns.Add("TempInvoice", "Temp Invoice(s)")

            dgvInvoiceSummary.Columns(GridInvoiceSummary.TotalInvoice).Width = 120
            dgvInvoiceSummary.Columns(GridInvoiceSummary.LockedInvoice).Width = 120
            dgvInvoiceSummary.Columns(GridInvoiceSummary.TempInvoice).Width = 120

            'dgvInvoiceSummary.Columns(GridInvoiceSummary.TotalInvoice).HeaderCell.Style.Font = New Font(dgvInvoiceSummary.Font, FontStyle.Bold)
            'dgvInvoiceSummary.Columns(GridInvoiceSummary.LockedInvoice).HeaderCell.Style.Font = New Font(dgvInvoiceSummary.Font, FontStyle.Bold)
            'dgvInvoiceSummary.Columns(GridInvoiceSummary.TempInvoice).HeaderCell.Style.Font = New Font(dgvInvoiceSummary.Font, FontStyle.Bold)

            dgvInvoiceSummary.Columns(GridInvoiceSummary.TotalInvoice).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvInvoiceSummary.Columns(GridInvoiceSummary.LockedInvoice).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvInvoiceSummary.Columns(GridInvoiceSummary.TempInvoice).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvInvoiceSummary.Columns(GridInvoiceSummary.TotalInvoice).ReadOnly = True
            dgvInvoiceSummary.Columns(GridInvoiceSummary.LockedInvoice).ReadOnly = True
            dgvInvoiceSummary.Columns(GridInvoiceSummary.TempInvoice).ReadOnly = True


            dgvInvoiceDetail.Columns.Clear()

            Dim objChkBox As New DataGridViewCheckBoxColumn
            objChkBox.Name = "Selection"
            objChkBox.HeaderText = " "

            dgvInvoiceDetail.Columns.Add(objChkBox)
            dgvInvoiceDetail.Columns.Add("DocNo", "Invoice No.")
            dgvInvoiceDetail.Columns.Add("IRN", "IRN")
            dgvInvoiceDetail.Columns.Add("EwayBillNo", "EwayBillNo.")

            dgvInvoiceDetail.Columns(GridInvoiceDetail.Selection).Width = 35
            dgvInvoiceDetail.Columns(GridInvoiceDetail.DocNo).Width = 110
            dgvInvoiceDetail.Columns(GridInvoiceDetail.IRNNo).Width = 220
            dgvInvoiceDetail.Columns(GridInvoiceDetail.EwayBillNo).Width = 130


            'dgvInvoiceDetail.Columns(GridInvoiceDetail.Selection).HeaderCell.Style.Font = New Font(dgvInvoiceDetail.Font, FontStyle.Bold)
            'dgvInvoiceDetail.Columns(GridInvoiceDetail.DocNo).HeaderCell.Style.Font = New Font(dgvInvoiceDetail.Font, FontStyle.Bold)
            'dgvInvoiceDetail.Columns(GridInvoiceDetail.IRNNo).HeaderCell.Style.Font = New Font(dgvInvoiceDetail.Font, FontStyle.Bold)
            'dgvInvoiceDetail.Columns(GridInvoiceDetail.EwayBillNo).HeaderCell.Style.Font = New Font(dgvInvoiceDetail.Font, FontStyle.Bold)

            dgvInvoiceDetail.Columns(GridInvoiceDetail.Selection).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvInvoiceDetail.Columns(GridInvoiceDetail.DocNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvInvoiceDetail.Columns(GridInvoiceDetail.IRNNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvInvoiceDetail.Columns(GridInvoiceDetail.EwayBillNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            dgvInvoiceDetail.Columns(GridInvoiceDetail.DocNo).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetail.IRNNo).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetail.EwayBillNo).ReadOnly = True

            dgvInvoiceDetail.Columns(GridInvoiceDetail.Selection).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvInvoiceDetail.Columns(GridInvoiceDetail.DocNo).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvInvoiceDetail.Columns(GridInvoiceDetail.IRNNo).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvInvoiceDetail.Columns(GridInvoiceDetail.EwayBillNo).SortMode = DataGridViewColumnSortMode.NotSortable

            dgvRGPDetail.Columns.Clear()

            dgvRGPDetail.Columns.Add("CloseBoxNo", "CloseBox#")
            dgvRGPDetail.Columns.Add("DocType", "Doc Type")
            dgvRGPDetail.Columns.Add("DocNo", "RGP / NRGP No.")

            dgvRGPDetail.Columns(GridRGPDetail.CloseBoxNo).Width = 100
            dgvRGPDetail.Columns(GridRGPDetail.DocType).Width = 80
            dgvRGPDetail.Columns(GridRGPDetail.DocNo).Width = 120


            'dgvRGPDetail.Columns(GridRGPDetail.CloseBoxNo).HeaderCell.Style.Font = New Font(dgvRGPDetail.Font, FontStyle.Bold)
            'dgvRGPDetail.Columns(GridRGPDetail.DocType).HeaderCell.Style.Font = New Font(dgvRGPDetail.Font, FontStyle.Bold)
            'dgvRGPDetail.Columns(GridRGPDetail.DocNo).HeaderCell.Style.Font = New Font(dgvRGPDetail.Font, FontStyle.Bold)

            dgvRGPDetail.Columns(GridRGPDetail.CloseBoxNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvRGPDetail.Columns(GridRGPDetail.DocType).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvRGPDetail.Columns(GridRGPDetail.DocNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            dgvRGPDetail.Columns(GridRGPDetail.CloseBoxNo).ReadOnly = True
            dgvRGPDetail.Columns(GridRGPDetail.DocType).ReadOnly = True
            dgvRGPDetail.Columns(GridRGPDetail.DocNo).ReadOnly = True

            dgvRGPDetail.Columns(GridRGPDetail.CloseBoxNo).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvRGPDetail.Columns(GridRGPDetail.DocType).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvRGPDetail.Columns(GridRGPDetail.DocNo).SortMode = DataGridViewColumnSortMode.NotSortable


            dgvRGPNRGPDetails.Columns.Clear()

            objChkBox = New DataGridViewCheckBoxColumn
            objChkBox.Name = "Selection"
            objChkBox.HeaderText = " "

            dgvRGPNRGPDetails.Columns.Add(objChkBox)
            dgvRGPNRGPDetails.Columns.Add("DocType", "Doc Type")
            dgvRGPNRGPDetails.Columns.Add("DocNo", "RGP / NRGP No.")

            dgvRGPNRGPDetails.Columns(GridRGPNRGPDetail.Selection).Width = 35
            dgvRGPNRGPDetails.Columns(GridRGPNRGPDetail.DocType).Width = 80
            dgvRGPNRGPDetails.Columns(GridRGPNRGPDetail.DocNo).Width = 120

            dgvRGPNRGPDetails.Columns(GridRGPNRGPDetail.Selection).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvRGPNRGPDetails.Columns(GridRGPNRGPDetail.DocType).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvRGPNRGPDetails.Columns(GridRGPNRGPDetail.DocNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            dgvRGPNRGPDetails.Columns(GridRGPNRGPDetail.DocType).ReadOnly = True
            dgvRGPNRGPDetails.Columns(GridRGPNRGPDetail.DocNo).ReadOnly = True

            dgvRGPNRGPDetails.Columns(GridRGPNRGPDetail.Selection).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvRGPNRGPDetails.Columns(GridRGPNRGPDetail.DocType).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvRGPNRGPDetails.Columns(GridRGPNRGPDetail.DocNo).SortMode = DataGridViewColumnSortMode.NotSortable
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FillInvoices()
        Dim ds As New DataSet
        Dim sqlCmd As New SqlCommand
        Try
            Dim i As Integer = 0
            With sqlCmd
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 300 ' 5 Minute
                .CommandText = "USP_BSR_GET_BULK_INVOICE_PRINTING_DATA"
                .Parameters.Clear()
                .Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
                .Parameters.AddWithValue("@DISP_LOT_NO", Trim(txtDocNo.Text))
                .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                ds = SqlConnectionclass.GetDataSet(sqlCmd)
                If Convert.ToString(.Parameters("@MESSAGE").Value) <> "" Then
                    MsgBox(Convert.ToString(.Parameters("@MESSAGE").Value), MsgBoxStyle.Exclamation, ResolveResString(100))
                Else
                    If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                        dgvInvoiceSummary.Rows.Clear()
                        If ds.Tables(0).Rows.Count > 0 Then
                            dgvInvoiceSummary.Rows.Add(ds.Tables(0).Rows.Count)
                            For Each dr As DataRow In ds.Tables(0).Rows
                                dgvInvoiceSummary.Rows(i).Cells(GridInvoiceSummary.TotalInvoice).Value = dr("TOTAL_INVOICE")
                                dgvInvoiceSummary.Rows(i).Cells(GridInvoiceSummary.LockedInvoice).Value = dr("LOCKED_INVOICE")
                                dgvInvoiceSummary.Rows(i).Cells(GridInvoiceSummary.TempInvoice).Value = dr("TEMP_INVOICE")
                                i += 1
                            Next
                        End If
                        i = 0
                        dgvInvoiceDetail.Rows.Clear()
                        If ds.Tables(1).Rows.Count > 0 Then
                            dgvInvoiceDetail.Rows.Add(ds.Tables(1).Rows.Count)
                            For Each dr As DataRow In ds.Tables(1).Rows
                                dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.Selection).Value = False
                                dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.DocNo).Value = dr("DOC_NO")
                                dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.IRNNo).Value = dr("IRN_NO")
                                dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.EwayBillNo).Value = dr("EWAY_BILL_NO")
                                i += 1
                            Next
                        End If
                        i = 0
                        dgvRGPDetail.Rows.Clear()
                        If ds.Tables(2).Rows.Count > 0 Then
                            dgvRGPDetail.Rows.Add(ds.Tables(2).Rows.Count)
                            For Each dr As DataRow In ds.Tables(2).Rows
                                dgvRGPDetail.Rows(i).Cells(GridRGPDetail.CloseBoxNo).Value = dr("DOC_NO")
                                dgvRGPDetail.Rows(i).Cells(GridRGPDetail.DocType).Value = dr("DOC_TYPE")
                                dgvRGPDetail.Rows(i).Cells(GridRGPDetail.DocNo).Value = dr("RGP_NRGP_NO")
                                i += 1
                            Next
                        End If
                        i = 0
                        dgvRGPNRGPDetails.Rows.Clear()
                        If ds.Tables(3).Rows.Count > 0 Then
                            dgvRGPNRGPDetails.Rows.Add(ds.Tables(3).Rows.Count)
                            For Each dr As DataRow In ds.Tables(3).Rows
                                dgvRGPNRGPDetails.Rows(i).Cells(GridRGPNRGPDetail.Selection).Value = False
                                dgvRGPNRGPDetails.Rows(i).Cells(GridRGPNRGPDetail.DocType).Value = dr("DOC_TYPE")
                                dgvRGPNRGPDetails.Rows(i).Cells(GridRGPNRGPDetail.DocNo).Value = dr("RGP_NRGP_NO")
                                i += 1
                            Next
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            If ds IsNot Nothing Then
                ds.Dispose()
            End If
            If sqlCmd IsNot Nothing Then
                sqlCmd.Dispose()
            End If
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            If MessageBox.Show("Are you sure to close?", "eMPro", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = Windows.Forms.DialogResult.Yes Then
                Me.Close()
            End If
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
                .Rows(i).Cells(GridInvoiceDetail.Selection).Value = rdbInvoiceCheckAll.Checked
            Next
        End With
    End Sub
    Private Sub CheckUncheckRGPAll()
        If dgvRGPNRGPDetails Is Nothing OrElse dgvRGPNRGPDetails.Rows.Count = 0 Then
            rdbRGPCheckAll.Checked = False
            rdbRGPUnCheckAll.Checked = False
            Exit Sub
        End If
        With dgvRGPNRGPDetails
            For i As Integer = 0 To .Rows.Count - 1
                .Rows(i).Cells(GridRGPNRGPDetail.Selection).Value = rdbRGPCheckAll.Checked
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

    Private Sub rdbRGPCheckAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbRGPCheckAll.CheckedChanged
        Try
            CheckUncheckRGPAll()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Try
            dgvInvoiceSummary.Rows.Clear()
            dgvInvoiceDetail.Rows.Clear()
            dgvRGPDetail.Rows.Clear()
            dgvRGPNRGPDetails.Rows.Clear()
            rdbInvoiceCheckAll.Checked = False
            rdbInvoiceUncheckAll.Checked = False
            rdbRGPCheckAll.Checked = False
            rdbRGPUnCheckAll.Checked = False
            txtCustomerCode.Text = String.Empty
            lblCustCodeDes.Text = String.Empty
            txtVehicleNo.Text = String.Empty
            txtDocNo.Text = String.Empty
            txtCustomerCode.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnPrintInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintInvoice.Click
        Try
            If dgvInvoiceDetail Is Nothing OrElse dgvInvoiceDetail.Rows.Count = 0 Then
                MsgBox("No invoice found to print.", MsgBoxStyle.Exclamation, ResolveResString(100))
                Exit Sub
            End If
            Dim flag As Boolean = False

            For i As Integer = 0 To dgvInvoiceDetail.Rows.Count - 1
                If Convert.ToBoolean(dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.Selection).Value) Then
                    flag = True
                    Exit For
                End If
            Next
            If Not flag Then
                MsgBox("Please select atleast one invoice to print.", MsgBoxStyle.Exclamation, ResolveResString(100))
                dgvInvoiceDetail.Focus()
                dgvInvoiceDetail.CurrentCell = dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.Selection)
                Exit Sub
            End If

            Dim docNo As String = String.Empty
            For i As Integer = 0 To dgvInvoiceDetail.Rows.Count - 1
                If Convert.ToBoolean(dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.Selection).Value) Then
                    docNo = Convert.ToString(dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.DocNo).Value).Trim()
                    If String.IsNullOrEmpty(docNo) Then
                        MsgBox("Doc No. should not be blank.", MsgBoxStyle.Exclamation, ResolveResString(100))
                        dgvInvoiceDetail.Focus()
                        dgvInvoiceDetail.CurrentCell = dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.Selection)
                        Exit Sub
                    End If
                End If
            Next

            If isRGPNRGPMandatoryForInvoicePrint Then
                If CheckCloseBoxExist() Then
                    If Not GenerateRGPNRGP(False) Then
                        Exit Sub
                    End If
                End If
            End If

            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)

            Dim objInvoiceLocking As New InvoiceLocking(InvoiceLocking.InvoiceLockingModes.Window, InvoiceLocking.InvoiceLockingOperations.ReprintInvoice, InvoiceLocking.InvoiceButtons.Print, InvoiceLocking.InvoiceReprintOptions.Print)
            For i As Integer = 0 To dgvInvoiceDetail.Rows.Count - 1
                If Convert.ToBoolean(dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.Selection).Value) Then
                    If Convert.ToString(dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.DocNo).Value) <> "" Then
                        objInvoiceLocking.StartOperation(dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.DocNo).Value.ToString())
                    End If
                End If
            Next

            FillInvoices()
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub PrintRGP(ByVal rgpNo As String, ByVal rgpDate As String)
        Try
            Dim strreportName As String = String.Empty
            Dim datDocument_date As Date
            Dim blnflag As Boolean = False
            Dim blnBarLoc As Boolean = False
            Dim strCST_No As String = String.Empty
            Dim strLST_No As String = String.Empty
            Dim strECC_No As String = String.Empty
            Dim strTIN_No As String = String.Empty
            Dim strSelectionFormula As String = String.Empty
            Dim strDeptSQL As String = String.Empty
            Dim strDeptCode As String = String.Empty
            Dim strEmpName As String = String.Empty
            Dim strDocCat As String = String.Empty
            Dim strQSNo As String = String.Empty
            Dim rstDept As ClsResultSetDB
            Dim RdAddSold As ReportDocument
            Dim Frm As New eMProCrystalReportViewer
            RdAddSold = Frm.GetReportDocument()
            Frm.ReportHeader = Me.ctlHeader.HeaderString()
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
            Using sqlcmd As SqlCommand = New SqlCommand
                With sqlcmd
                    .CommandText = "USP_MULTIPLE_COPY_RPT"
                    .CommandTimeout = 0
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 20).Value = gstrUnitId
                    .Parameters.Add("@DOC_TYPE", SqlDbType.VarChar, 20).Value = "22"
                    .Parameters.Add("@LOCATION_CODE", SqlDbType.VarChar, 20).Value = fromLocation
                    .Parameters.Add("@DOC_NO", SqlDbType.Int).Value = rgpNo
                    .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                    SqlConnectionclass.ExecuteNonQuery(sqlcmd)
                End With
            End Using
            With RdAddSold
                '----------------------------10192393--------------------------------------------------------------------
                strreportName = GetPlantName()
                strreportName = "\Reports\rptRGPPrinting_" & strreportName & ".rpt"
                If Not CheckFile(strreportName) Then
                    strreportName = "\Reports\rptRGPPrinting.rpt"
                End If
                .Load(My.Application.Info.DirectoryPath & strreportName)
                blnflag = True
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
                strCST_No = SelectData("Cst_No", "Company_Mst", " where Unit_code = '" & gstrUnitId & "'")
                strLST_No = SelectData("Lst_No", "Company_Mst", " where Unit_code = '" & gstrUnitId & "'")
                strECC_No = SelectData("ECC_No", "Company_Mst", " where Unit_code = '" & gstrUnitId & "'")
                strTIN_No = SelectData("Tin_No", "Company_Mst", " where Unit_code = '" & gstrUnitId & "'")
                'blnBarLoc = False
                'If gblnBarcodeProcess Then
                '    blnBarLoc = Barcode_Location(fromLocation, fromLocation)
                '    If blnBarLoc = False Then GoTo nextstep
                'End If
                'stmt = "select rgp_dtl.actual_quantity,rgp_dtl.tmp_qty,item_mst.barcode_tracking from rgp_dtl inner join item_mst on rgp_dtl.item_code=item_mst.item_code AND rgp_dtl.UNIT_CODE = item_mst.UNIT_CODE where doc_no='" & txtDocNo.Text & "' AND DOC_TYPE=22 AND rgp_dtl.UNIT_CODE = '" & gstrUnitId & "'"
                'If rstemp.State = 1 Then rstemp.Close()
                'rstemp = mP_Connection.Execute(stmt)
                'If Not rstemp.EOF Then
                '    While Not rstemp.EOF
                '        If rstemp.Fields("barcode_tracking").Value = True Then
                '            'If rstemp.Fields("actual_quantity").Value = IIf(IsDBNull(rstemp.Fields("tmp_qty").Value), 0, rstemp.Fields("tmp_qty").Value) Then
                '            '-------------------------10189668------------------------------------------
                '            If Val(rstemp.Fields("actual_quantity").Value) > 0 Then
                '                blnflag = True
                '            Else
                '                blnflag = False
                '                GoTo nextstep
                '            End If
                '        End If
                '        rstemp.MoveNext()
                '    End While
                'End If
                'nextstep:

                If strreportName.ToLower = "\Reports\rptRGPPrinting.rpt" Or strreportName.ToUpper = "\Reports\rptRGPPrinting.rpt" Then
                    strSelectionFormula = "{rgp_hdr.doc_type}= 22 and {RGP_Hdr.from_location} ='" & fromLocation & "' AND {RGP_Hdr.unit_code} = '" & gstrUnitId & "' and {RGP_Hdr.Doc_No} =  " & rgpNo & "  AND  {TEMP_MULTIPLE_COPY.ip_address} = '" & gstrIpaddressWinSck & "' "
                Else
                    strSelectionFormula = "{rgp_hdr.doc_type}= 22 and {RGP_Hdr.from_location} ='" & fromLocation & "' AND {RGP_Hdr.unit_code} = '" & gstrUnitId & "' and {RGP_Hdr.Doc_No} =  " & rgpNo & ""

                End If
                ' To send the report to Printer
                .DataDefinition.FormulaFields("empcode").Text = "'" & Trim(mP_User) & "'"
                strDeptSQL = "SELECT description,name FROM employee_mst,profit_center_mst WHERE employee_code='" & Trim(mP_User) & "' and profit_center_code=department_code AND employee_mst.UNIT_CODE = '" & gstrUnitId & "'"
                rstDept = New ClsResultSetDB
                Call rstDept.GetResult(strDeptSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rstDept.GetNoRows > 0 Then
                    strDeptCode = rstDept.GetValue("description")
                    strEmpName = rstDept.GetValue("name")
                    .DataDefinition.FormulaFields("empname").Text = "'" & strEmpName & "'"
                    .DataDefinition.FormulaFields("departmentcode").Text = "'" & strDeptCode & "'"
                End If
                rstDept.ResultSetClose()
                rstDept = Nothing
                strDeptSQL = "SELECT doc_category FROM rgp_hdr WHERE doc_type=22 and from_location ='" & fromLocation & "' AND UNIT_CODE = '" & gstrUnitId & "' and doc_no=" & rgpNo
                rstDept = New ClsResultSetDB
                Call rstDept.GetResult(strDeptSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rstDept.GetNoRows > 0 Then
                    If UCase(rstDept.GetValue("doc_category")) = "J" Then
                        strDocCat = "JobWork"
                    Else
                        strDocCat = "Miscellaneous"
                    End If
                    .DataDefinition.FormulaFields("doccatdesc").Text = "'" & strDocCat & "'"
                End If
                rstDept.ResultSetClose()
                rstDept = Nothing
                rstDept = New ClsResultSetDB
                Call rstDept.GetResult("SELECT phone FROM company_mst WHERE UNIT_CODE = '" & gstrUnitId & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                .DataDefinition.FormulaFields("CompanyName").Text = "'" & gstrCOMPANY & "'"
                .DataDefinition.FormulaFields("WorkAddress").Text = "'" & gstr_WRK_ADDRESS1 & " " & gstr_WRK_ADDRESS2 & "'"
                '10577079-Change in Register address Format.
                .DataDefinition.FormulaFields("RegisteredOfficeAddress1").Text = "'" & gstr_RGN_ADDRESS1 + gstr_RGN_ADDRESS2 & "'"
                .DataDefinition.FormulaFields("CST_No").Text = "'" & strCST_No & "'"
                .DataDefinition.FormulaFields("LST_No").Text = "'" & strLST_No & "'"
                .DataDefinition.FormulaFields("ECC_No").Text = "'" & strECC_No & "'"
                .DataDefinition.FormulaFields("fmlflag").Text = "" & blnflag & ""
                .DataDefinition.FormulaFields("Tin_No").Text = "'" & strTIN_No & "'"
                .DataDefinition.FormulaFields("WorkTelephoneNo").Text = "'" & rstDept.GetValue("phone") & "'"
                If gblnStoreConfig.BatchTracking = True Then '   @@
                    .DataDefinition.FormulaFields("SuppressBatches").Text = "False"
                Else
                    .DataDefinition.FormulaFields("SuppressBatches").Text = "True"
                End If


                If QSRequired() = True Then 'To check if the QS Format no. is to printed or not
                    datDocument_date = ConvertToDate(rgpDate) 'date to be passed
                    strQSNo = QSFormatNumber("rptRGPPrinting", datDocument_date)
                    .DataDefinition.FormulaFields("QSFormatNo").Text = "'" & Trim(strQSNo) & "'"
                Else
                    .DataDefinition.FormulaFields("QSFormatNo").Text = "'" & "" & "'"
                End If
                .RecordSelectionFormula = strSelectionFormula
                'Frm.Show()
                Frm.SetReportDocument()
                For i As Integer = 0 To printCountRGPNRGP - 1
                    .PrintToPrinter(1, False, 0, 0)
                Next
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Function SelectData(ByVal pstrFName As String, ByVal pstrTName As String, Optional ByVal pstrCon As String = "") As String
        '****************************************************
        'Description    -  Getting the Field name,table name and condition information and returning the field's information.
        'Arguments      -  pstrFName - Field Name,pstrTName - Table Name,pstrCon - Condition
        'Return Value   -  Field's Information
        '****************************************************
        On Error GoTo ErrHandler
        Dim strSelectSql As String 'Declared To Make Select Query
        Dim rsGetDes As ClsResultSetDB
        strSelectSql = "Select " & Trim(pstrFName) & " from " & Trim(pstrTName) & " " & pstrCon
        rsGetDes = New ClsResultSetDB
        rsGetDes.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsGetDes.GetNoRows > 0 Then
            SelectData = rsGetDes.GetValue(pstrFName)
        Else
            SelectData = ""
        End If
        rsGetDes.ResultSetClose()
        rsGetDes = Nothing
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function

    Private Sub PrintNRGP(ByVal nrgpNo As String, ByVal nrgpDate As String)
        Try
            Dim strSelectionFormula As String = String.Empty
            Dim strQSNo As String = String.Empty
            Dim datDocument_date As Date
            Dim strDeptCode As String = String.Empty
            Dim strDocCat As String = String.Empty
            Dim strEmpName As String = String.Empty
            Dim strCST_No As String = String.Empty
            Dim strLST_No As String = String.Empty
            Dim strECC_No As String = String.Empty
            Dim strTIN_No As String = String.Empty
            Dim strPhoneNo As String = String.Empty
            Dim strsql As String = String.Empty
            Dim oRs As ADODB.Recordset
            Dim blnNRGPCompleted As Boolean = True
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
            Using sqlcmd As SqlCommand = New SqlCommand
                With sqlcmd
                    .CommandText = "USP_MULTIPLE_COPY_RPT"
                    .CommandTimeout = 0
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 20).Value = gstrUnitId
                    .Parameters.Add("@DOC_TYPE", SqlDbType.VarChar, 20).Value = "23"
                    .Parameters.Add("@LOCATION_CODE", SqlDbType.VarChar, 20).Value = fromLocation
                    .Parameters.Add("@DOC_NO", SqlDbType.Int).Value = nrgpNo
                    .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                    SqlConnectionclass.ExecuteNonQuery(sqlcmd)
                End With
            End Using
            'blnBarcodeLoc = Barcode_Location(txtLocFrom.Text.Trim, txtLocTo.Text.Trim)
            'vardata = Nothing
            'Call SpRGP.GetText(enmNRGPGrid.Col_Item_Code, 1, vardata)
            'If (gblnBarcodeProcess = True) And (blnBarcodeLoc = True) And IsValidItem(CStr(vardata)) Then
            '    strsql = "SELECT DBO.UFN_GET_NRGP_COMPLETED" & _
            '            "('" & gstrUnitId & "','" & _
            '            txtLocFrom.Text.Trim & "'," & _
            '            txtDocNo.Text.Trim & _
            '            ") AS FIELD"
            '    oRs = mP_Connection.Execute(strsql)
            '    blnNRGPCompleted = oRs.Fields("Field").Value
            '    oRs.Close()
            '    oRs = Nothing
            'Else
            '    blnNRGPCompleted = True
            'End If
            strCST_No = String.Empty
            strLST_No = String.Empty
            strECC_No = String.Empty
            strTIN_No = String.Empty
            strPhoneNo = String.Empty
            strsql = "SELECT ISNULL(PHONE,'') AS PHONE,ISNULL(Cst_No,'') AS Cst_No,ISNULL(Lst_No,'') AS Lst_No," & _
                    "ISNULL(ECC_No,'') AS ECC_No,ISNULL(TIN_No,'') AS TIN_No" & _
                    " FROM Company_Mst WHERE UNIT_CODE = '" & gstrUnitId & "'"
            oRs = mP_Connection.Execute(strsql)
            If Not (oRs.EOF And oRs.BOF) Then
                strCST_No = oRs.Fields("Cst_No").Value
                strLST_No = oRs.Fields("Lst_No").Value
                strECC_No = oRs.Fields("ECC_No").Value
                strTIN_No = oRs.Fields("tin_no").Value
                strPhoneNo = oRs.Fields("Phone").Value
            End If
            oRs.Close()
            oRs = Nothing
            strDeptCode = String.Empty
            strEmpName = String.Empty
            strsql = "SELECT DESCRIPTION,NAME" & _
                    " FROM EMPLOYEE_MST E" & _
                    " INNER JOIN PROFIT_CENTER_MST P" & _
                    " ON" & _
                    " (PROFIT_CENTER_CODE = DEPARTMENT_CODE AND E.UNIT_CODE = P.UNIT_CODE" & _
                    " AND EMPLOYEE_CODE='" & mP_User.Trim & "' AND E.UNIT_CODE = '" & gstrUnitId & "')"
            oRs = mP_Connection.Execute(strsql)
            If Not (oRs.EOF And oRs.BOF) Then
                strDeptCode = oRs.Fields("Description").Value
                strEmpName = oRs.Fields("Name").Value
            End If
            oRs.Close()
            oRs = Nothing
            strDocCat = String.Empty
            strsql = "SELECT doc_category FROM rgp_hdr" & _
                    " WHERE doc_type=23" & _
                    " and from_location='" & fromLocation & "'" & _
                    "  AND UNIT_CODE = '" & gstrUnitId & "' and doc_no = " & nrgpNo
            oRs = mP_Connection.Execute(strsql)
            If Not (oRs.EOF And oRs.BOF) Then
                If UCase(oRs.Fields("Doc_Category").Value) = "J" Then
                    strDocCat = "JobWork"
                Else
                    strDocCat = "Miscellaneous"
                End If
            End If
            oRs.Close()
            oRs = Nothing
            Dim RdAddSold As ReportDocument
            Dim strflag As Boolean = False
            Dim Frm As New eMProCrystalReportViewer
            RdAddSold = Frm.GetReportDocument()
            Frm.ReportHeader = Me.ctlHeader.HeaderString()
            With RdAddSold
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
                .Load(My.Application.Info.DirectoryPath & "\Reports\rptNRGPPrinting.rpt")
                strSelectionFormula = "{rgp_hdr.doc_type}= 23" & _
                "  and {rgp_hdr.unit_code}= '" & gstrUnitId & "' and {RGP_Hdr.from_location}='" & fromLocation & "'" & _
                " and {RGP_Hdr.Doc_No} = " & nrgpNo & "" & _
                " and  {TEMP_MULTIPLE_COPY.ip_address} = '" & gstrIpaddressWinSck & "'"
                .RecordSelectionFormula = strSelectionFormula
                .DataDefinition.FormulaFields("empcode").Text = "'" & mP_User.Trim & "'"
                .DataDefinition.FormulaFields("empname").Text = "'" & strEmpName & "'"
                .DataDefinition.FormulaFields("departmentcode").Text = "'" & strDeptCode & "'"
                .DataDefinition.FormulaFields("doccatdesc").Text = "'" & strDocCat & "'"
                .DataDefinition.FormulaFields("CompanyName").Text = "'" & gstrCOMPANY & "'"
                .DataDefinition.FormulaFields("WorkAddress").Text = "'" & gstr_WRK_ADDRESS1 & " " & gstr_WRK_ADDRESS2 & "'"
                '10577079-Change in Register address Format.
                .DataDefinition.FormulaFields("RegisteredOfficeAddress1").Text = "'" & gstr_RGN_ADDRESS1 + gstr_RGN_ADDRESS2 & "'"
                .DataDefinition.FormulaFields("CST_No").Text = "'" & strCST_No & "'"
                .DataDefinition.FormulaFields("LST_No").Text = "'" & strLST_No & "'"
                .DataDefinition.FormulaFields("WorkTelephoneNo").Text = "'" & strPhoneNo & "'"
                If (gblnStoreConfig.BatchTracking = True) Then
                    .DataDefinition.FormulaFields("SuppressBatches").Text = "False"
                Else
                    .DataDefinition.FormulaFields("SuppressBatches").Text = "True"
                End If
                If QSRequired() = True Then
                    datDocument_date = ConvertToDate(nrgpDate)
                    strQSNo = QSFormatNumber("rptNRGPPrinting", datDocument_date)
                    .DataDefinition.FormulaFields("QSFormatNo").Text = "'" & Trim(strQSNo) & "'"
                Else
                    .DataDefinition.FormulaFields("QSFormatNo").Text = "'" & String.Empty & "'"
                End If
                .DataDefinition.FormulaFields("ECC_No").Text = "'" & strECC_No & "'"
                'If (OptCustType(2).Checked = True) Then
                '    'Modified by -Prachi Jain ,Issue id-10519838
                '    strsql = "SELECT employee_code,name,isnull(address_1,'') AS address_1," & _
                '            " ISNULL(address_2,'') AS address_2,ISNULL(city,'') AS city,ISNULL(state,'') AS state" & _
                '            " FROM employee_mst" & _
                '            " WHERE employee_code='" & txtLocTo.Text.Trim & "'  AND UNIT_CODE = '" & gstrUnitId & "'  AND  ISNULL(LEFT_DATE,'')=''"
                '    oRs = mP_Connection.Execute(strsql)
                '    .DataDefinition.FormulaFields("Employee_code").Text = "'" & Trim(oRs.Fields("EMPLOYEE_CODE").Value) & "'"
                '    .DataDefinition.FormulaFields("Employee_name").Text = "'" & Trim(oRs.Fields("Name").Value) & "'"
                '    .DataDefinition.FormulaFields("Employee_address_1").Text = "'" & Trim(oRs.Fields("Address_1").Value) & "'"
                '    .DataDefinition.FormulaFields("Employee_address_2").Text = "'" & Trim(oRs.Fields("Address_2").Value) & "'"
                '    .DataDefinition.FormulaFields("Employee_city").Text = "'" & Trim(oRs.Fields("City").Value) & "'"
                '    .DataDefinition.FormulaFields("Employee_state").Text = "'" & Trim(oRs.Fields("State").Value) & "'"
                '    oRs.Close()
                '    oRs = Nothing
                'End If
                .DataDefinition.FormulaFields("FMLFLAG").Text = "" & blnNRGPCompleted
                .DataDefinition.FormulaFields("TIN_No").Text = "'" & strTIN_No & "'"
                'Frm.Show()
                Frm.SetReportDocument()
                For i As Integer = 0 To printCountRGPNRGP - 1
                    .PrintToPrinter(1, False, 0, 0)
                Next
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub btnExceptionInvoices_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExceptionInvoices.Click
        Try
            Dim objExceptionInvoices As New frmExceptionInvoices
            If Not String.IsNullOrEmpty(txtDocNo.Text.Trim()) Then
                objExceptionInvoices.SetDispLotNo = txtDocNo.Text
            End If
            objExceptionInvoices.ShowDialog()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Function GenerateRGPNRGP(ByVal blnIsMsgShowing As Boolean) As Boolean
        Dim result As Boolean = False
        Dim sqlCmd As New SqlCommand
        Try
            If dtpDateFrom.Value > dtpDateTo.Value Then
                MessageBox.Show("[Date From] should be less than or equal to [Date To].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                dtpDateFrom.Focus()
                Return result
            ElseIf String.IsNullOrEmpty(txtCustomerCode.Text) Then
                MessageBox.Show("Please first select [Customer Code].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtCustomerCode.Text = String.Empty
                lblCustCodeDes.Text = String.Empty
                txtCustomerCode.Focus()
                Return result
            ElseIf String.IsNullOrEmpty(txtVehicleNo.Text) Then
                MessageBox.Show("Please first select [Vehicle No].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtVehicleNo.Text = String.Empty
                txtVehicleNo.Focus()
                Return result
            ElseIf String.IsNullOrEmpty(txtDocNo.Text) Then
                MessageBox.Show("Please first select [Disp. Lot No].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtDocNo.Text = String.Empty
                txtDocNo.Focus()
                Return result
            End If

            If blnIsMsgShowing Then
                If MessageBox.Show("Are you sure to create RGP / NRGP?", "eMPro", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly) = Windows.Forms.DialogResult.No Then
                    Return result
                End If
            End If

            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)

            With sqlCmd
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 300 ' 5 Minute
                .CommandText = "USP_BSR_AUTO_RGP_NRGP"
                .Parameters.Clear()
                .Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
                .Parameters.AddWithValue("@CUSTOMER_CODE", Trim(txtCustomerCode.Text))
                .Parameters.AddWithValue("@DISPATCH_LOT_NO", Trim(txtDocNo.Text))
                .Parameters.AddWithValue("@IP_ADDRESS", gstrIpaddressWinSck)
                .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                If Convert.ToString(.Parameters("@MESSAGE").Value) <> "" Then
                    MsgBox(Convert.ToString(.Parameters("@MESSAGE").Value), MsgBoxStyle.Exclamation, ResolveResString(100))
                    Return result
                Else
                    result = True
                    If blnIsMsgShowing Then
                        MessageBox.Show("RGP / NRGP created successfully.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    End If
                End If
            End With
        Catch ex As Exception
            result = False
            RaiseException(ex)
        Finally
            If sqlCmd IsNot Nothing Then
                sqlCmd.Dispose()
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        End Try
        Return result
    End Function

    Private Sub btnGenerateRGPNRGP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenerateRGPNRGP.Click
        Try
            If GenerateRGPNRGP(True) Then
                FillInvoices()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnPrintRGP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintRGP.Click
        Try
            If dgvRGPNRGPDetails Is Nothing OrElse dgvRGPNRGPDetails.Rows.Count = 0 Then
                MsgBox("No RGP / NRGP found to print.", MsgBoxStyle.Exclamation, ResolveResString(100))
                Exit Sub
            End If
            Dim flag As Boolean = False

            For i As Integer = 0 To dgvRGPNRGPDetails.Rows.Count - 1
                If Convert.ToBoolean(dgvRGPNRGPDetails.Rows(i).Cells(GridRGPNRGPDetail.Selection).Value) Then
                    flag = True
                    Exit For
                End If
            Next
            If Not flag Then
                MsgBox("Please select atleast one RGP / NRGP to print.", MsgBoxStyle.Exclamation, ResolveResString(100))
                dgvRGPNRGPDetails.Focus()
                dgvRGPNRGPDetails.CurrentCell = dgvRGPNRGPDetails.Rows(i).Cells(GridRGPNRGPDetail.Selection)
                Exit Sub
            End If

            Dim docNo As String = String.Empty
            Dim docType As String = String.Empty
            Dim docDt As String = String.Empty

            For i As Integer = 0 To dgvRGPNRGPDetails.Rows.Count - 1
                If Convert.ToBoolean(dgvRGPNRGPDetails.Rows(i).Cells(GridRGPNRGPDetail.Selection).Value) Then
                    docNo = Convert.ToString(dgvRGPNRGPDetails.Rows(i).Cells(GridRGPNRGPDetail.DocNo).Value).Trim()
                    docType = Convert.ToString(dgvRGPNRGPDetails.Rows(i).Cells(GridRGPNRGPDetail.DocType).Value).Trim()
                    If String.IsNullOrEmpty(docType) Then
                        MsgBox("Doc Type should not be blank.", MsgBoxStyle.Exclamation, ResolveResString(100))
                        dgvRGPNRGPDetails.Focus()
                        dgvRGPNRGPDetails.CurrentCell = dgvRGPNRGPDetails.Rows(i).Cells(GridRGPNRGPDetail.Selection)
                        Exit Sub
                    End If
                    If String.IsNullOrEmpty(docNo) Then
                        MsgBox("Doc No. should not be blank.", MsgBoxStyle.Exclamation, ResolveResString(100))
                        dgvRGPNRGPDetails.Focus()
                        dgvRGPNRGPDetails.CurrentCell = dgvRGPNRGPDetails.Rows(i).Cells(GridRGPNRGPDetail.Selection)
                        Exit Sub
                    End If
                End If
            Next

            docNo = String.Empty
            docType = String.Empty
            docDt = String.Empty
            For i As Integer = 0 To dgvRGPNRGPDetails.Rows.Count - 1
                If Convert.ToBoolean(dgvRGPNRGPDetails.Rows(i).Cells(GridRGPNRGPDetail.Selection).Value) Then
                    docType = Convert.ToString(dgvRGPNRGPDetails.Rows(i).Cells(GridRGPNRGPDetail.DocType).Value).Trim()
                    docNo = Convert.ToString(dgvRGPNRGPDetails.Rows(i).Cells(GridRGPNRGPDetail.DocNo).Value).Trim()
                    If docType <> "" AndAlso docNo <> "" Then
                        If docType.ToUpper() = "RGP" Then
                            docDt = Convert.ToString(SqlConnectionclass.ExecuteScalar("SELECT CONVERT(VARCHAR(15),RGP_DATE,103) AS RGP_DATE FROM RGP_HDR (NOLOCK) WHERE UNIT_CODE='" & gstrUnitId & "' AND DOC_TYPE='22' AND DOC_NO=" & docNo & " AND FROM_LOCATION='" & fromLocation & "'"))
                            PrintRGP(docNo, docDt)
                        ElseIf docType.ToUpper() = "NRGP" Then
                            docDt = Convert.ToString(SqlConnectionclass.ExecuteScalar("SELECT CONVERT(VARCHAR(15),RGP_DATE,103) AS RGP_DATE FROM RGP_HDR (NOLOCK) WHERE UNIT_CODE='" & gstrUnitId & "' AND DOC_TYPE='23' AND DOC_NO=" & docNo & " AND FROM_LOCATION='" & fromLocation & "'"))
                            PrintNRGP(docNo, docDt)
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Function CheckCloseBoxExist() As Boolean
        Dim result As Boolean = False
        result = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("SELECT DBO.UDF_BSR_CHECK_RGP_NRGP_EXIST('" & gstrUnitId & "','" & Trim(txtCustomerCode.Text) & "'," & Trim(txtDocNo.Text) & ")"))
        Return result
    End Function
End Class