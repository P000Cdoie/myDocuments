Imports System.Data
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
'*********************************************************************************************************************
'Copyright(c)       - MIND
'Name of Module     - Bulk Invoice Printing (Generic)
'Name of Form       - FRMMKTTRN0125  , Bulk Invoice Printing (Generic)
'Created by         - Ashish sharma
'Created Date       - 07 JUL 2022
'description        - Bulk invoice Printing (Generic)
'*********************************************************************************************************************
Public Class FRMMKTTRN0125
    Private Enum GridInvoiceDetail
        Selection = 0
        DocNo
        InvoiceType
        InvoiceSubType
        IRNNo
        EwayBillNo
    End Enum

    Private Sub FRMMKTTRN0125_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            txtCustomerCode.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FRMMKTTRN0125_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
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

    Private Sub FRMMKTTRN0125_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Call FitToClient(Me, GrpMain, ctlHeader, GrpBoxButtons, 550)
            Me.MdiParent = mdifrmMain
            dtpDateFrom.Value = GetServerDate()
            dtpDateTo.Value = GetServerDate()
            '' PRAVEEN DIGITAL SIGN
            mblnISTrueSignRequired = CBool(Find_Value("SELECT ISNULL(IS_TRUE_SIGN_REQUIRED,0) FROM gen_unitmaster (NOLOCK) WHERE Unt_CodeID='" + gstrUNITID + "'"))
            mblnAPIUrl = Find_Value("Select API_Url from gen_unitmaster (NOLOCK) where Unt_CodeID = '" & gstrUNITID & "'")
            mblnPFX_ID = Find_Value("Select PFX_ID from gen_unitmaster (NOLOCK) where Unt_CodeID = '" & gstrUNITID & "'")
            mblnPFX_Pass = Find_Value("Select PFX_password from gen_unitmaster (NOLOCK) where Unt_CodeID = '" & gstrUNITID & "'")
            mblnAPI_Key = Find_Value("Select API_Key from gen_unitmaster (NOLOCK) where Unt_CodeID = '" & gstrUNITID & "'")
            ConfigureGridColumn()
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
    Private Sub ConfigureGridColumn()
        Try
            dgvInvoiceDetail.Columns.Clear()

            Dim objChkBox As New DataGridViewCheckBoxColumn
            objChkBox.Name = "Selection"
            objChkBox.HeaderText = " "

            dgvInvoiceDetail.Columns.Add(objChkBox)
            dgvInvoiceDetail.Columns.Add("DocNo", "Invoice No.")
            dgvInvoiceDetail.Columns.Add("InvoiceType", "Invoice Type")
            dgvInvoiceDetail.Columns.Add("InvoiceSubType", "Invoice Sub Type")
            dgvInvoiceDetail.Columns.Add("IRN", "IRN")
            dgvInvoiceDetail.Columns.Add("EwayBillNo", "EwayBillNo.")

            dgvInvoiceDetail.Columns(GridInvoiceDetail.Selection).Width = 35
            dgvInvoiceDetail.Columns(GridInvoiceDetail.DocNo).Width = 120
            dgvInvoiceDetail.Columns(GridInvoiceDetail.InvoiceType).Width = 120
            dgvInvoiceDetail.Columns(GridInvoiceDetail.InvoiceSubType).Width = 120
            dgvInvoiceDetail.Columns(GridInvoiceDetail.IRNNo).Width = 250
            dgvInvoiceDetail.Columns(GridInvoiceDetail.EwayBillNo).Width = 150

            dgvInvoiceDetail.Columns(GridInvoiceDetail.Selection).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvInvoiceDetail.Columns(GridInvoiceDetail.DocNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvInvoiceDetail.Columns(GridInvoiceDetail.InvoiceType).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvInvoiceDetail.Columns(GridInvoiceDetail.InvoiceSubType).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvInvoiceDetail.Columns(GridInvoiceDetail.IRNNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvInvoiceDetail.Columns(GridInvoiceDetail.EwayBillNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            dgvInvoiceDetail.Columns(GridInvoiceDetail.DocNo).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetail.InvoiceType).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetail.InvoiceSubType).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetail.IRNNo).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetail.EwayBillNo).ReadOnly = True

            dgvInvoiceDetail.Columns(GridInvoiceDetail.Selection).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvInvoiceDetail.Columns(GridInvoiceDetail.DocNo).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvInvoiceDetail.Columns(GridInvoiceDetail.InvoiceType).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvInvoiceDetail.Columns(GridInvoiceDetail.InvoiceSubType).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvInvoiceDetail.Columns(GridInvoiceDetail.IRNNo).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvInvoiceDetail.Columns(GridInvoiceDetail.EwayBillNo).SortMode = DataGridViewColumnSortMode.NotSortable
        Catch ex As Exception
            RaiseException(ex)
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
                strSql = "SELECT CUST_CODE,CUST_NAME FROM DBO.UDF_GET_CUSTOMERS_BULK_INVOICE_PRINTING('" & gstrUnitId & "','" & dtpDateFrom.Value.ToString("dd MMM yyyy") & "','" & dtpDateTo.Value.ToString("dd MMM yyyy") & "') ORDER BY CUST_CODE"
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
                strSql = "SELECT VEHICLE_NO,VEHICLE_NO AS VEHICLE_NO_DESC FROM DBO.UDF_GET_VEHICLE_BULK_INVOICE_PRINTING('" & gstrUnitId & "','" & Trim(txtCustomerCode.Text) & "','" & dtpDateFrom.Value.ToString("dd MMM yyyy") & "','" & dtpDateTo.Value.ToString("dd MMM yyyy") & "') ORDER BY VEHICLE_NO"
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
            dgvInvoiceDetail.Rows.Clear()
            rdbInvoiceCheckAll.Checked = False
            rdbInvoiceUncheckAll.Checked = False
        Catch ex As Exception
            RaiseException(ex)
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

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Try
            dgvInvoiceDetail.Rows.Clear()
            rdbInvoiceCheckAll.Checked = False
            rdbInvoiceUncheckAll.Checked = False
            txtCustomerCode.Text = String.Empty
            lblCustCodeDes.Text = String.Empty
            txtVehicleNo.Text = String.Empty
            txtCustomerCode.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub rdbInvoiceCheckAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbInvoiceCheckAll.CheckedChanged
        Try
            CheckUncheckInvoicesAll()
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
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
            FillInvoices()
            rdbInvoiceCheckAll.Checked = False
            rdbInvoiceUncheckAll.Checked = False
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub btnExceptionInvoices_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExceptionInvoices.Click
        Try
            Dim objExceptionInvoices As New frmExceptionInvoices
            objExceptionInvoices.ShowDialog()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FillInvoices()
        Dim dt As New DataTable
        Dim sqlCmd As New SqlCommand
        Try
            Dim i As Integer = 0
            With sqlCmd
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 300 ' 5 Minute
                .CommandText = "USP_GET_BULK_INVOICE_PRINTING_DATA"
                .Parameters.Clear()
                .Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
                .Parameters.AddWithValue("@FROM_DATE", dtpDateFrom.Value.ToString("dd MMM yyyy"))
                .Parameters.AddWithValue("@TO_DATE", dtpDateTo.Value.ToString("dd MMM yyyy"))
                .Parameters.AddWithValue("@CUSTOMER_CODE", Trim(txtCustomerCode.Text))
                If Not String.IsNullOrEmpty(txtVehicleNo.Text) Then
                    .Parameters.AddWithValue("@VEHICLE_NO", Trim(txtVehicleNo.Text))
                End If
                .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                dt = SqlConnectionclass.GetDataTable(sqlCmd)
                If Convert.ToString(.Parameters("@MESSAGE").Value) <> "" Then
                    MsgBox(Convert.ToString(.Parameters("@MESSAGE").Value), MsgBoxStyle.Exclamation, ResolveResString(100))
                Else
                    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                        dgvInvoiceDetail.Rows.Clear()
                        dgvInvoiceDetail.Rows.Add(dt.Rows.Count)
                        For Each dr As DataRow In dt.Rows
                            dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.Selection).Value = False
                            dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.DocNo).Value = dr("DOC_NO")
                            dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.InvoiceType).Value = dr("INVOICE_TYPE")
                            dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.InvoiceSubType).Value = dr("SUB_TYPE")
                            dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.IRNNo).Value = dr("IRN_NO")
                            dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.EwayBillNo).Value = dr("EWAY_BILL_NO")
                            i += 1
                        Next
                    End If
                End If
            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            If dt IsNot Nothing Then
                dt.Dispose()
            End If
            If sqlCmd IsNot Nothing Then
                sqlCmd.Dispose()
            End If
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
End Class