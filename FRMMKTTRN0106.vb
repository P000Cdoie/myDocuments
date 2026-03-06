Imports System.Data.SqlClient
'*********************************************************************************************************************
'Copyright(c)       - MIND
'Name of Module     - e-Invoice Manual Submission
'Name of Form       - FRMMKTTRN0106  , e-Invoice Manual Submission
'Created by         - Ashish sharma
'Created Date       - 17 MAY 2018
'description        - e-Invoice Manual Submission (New Development)
'*********************************************************************************************************************
Public Class FRMMKTTRN0106
    Private Const Normal As String = "NORMAL"
    Private Enum enumBillDetail
        TICK = 1
        DOC_NO
        DOC_DATE
        VOUCHER_NO
        DRCR
        SUPP_INV_TYPE
    End Enum
    Private Sub FRMMKTTRN0106_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        cmbInvoiceType.Focus()
    End Sub
    Private Sub FRMMKTTRN0106_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Call FitToClient(Me, GrpMain, ctlHeader, GrpBoxButtons, 600)
            Me.MdiParent = mdifrmMain
            dtpDateFrom.Value = GetServerDate()
            dtpDateTo.Value = GetServerDate()
            cmbInvoiceType.SelectedIndex = 0
            InitializeSpread()
            lblMessage.BackColor = Color.Coral
            lblMessage.ForeColor = Color.DarkBlue
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub cmdInvoiceFrom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInvoiceFrom.Click
        Dim strHelp() As String = Nothing
        Dim strSql As String = String.Empty
        Try
            If dtpDateFrom.Value > dtpDateTo.Value Then
                MessageBox.Show("[Date From] should be less than or equal to [Date To].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                dtpDateFrom.Focus()
                Exit Sub
            ElseIf String.IsNullOrEmpty(cmbInvoiceType.Text) Then
                MessageBox.Show("Please first select [Invoice Type].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                cmbInvoiceType.Focus()
                Exit Sub
            ElseIf String.IsNullOrEmpty(txtCustomerCode.Text) Then
                MessageBox.Show("Please first select [Customer Code].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtCustomerCode.Focus()
                Exit Sub
            Else
                If cmbInvoiceType.Text.ToUpper() = FRMMKTTRN0106.Normal Then
                    strSql = "SELECT INVOICE_NO,INVOICE_DATE FROM DBO.UDF_eINVOICE_MANUAL_SUBMISSION_INVOICE_HELP('" & gstrUnitId & "','" & dtpDateFrom.Value.ToString("dd MMM yyyy") & "','" & dtpDateTo.Value.ToString("dd MMM yyyy") & "','" & cmbInvoiceType.Text.ToUpper() & "','" & txtCustomerCode.Text & "') ORDER BY INVOICE_NO"
                Else
                    strSql = "SELECT INVOICE_NO,INVOICE_DATE,VOUCHER_NO FROM DBO.UDF_eINVOICE_MANUAL_SUBMISSION_INVOICE_HELP('" & gstrUnitId & "','" & dtpDateFrom.Value.ToString("dd MMM yyyy") & "','" & dtpDateTo.Value.ToString("dd MMM yyyy") & "','" & cmbInvoiceType.Text.ToUpper() & "','" & txtCustomerCode.Text & "') ORDER BY INVOICE_NO"
                End If
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
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtInvoiceFrom_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtInvoiceFrom.KeyDown
        Dim objText As TextBox = DirectCast(sender, TextBox)
        If e.KeyCode = Keys.Delete Then
            objText.Text = String.Empty
        ElseIf e.KeyCode = Keys.F1 Then
            cmdInvoiceFrom_Click(sender, e)
        End If
    End Sub

    Private Sub CmdCustCodeHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdCustCodeHelp.Click
        Try
            Dim strHelp() As String = Nothing
            Dim strSql As String = String.Empty
            Try
                If dtpDateFrom.Value > dtpDateTo.Value Then
                    MessageBox.Show("[Date From] should be less than or equal to [Date To].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    dtpDateFrom.Focus()
                    Exit Sub
                ElseIf String.IsNullOrEmpty(cmbInvoiceType.Text) Then
                    MessageBox.Show("Please first select [Invoice Type].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    cmbInvoiceType.Focus()
                    Exit Sub
                Else
                    strSql = "SELECT CUSTOMER_CODE,CUSTOMER_NAME FROM DBO.UDF_eINVOICE_MANUAL_SUBMISSION_CUSTOMER_HELP('" & gstrUnitId & "') ORDER BY CUSTOMER_CODE"
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
            ElseIf String.IsNullOrEmpty(cmbInvoiceType.Text) Then
                MessageBox.Show("Please first select [Invoice Type].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                cmbInvoiceType.Focus()
                Exit Sub
            ElseIf String.IsNullOrEmpty(txtCustomerCode.Text) Then
                MessageBox.Show("Please first select [Customer Code].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtCustomerCode.Focus()
                Exit Sub
            ElseIf String.IsNullOrEmpty(txtInvoiceFrom.Text) Then
                MessageBox.Show("Please first select [Invoice From].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtInvoiceFrom.Focus()
                Exit Sub
            ElseIf String.IsNullOrEmpty(txtInvoiceTo.Text) Then
                MessageBox.Show("Please first select [Invoice To].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtInvoiceTo.Focus()
                Exit Sub
            ElseIf Not String.IsNullOrEmpty(txtInvoiceTo.Text) And Not String.IsNullOrEmpty(txtInvoiceFrom.Text) Then
                If Val(txtInvoiceFrom.Text) > Val(txtInvoiceTo.Text) Then
                    MessageBox.Show("[Invoice From] should be less than or equal to [Invoice To].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    txtInvoiceTo.Focus()
                    Exit Sub
                End If
            End If
            FillInvoices()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            If MessageBox.Show("Are you sure you want close?", "eMPro", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = Windows.Forms.DialogResult.Yes Then
                Me.Close()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            cmbInvoiceType.SelectedIndex = 0
            txtCustomerCode.Text = String.Empty
            lblCustCodeDes.Text = String.Empty
            txtInvoiceFrom.Text = String.Empty
            txtInvoiceTo.Text = String.Empty
            fspInvoices.MaxRows = 0
            rdbCheckAll.Checked = False
            rdbUncheckAll.Checked = False
            cmbInvoiceType.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub cmdInvoiceTo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInvoiceTo.Click
        Dim strHelp() As String = Nothing
        Dim strSql As String = String.Empty
        Try
            If dtpDateFrom.Value > dtpDateTo.Value Then
                MessageBox.Show("[Date From] should be less than or equal to [Date To].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                dtpDateFrom.Focus()
                Exit Sub
            ElseIf String.IsNullOrEmpty(cmbInvoiceType.Text) Then
                MessageBox.Show("Please first select [Invoice Type].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                cmbInvoiceType.Focus()
                Exit Sub
            ElseIf String.IsNullOrEmpty(txtCustomerCode.Text) Then
                MessageBox.Show("Please first select [Customer Code].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtCustomerCode.Focus()
                Exit Sub
            ElseIf String.IsNullOrEmpty(txtInvoiceFrom.Text) Then
                MessageBox.Show("Please first select [Invoice From].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtInvoiceFrom.Focus()
                Exit Sub
            Else
                If cmbInvoiceType.Text.ToUpper() = FRMMKTTRN0106.Normal Then
                    strSql = "SELECT INVOICE_NO,INVOICE_DATE FROM DBO.UDF_eINVOICE_MANUAL_SUBMISSION_INVOICE_HELP('" & gstrUnitId & "','" & dtpDateFrom.Value.ToString("dd MMM yyyy") & "','" & dtpDateTo.Value.ToString("dd MMM yyyy") & "','" & cmbInvoiceType.Text.ToUpper() & "','" & txtCustomerCode.Text & "') ORDER BY INVOICE_NO"
                Else
                    strSql = "SELECT INVOICE_NO,INVOICE_DATE,VOUCHER_NO FROM DBO.UDF_eINVOICE_MANUAL_SUBMISSION_INVOICE_HELP('" & gstrUnitId & "','" & dtpDateFrom.Value.ToString("dd MMM yyyy") & "','" & dtpDateTo.Value.ToString("dd MMM yyyy") & "','" & cmbInvoiceType.Text.ToUpper() & "','" & txtCustomerCode.Text & "') ORDER BY INVOICE_NO"
                End If
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
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtInvoiceTo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtInvoiceTo.KeyDown
        Dim objText As TextBox = DirectCast(sender, TextBox)
        If e.KeyCode = Keys.Delete Then
            objText.Text = String.Empty
        ElseIf e.KeyCode = Keys.F1 Then
            cmdInvoiceTo_Click(sender, e)
        End If
    End Sub

    Private Sub dtpDateFrom_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDateFrom.ValueChanged
        Try
            txtInvoiceFrom.Text = String.Empty
            txtInvoiceTo.Text = String.Empty
            fspInvoices.MaxRows = 0
            rdbCheckAll.Checked = False
            rdbUncheckAll.Checked = False
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub dtpDateTo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDateTo.ValueChanged
        Try
            txtInvoiceFrom.Text = String.Empty
            txtInvoiceTo.Text = String.Empty
            fspInvoices.MaxRows = 0
            rdbCheckAll.Checked = False
            rdbUncheckAll.Checked = False
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtCustomerCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Try
            Dim objText As TextBox = DirectCast(sender, TextBox)
            If e.KeyCode = Keys.Delete Then
                objText.Text = String.Empty
                If objText.Name.ToUpper = "txtCustomerCode".ToUpper Then
                    lblCustCodeDes.Text = String.Empty
                End If
            ElseIf e.KeyCode = Keys.F1 Then
                CmdCustCodeHelp_Click(sender, e)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtCustomerCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustomerCode.TextChanged
        Try
            txtInvoiceFrom.Text = String.Empty
            txtInvoiceTo.Text = String.Empty
            fspInvoices.MaxRows = 0
            rdbCheckAll.Checked = False
            rdbUncheckAll.Checked = False
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtInvoiceFrom_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtInvoiceFrom.TextChanged
        Try
            txtInvoiceTo.Text = String.Empty
            fspInvoices.MaxRows = 0
            rdbCheckAll.Checked = False
            rdbUncheckAll.Checked = False
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FillInvoices()
        Dim dt As New DataTable
        Try
            Dim sqlCmd As New SqlCommand
            With sqlCmd
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 300 ' 5 Minute
                .CommandText = "USP_eINVOICE_MANUAL_SUBMISSION"
                .Parameters.Clear()
                .Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
                .Parameters.AddWithValue("@CUSTOMER_CODE", txtCustomerCode.Text.Trim())
                .Parameters.AddWithValue("@FROM_INVOICE", txtInvoiceFrom.Text)
                .Parameters.AddWithValue("@TO_INVOICE", txtInvoiceTo.Text)
                .Parameters.AddWithValue("@INVOICE_TYPE", cmbInvoiceType.Text.ToUpper())
                .Parameters.AddWithValue("@OPERATION_FLAG", "GET_INVOICES")
                .Parameters.AddWithValue("@USER_ID", mP_User)
                .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                dt = SqlConnectionclass.GetDataTable(sqlCmd)
                If Convert.ToString(.Parameters("@MESSAGE").Value) <> "" Then
                    MsgBox(Convert.ToString(.Parameters("@MESSAGE").Value), MsgBoxStyle.Exclamation, ResolveResString(100))
                Else
                    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                        InitializeSpread()
                        For i As Integer = 0 To dt.Rows.Count - 1
                            With fspInvoices
                                AddRow()
                                .SetText(enumBillDetail.TICK, i + 1, False)
                                .SetText(enumBillDetail.DOC_NO, i + 1, dt.Rows(i).Item("DOC_NO").ToString.Trim)
                                .SetText(enumBillDetail.DOC_DATE, i + 1, dt.Rows(i).Item("INVOICE_DATE"))
                                If cmbInvoiceType.Text.ToUpper() <> FRMMKTTRN0106.Normal Then
                                    .SetText(enumBillDetail.VOUCHER_NO, i + 1, dt.Rows(i).Item("VOUCHER_NO"))
                                    .SetText(enumBillDetail.DRCR, i + 1, dt.Rows(i).Item("DRCR"))
                                    .SetText(enumBillDetail.SUPP_INV_TYPE, i + 1, dt.Rows(i).Item("SUPPINVTYPE"))
                                End If
                            End With
                        Next
                    End If
                End If
            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            dt.Dispose()
        End Try
    End Sub

    Private Sub InitializeSpread()
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            With Me.fspInvoices
                .MaxRows = 0
                .MaxCols = [Enum].GetValues(GetType(enumBillDetail)).Length
                .set_RowHeight(0, 20)
                .Row = 0 : .Col = enumBillDetail.TICK : .Text = "Select" : .set_ColWidth(.Col, 5) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBillDetail.DOC_NO : .Text = "Invoice No." : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBillDetail.DOC_DATE : .Text = "Invoice Date" : .set_ColWidth(.Col, 12) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                If cmbInvoiceType.Text.ToUpper() = FRMMKTTRN0106.Normal Then
                    .Row = 0 : .Col = enumBillDetail.VOUCHER_NO : .Text = "Voucher No." : .set_ColWidth(.Col, 12) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .ColHidden = True
                    .Row = 0 : .Col = enumBillDetail.DRCR : .Text = "DR/CR" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .ColHidden = True
                    .Row = 0 : .Col = enumBillDetail.SUPP_INV_TYPE : .Text = "Invoice Type" : .set_ColWidth(.Col, 12) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .ColHidden = True
                Else
                    .Row = 0 : .Col = enumBillDetail.VOUCHER_NO : .Text = "Voucher No." : .set_ColWidth(.Col, 12) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .ColHidden = False
                    .Row = 0 : .Col = enumBillDetail.DRCR : .Text = "DR/CR" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .ColHidden = False
                    .Row = 0 : .Col = enumBillDetail.SUPP_INV_TYPE : .Text = "Invoice Type" : .set_ColWidth(.Col, 12) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .ColHidden = False
                End If
                .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            End With
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Public Sub AddRow()
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            With fspInvoices
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows : .Col = enumBillDetail.TICK : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumBillDetail.DOC_NO : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumBillDetail.DOC_DATE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeDate : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .TypeDateFormat = FPSpreadADO.TypeDateFormatConstants.TypeDateFormatDDMMYY
                .Row = .MaxRows : .Col = enumBillDetail.VOUCHER_NO : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .TypeMaxEditLen = 15
                .Row = .MaxRows : .Col = enumBillDetail.DRCR : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumBillDetail.SUPP_INV_TYPE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
            End With
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnResubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnResubmit.Click
        Try
            If ValidateInvoices() Then
                Exit Sub
            End If
            If MessageBox.Show("Are you sure you want to Re-Submit selected Invoices?", "eMPro", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = Windows.Forms.DialogResult.Yes Then
                Resubmission()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub Resubmission()
        Dim dtInvoices As New DataTable
        Try
            dtInvoices.Columns.Add("DOC_NO", GetType(System.Int64))
            dtInvoices.Columns.Add("VOUCHER_NO", GetType(System.String))
            dtInvoices.Columns.Add("UNIT_CODE", GetType(System.String))
            Dim drInvoices As DataRow
            Dim isChecked As String = String.Empty
            With fspInvoices
                For i As Integer = 1 To .MaxRows
                    .Row = i
                    .Col = enumBillDetail.TICK
                    isChecked = Convert.ToString(.Value)
                    If isChecked = "1" Then
                        drInvoices = dtInvoices.NewRow()
                        .Row = i
                        .Col = enumBillDetail.DOC_NO
                        drInvoices("DOC_NO") = Convert.ToInt64(.Text)
                        If cmbInvoiceType.Text.ToUpper() <> FRMMKTTRN0106.Normal Then
                            .Row = i
                            .Col = enumBillDetail.VOUCHER_NO
                            drInvoices("VOUCHER_NO") = .Text
                        Else
                            drInvoices("VOUCHER_NO") = ""
                        End If
                        drInvoices("UNIT_CODE") = gstrUnitId
                        dtInvoices.Rows.Add(drInvoices)
                    End If
                Next
            End With
            If dtInvoices IsNot Nothing AndAlso dtInvoices.Rows.Count > 0 Then
                Dim sqlCmd As New SqlCommand
                With sqlCmd
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 300 ' 5 Minute
                    .CommandText = "USP_eINVOICE_MANUAL_SUBMISSION"
                    .Parameters.Clear()
                    .Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
                    .Parameters.AddWithValue("@INVOICE_TYPE", cmbInvoiceType.Text.ToUpper())
                    .Parameters.AddWithValue("@OPERATION_FLAG", "SUBMIT")
                    .Parameters.AddWithValue("@USER_ID", mP_User)
                    .Parameters.AddWithValue("@UDT_EINVOICE_MANUAL_SUBMISSION", dtInvoices)
                    .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                    SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                    If Convert.ToString(.Parameters("@MESSAGE").Value) <> "" Then
                        MsgBox(Convert.ToString(.Parameters("@MESSAGE").Value), MsgBoxStyle.Exclamation, ResolveResString(100))
                    Else
                        MsgBox("Selected Invoices Resubmit successfully.")
                        btnCancel.PerformClick()
                    End If
                End With
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            dtInvoices.Dispose()
        End Try
    End Sub

    Private Function ValidateInvoices() As Boolean
        Dim result As Boolean = False
        Dim chkSelect As String = String.Empty
        Dim countCheck As Integer = 0
        Dim transportMode As String = String.Empty
        Try
            If fspInvoices Is Nothing OrElse fspInvoices.MaxRows = 0 Then
                MessageBox.Show("No Invoice for resubmission.Please first search invoices.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                result = True
                Return result
            End If
            With fspInvoices
                For i As Integer = 1 To .MaxRows
                    .Row = i
                    .Col = enumBillDetail.TICK
                    chkSelect = Convert.ToString(.Value)
                    If chkSelect = "1" Then
                        countCheck += 1
                        Exit For
                    End If
                Next
                If countCheck = 0 Then
                    MessageBox.Show("Please select atleast one invoice for resubmission.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    result = True
                    Return result
                End If
            End With
            Return result
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function

    Private Sub cmbProcessType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbInvoiceType.SelectedIndexChanged
        Try
            txtCustomerCode.Text = String.Empty
            lblCustCodeDes.Text = String.Empty
            txtInvoiceFrom.Text = String.Empty
            txtInvoiceTo.Text = String.Empty
            fspInvoices.MaxRows = 0
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtInvoiceTo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtInvoiceTo.TextChanged
        Try
            fspInvoices.MaxRows = 0
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub rdbCheckAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbCheckAll.CheckedChanged
        Try
            CheckUncheckAll("1")
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub rdbUncheckAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbUncheckAll.CheckedChanged
        Try
            CheckUncheckAll("0")
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub CheckUncheckAll(ByVal tick As String)
        If fspInvoices Is Nothing OrElse fspInvoices.MaxRows = 0 Then
            rdbCheckAll.Checked = False
            rdbUncheckAll.Checked = False
            Exit Sub
        End If
        With fspInvoices
            For i As Integer = 1 To .MaxRows
                .Row = i
                .Col = enumBillDetail.TICK
                .Value = tick
            Next
        End With
    End Sub
End Class