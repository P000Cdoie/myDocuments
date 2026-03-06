Imports System
Imports System.Globalization
Imports System.Data.SqlClient
'*********************************************************************************************************************
'Copyright(c)       - MIND
'Name of Module     - e-Way Bill Detail
'Name of Form       - FRMMKTTRN0102  , e-Way Bill Detail
'Created by         - Ashish sharma
'Created Date       - 31 JAN 2018
'description        - Update e-Way Bill Detail (New Development)
'*********************************************************************************************************************
'Created by         - Ashish sharma
'Created Date       - 14 AUG 2020
'description        - 102027599 - IRN CHANGES
'*********************************************************************************************************************

Public Class FRMMKTTRN0102
    Private Const Invoice As String = "INVOICES"
    Private Enum enumBillDetail
        TICK = 1
        DOC_NO
        VOUCHER_NO
        DR_CR
        CUSTOMER_VENDOR_CODE
        EWAY_IRN_NO
        EWAY_IRN_DATE
        IRN_BARCODE_STRING
        EWAY_IRN_CANCEL
        EWAY_IRN_CANCEL_DATE
        EWAY_BILL_NO
        EWAY_TRANSPORT_MODE
        EWAY_TRANSPORTER_ID
        EWAY_TRANSPORTER_DOC_NO
        EWAY_TRANSPORTER_DOC_DATE
        EWAY_VEHICLE_NO
        EWAY_TYPE
        EWAY_VALID_FROM_DATE
        EWAY_VALID_FROM_TIME
        EWAY_VALID_TO_DATE
        EWAY_VALID_TO_TIME
    End Enum
    Private Sub FRMMKTTRN0102_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        cmbProcessType.Focus()
    End Sub
    Private Sub FRMMKTTRN0102_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Call FitToClient(Me, GrpMain, ctlHeader, GrpBoxButtons, 600)
            Me.MdiParent = mdifrmMain
            dtpDateFrom.Value = GetServerDate()
            dtpDateTo.Value = GetServerDate()
            cmbProcessType.SelectedIndex = 0
            InitializeSpread()
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
            ElseIf String.IsNullOrEmpty(cmbProcessType.Text) Then
                MessageBox.Show("Please first select [Process Type].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                cmbProcessType.Focus()
                Exit Sub
            Else
                If cmbProcessType.Text.ToUpper() = FRMMKTTRN0102.Invoice Then
                    strSql = " SELECT CAST(Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),Invoice_Date,103) as Invoice_Date" & _
                             " FROM SALESCHALLAN_DTL " & _
                             " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtpDateFrom.Text & "',103) AND Convert(Date,'" & dtpDateTo.Text & "',103)" & _
                             " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 " & _
                             " AND UNIT_CODE = '" & gstrUnitId & "' "
                    If Not String.IsNullOrEmpty(txtCustomerCode.Text) Then
                        strSql = strSql & " AND Account_Code='" & txtCustomerCode.Text.Trim() & "'"
                    End If
                    strSql = strSql & " ORDER BY Invoice_No"
                    strHelp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "Invoice No. Help")
                Else
                    strSql = " SELECT CAST(Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),Invoice_Date,103) as Invoice_Date,Voucher_No" & _
                            " FROM SUPPLEMENTARYINV_HDR (NOLOCK) " & _
                            " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtpDateFrom.Text & "',103) AND Convert(Date,'" & dtpDateTo.Text & "',103)" & _
                            " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 " & _
                            " AND UNIT_CODE = '" & gstrUnitId & "' "
                    If Not String.IsNullOrEmpty(txtCustomerCode.Text) Then
                        strSql = strSql & " AND Account_Code='" & txtCustomerCode.Text.Trim() & "'"
                    End If
                    strSql = strSql & " ORDER BY Invoice_No"
                    strHelp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "Supp .Invoice No. Help")
                End If
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
        If e.KeyCode = Keys.F1 Then
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
                ElseIf String.IsNullOrEmpty(cmbProcessType.Text) Then
                    MessageBox.Show("Please first select [Process Type].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    cmbProcessType.Focus()
                    Exit Sub
                Else
                    strSql = " Select CUSTOMER_CODE [CUSTOMER_VENDOR_CODE],CUST_NAME [CUSTOMER_VENDOR_NAME] FROM CUSTOMER_VENDOR_VW WHERE UNIT_CODE='" & gstrUnitId & "'"
                    strHelp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "Customer/Vendor Help")
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

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Try
            If dtpDateFrom.Value > dtpDateTo.Value Then
                MessageBox.Show("[Date From] should be less than or equal to [Date To].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                dtpDateFrom.Focus()
                Exit Sub
            ElseIf String.IsNullOrEmpty(cmbProcessType.Text) Then
                MessageBox.Show("Please first select [Process Type].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                cmbProcessType.Focus()
                Exit Sub
            ElseIf Not String.IsNullOrEmpty(txtInvoiceFrom.Text) And String.IsNullOrEmpty(txtInvoiceTo.Text) Then
                MessageBox.Show("Please first enter [Invoice To].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtInvoiceTo.Focus()
                Exit Sub
            ElseIf Not String.IsNullOrEmpty(txtInvoiceTo.Text) And String.IsNullOrEmpty(txtInvoiceFrom.Text) Then
                MessageBox.Show("Please first enter [Invoice From].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
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
        Me.Close()
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Try
            'cmbProcessType.SelectedIndex = 0
            'dtpDateFrom.Value = GetServerDate()
            'dtpDateTo.Value = GetServerDate()
            txtCustomerCode.Text = String.Empty
            lblCustCodeDes.Text = String.Empty
            txtInvoiceFrom.Text = String.Empty
            txtInvoiceTo.Text = String.Empty
            fspEWayBillDetail.MaxRows = 0
            cmbProcessType.Focus()
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
            ElseIf String.IsNullOrEmpty(cmbProcessType.Text) Then
                MessageBox.Show("Please first select [Process Type].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                cmbProcessType.Focus()
                Exit Sub
            ElseIf String.IsNullOrEmpty(txtInvoiceFrom.Text) Then
                MessageBox.Show("Please first select [Invoice From].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtInvoiceFrom.Focus()
                Exit Sub
            Else
                If cmbProcessType.Text.ToUpper() = FRMMKTTRN0102.Invoice Then
                    strSql = " SELECT CAST(Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),Invoice_Date,103) as Invoice_Date" & _
                             " FROM SALESCHALLAN_DTL " & _
                             " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtpDateFrom.Text & "',103) AND Convert(Date,'" & dtpDateTo.Text & "',103)" & _
                             " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 " & _
                             " AND UNIT_CODE = '" & gstrUnitId & "' "
                    If Not String.IsNullOrEmpty(txtCustomerCode.Text) Then
                        strSql = strSql & " AND Account_Code='" & txtCustomerCode.Text.Trim() & "'"
                    End If
                    strSql = strSql & " ORDER BY Invoice_No"
                    strHelp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "Invoice No. Help")
                Else
                    strSql = " SELECT CAST(Doc_No AS VARCHAR(20)) As Invoice_No,CONVERT(VARCHAR(11),Invoice_Date,103) as Invoice_Date,Voucher_No" & _
                            " FROM SUPPLEMENTARYINV_HDR (NOLOCK) " & _
                            " WHERE Convert(Date,Invoice_Date,103) BETWEEN Convert(Date,'" & dtpDateFrom.Text & "',103) AND Convert(Date,'" & dtpDateTo.Text & "',103)" & _
                            " AND Bill_Flag = 1 And ISNULL(CANCEL_FLAG, 0) = 0 " & _
                            " AND UNIT_CODE = '" & gstrUnitId & "' "
                    If Not String.IsNullOrEmpty(txtCustomerCode.Text) Then
                        strSql = strSql & " AND Account_Code='" & txtCustomerCode.Text.Trim() & "'"
                    End If
                    strSql = strSql & " ORDER BY Invoice_No"
                    strHelp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "Supp .Invoice No. Help")
                End If
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
        If e.KeyCode = Keys.F1 Then
            cmdInvoiceTo_Click(sender, e)
        End If
    End Sub

    Private Sub dtpDateFrom_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDateFrom.ValueChanged
        Try
            txtInvoiceFrom.Text = String.Empty
            txtInvoiceTo.Text = String.Empty
            fspEWayBillDetail.MaxRows = 0
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub dtpDateTo_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpDateTo.ValueChanged
        Try
            txtInvoiceFrom.Text = String.Empty
            txtInvoiceTo.Text = String.Empty
            fspEWayBillDetail.MaxRows = 0
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtCustomerCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown, txtInvoiceFrom.KeyDown, txtInvoiceTo.KeyDown
        Try
            Dim objText As TextBox = DirectCast(sender, TextBox)
            If e.KeyCode = Keys.Delete Then
                objText.Text = String.Empty
                If objText.Name.ToUpper = "txtCustomerCode".ToUpper Then
                    lblCustCodeDes.Text = String.Empty
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtCustomerCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustomerCode.TextChanged
        Try
            txtInvoiceFrom.Text = String.Empty
            txtInvoiceTo.Text = String.Empty
            fspEWayBillDetail.MaxRows = 0
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtInvoiceFrom_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtInvoiceFrom.TextChanged
        Try
            txtInvoiceTo.Text = String.Empty
            fspEWayBillDetail.MaxRows = 0
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
                .CommandText = "USP_UPDATE_EWAY_DETAIL_IN_INVOICE"
                .Parameters.Clear()
                .Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
                If Not String.IsNullOrEmpty(txtCustomerCode.Text) Then
                    .Parameters.AddWithValue("@ACCOUNT_CODE", txtCustomerCode.Text.Trim())
                End If
                .Parameters.AddWithValue("@DATE_FROM", dtpDateFrom.Value)
                .Parameters.AddWithValue("@DATE_To", dtpDateTo.Value)
                If Not String.IsNullOrEmpty(txtInvoiceFrom.Text) Then
                    .Parameters.AddWithValue("@INVOICE_FROM", txtInvoiceFrom.Text)
                End If
                If Not String.IsNullOrEmpty(txtInvoiceTo.Text) Then
                    .Parameters.AddWithValue("@INVOICE_TO", txtInvoiceTo.Text)
                End If
                If cmbProcessType.Text.ToUpper() = Me.Invoice Then
                    .Parameters.AddWithValue("@OPERATION_FLAG", "GET_INVOICE")
                Else
                    .Parameters.AddWithValue("@OPERATION_FLAG", "GET_SUPP_INVOICE")
                End If
                .Parameters.Add("@MSG", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                dt = SqlConnectionclass.GetDataTable(sqlCmd)
                If Convert.ToString(.Parameters("@MSG").Value) <> "" Then
                    MsgBox(Convert.ToString(.Parameters("@MSG").Value), MsgBoxStyle.Exclamation, ResolveResString(100))
                Else
                    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                        InitializeSpread()
                        For i As Integer = 0 To dt.Rows.Count - 1
                            With fspEWayBillDetail
                                AddRow()
                                .SetText(enumBillDetail.TICK, i + 1, False)
                                .SetText(enumBillDetail.DOC_NO, i + 1, dt.Rows(i).Item("DOC_NO").ToString.Trim)
                                .SetText(enumBillDetail.CUSTOMER_VENDOR_CODE, i + 1, dt.Rows(i).Item("ACCOUNT_CODE"))
                                .SetText(enumBillDetail.EWAY_IRN_NO, i + 1, dt.Rows(i).Item("IRN_NO"))
                                If Not IsDBNull(dt.Rows(i).Item("IRN_DATE")) Then
                                    .SetText(enumBillDetail.EWAY_IRN_DATE, i + 1, dt.Rows(i).Item("IRN_DATE"))
                                End If
                                .SetText(enumBillDetail.IRN_BARCODE_STRING, i + 1, dt.Rows(i).Item("BARCODE_DATA"))
                                .SetText(enumBillDetail.EWAY_IRN_CANCEL, i + 1, dt.Rows(i).Item("IRN_CANCEL"))
                                If Not IsDBNull(dt.Rows(i).Item("IRN_CANCEL_DATE")) Then
                                    .SetText(enumBillDetail.EWAY_IRN_CANCEL_DATE, i + 1, dt.Rows(i).Item("IRN_CANCEL_DATE"))
                                End If
                                .SetText(enumBillDetail.EWAY_BILL_NO, i + 1, dt.Rows(i).Item("EWAY_BILL_NO"))
                                .SetText(enumBillDetail.EWAY_TRANSPORT_MODE, i + 1, dt.Rows(i).Item("EWAY_TRANSPORT_MODE"))
                                .SetText(enumBillDetail.EWAY_TRANSPORTER_ID, i + 1, dt.Rows(i).Item("EWAY_TRANSPORTER_ID"))
                                .SetText(enumBillDetail.EWAY_TRANSPORTER_DOC_NO, i + 1, dt.Rows(i).Item("EWAY_TRANSPORTER_DOC_NO").ToString.Trim)
                                If Not IsDBNull(dt.Rows(i).Item("EWAY_TRANSPORTER_DOC_DATE")) Then
                                    .SetText(enumBillDetail.EWAY_TRANSPORTER_DOC_DATE, i + 1, dt.Rows(i).Item("EWAY_TRANSPORTER_DOC_DATE"))
                                End If
                                .SetText(enumBillDetail.EWAY_VEHICLE_NO, i + 1, dt.Rows(i).Item("EWAY_VEHICLE_NO").ToString.Trim)

                                .SetText(enumBillDetail.EWAY_TYPE, i + 1, dt.Rows(i).Item("EWAY_TYPE").ToString.Trim)

                                If dt.Rows(i).Item("EWAY_TYPE").ToString.Trim.ToUpper() = "B" Then
                                    If Not IsDBNull(dt.Rows(i).Item("EWAY_UPD_DT")) Then
                                        .SetText(enumBillDetail.EWAY_VALID_FROM_DATE, i + 1, dt.Rows(i).Item("EWAY_UPD_DT"))
                                        .SetText(enumBillDetail.EWAY_VALID_FROM_TIME, i + 1, Convert.ToDateTime(dt.Rows(i).Item("EWAY_UPD_DT")).ToString("hh:mm tt"))
                                    End If
                                    If Not IsDBNull(dt.Rows(i).Item("EWAY_EXP_DT")) Then
                                        .SetText(enumBillDetail.EWAY_VALID_TO_DATE, i + 1, dt.Rows(i).Item("EWAY_EXP_DT"))
                                        .SetText(enumBillDetail.EWAY_VALID_TO_TIME, i + 1, Convert.ToDateTime(dt.Rows(i).Item("EWAY_EXP_DT")).ToString("hh:mm tt"))
                                    End If
                                End If
                                If cmbProcessType.Text.ToUpper() <> Me.Invoice Then
                                    .SetText(enumBillDetail.VOUCHER_NO, i + 1, dt.Rows(i).Item("VOUCHER_NO").ToString.Trim)
                                    .SetText(enumBillDetail.DR_CR, i + 1, dt.Rows(i).Item("DRCR"))
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
            With Me.fspEWayBillDetail
                .MaxRows = 0
                .MaxCols = [Enum].GetValues(GetType(enumBillDetail)).Length
                .set_RowHeight(0, 20)
                .Row = 0 : .Col = enumBillDetail.TICK : .Text = "Select" : .set_ColWidth(.Col, 5) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBillDetail.DOC_NO : .Text = "Doc No." : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                If cmbProcessType.Text.ToUpper() = FRMMKTTRN0102.Invoice Then
                    .Row = 0 : .Col = enumBillDetail.VOUCHER_NO : .Text = "Voucher No." : .set_ColWidth(.Col, 12) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .ColHidden = True
                    .Row = 0 : .Col = enumBillDetail.DR_CR : .Text = "DR/CR" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .ColHidden = True
                Else
                    .Row = 0 : .Col = enumBillDetail.VOUCHER_NO : .Text = "Voucher No." : .set_ColWidth(.Col, 12) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .ColHidden = False
                    .Row = 0 : .Col = enumBillDetail.DR_CR : .Text = "DR/CR" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .ColHidden = False
                End If
                .Row = 0 : .Col = enumBillDetail.CUSTOMER_VENDOR_CODE : .Text = "Customer/Vendor Code" : .set_ColWidth(.Col, 12) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBillDetail.EWAY_IRN_NO : .Text = "IRN No." : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBillDetail.EWAY_IRN_DATE : .Text = "IRN Date" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBillDetail.IRN_BARCODE_STRING : .Text = "IRN Barcode" : .set_ColWidth(.Col, 15) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBillDetail.EWAY_IRN_CANCEL : .Text = "IRN Cancel" : .set_ColWidth(.Col, 12) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBillDetail.EWAY_IRN_CANCEL_DATE : .Text = "IRN Cancel Date" : .set_ColWidth(.Col, 12) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBillDetail.EWAY_BILL_NO : .Text = "e-Way Bill No." : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBillDetail.EWAY_TRANSPORT_MODE : .Text = "Transport Mode(F1)" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBillDetail.EWAY_TRANSPORTER_ID : .Text = "Transporter ID(F1)" : .set_ColWidth(.Col, 12) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBillDetail.EWAY_TRANSPORTER_DOC_NO : .Text = "Transporter Doc No." : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBillDetail.EWAY_TRANSPORTER_DOC_DATE : .Text = "Transporter Doc Dt." : .set_ColWidth(.Col, 12) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBillDetail.EWAY_VEHICLE_NO : .Text = "Vehicle No." : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBillDetail.EWAY_TYPE : .Text = "Type" : .set_ColWidth(.Col, 8) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBillDetail.EWAY_VALID_FROM_DATE : .Text = "Valid From Date" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBillDetail.EWAY_VALID_FROM_TIME : .Text = "Valid From Time" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBillDetail.EWAY_VALID_TO_DATE : .Text = "Valid To Date" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBillDetail.EWAY_VALID_TO_TIME : .Text = "Valid To Time" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
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
            With fspEWayBillDetail
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows : .Col = enumBillDetail.TICK : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumBillDetail.DOC_NO : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumBillDetail.VOUCHER_NO : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumBillDetail.DR_CR : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumBillDetail.CUSTOMER_VENDOR_CODE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumBillDetail.EWAY_IRN_NO : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .TypeMaxEditLen = 64
                .Row = .MaxRows : .Col = enumBillDetail.EWAY_IRN_DATE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeDate : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .TypeDateFormat = FPSpreadADO.TypeDateFormatConstants.TypeDateFormatDDMMYY
                .Row = .MaxRows : .Col = enumBillDetail.IRN_BARCODE_STRING : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .TypeMaxEditLen = 2000
                .Row = .MaxRows : .Col = enumBillDetail.EWAY_IRN_CANCEL : .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .TypeComboBoxEditable = False : .TypeComboBoxList = "" & Chr(9) & "NO" & Chr(9) & "YES" : .TypeComboBoxCurSel = 0
                .Row = .MaxRows : .Col = enumBillDetail.EWAY_IRN_CANCEL_DATE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeDate : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .TypeDateFormat = FPSpreadADO.TypeDateFormatConstants.TypeDateFormatDDMMYY
                .Row = .MaxRows : .Col = enumBillDetail.EWAY_BILL_NO : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .TypeMaxEditLen = 15
                .Row = .MaxRows : .Col = enumBillDetail.EWAY_TRANSPORT_MODE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumBillDetail.EWAY_TRANSPORTER_ID : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumBillDetail.EWAY_TRANSPORTER_DOC_NO : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .TypeMaxEditLen = 15
                .Row = .MaxRows : .Col = enumBillDetail.EWAY_TRANSPORTER_DOC_DATE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeDate : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .TypeDateFormat = FPSpreadADO.TypeDateFormatConstants.TypeDateFormatDDMMYY
                .Row = .MaxRows : .Col = enumBillDetail.EWAY_VEHICLE_NO : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .TypeMaxEditLen = 10
                .Row = .MaxRows : .Col = enumBillDetail.EWAY_TYPE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .TypeComboBoxEditable = False : .TypeComboBoxList = "" & Chr(9) & "A" & Chr(9) & "B" : .TypeComboBoxCurSel = 0
                .Row = .MaxRows : .Col = enumBillDetail.EWAY_VALID_FROM_DATE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeDate : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .TypeDateFormat = FPSpreadADO.TypeDateFormatConstants.TypeDateFormatDDMMYY
                .Row = .MaxRows : .Col = enumBillDetail.EWAY_VALID_FROM_TIME : .CellType = FPSpreadADO.CellTypeConstants.CellTypeTime : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .TypeTime24Hour = FPSpreadADO.TypeTime24HourConstants.TypeTime24Hour12HourClock
                .Row = .MaxRows : .Col = enumBillDetail.EWAY_VALID_TO_DATE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeDate : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .TypeDateFormat = FPSpreadADO.TypeDateFormatConstants.TypeDateFormatDDMMYY
                .Row = .MaxRows : .Col = enumBillDetail.EWAY_VALID_TO_TIME : .CellType = FPSpreadADO.CellTypeConstants.CellTypeTime : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .TypeTime24Hour = FPSpreadADO.TypeTime24HourConstants.TypeTime24Hour12HourClock
            End With
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub fspEWayBillDetail_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles fspEWayBillDetail.ButtonClicked
        Try
            Dim isChecked As String = String.Empty
            With fspEWayBillDetail
                If e.col = enumBillDetail.TICK Then
                    .Row = .ActiveRow
                    .Col = .ActiveCol
                    isChecked = Convert.ToString(.Value)
                    If isChecked <> "1" Then
                        .Row = .ActiveRow
                        .Col = enumBillDetail.EWAY_IRN_NO
                        .Lock = True
                        .Col = enumBillDetail.EWAY_IRN_DATE
                        .Lock = True
                        .Col = enumBillDetail.IRN_BARCODE_STRING
                        .Lock = True
                        .Col = enumBillDetail.EWAY_IRN_CANCEL
                        .Lock = True
                        .Col = enumBillDetail.EWAY_IRN_CANCEL_DATE
                        .Lock = True
                        .Col = enumBillDetail.EWAY_BILL_NO
                        .Lock = True
                        .Col = enumBillDetail.EWAY_TRANSPORT_MODE
                        .Lock = True
                        .Col = enumBillDetail.EWAY_TRANSPORTER_ID
                        .Lock = True
                        .Col = enumBillDetail.EWAY_TRANSPORTER_DOC_NO
                        .Lock = True
                        .Col = enumBillDetail.EWAY_TRANSPORTER_DOC_DATE
                        .Lock = True
                        .Col = enumBillDetail.EWAY_VEHICLE_NO
                        .Lock = True
                        .Col = enumBillDetail.EWAY_TYPE
                        .Lock = True
                        .Col = enumBillDetail.EWAY_VALID_FROM_DATE
                        .Lock = True
                        .Col = enumBillDetail.EWAY_VALID_FROM_TIME
                        .Lock = True
                        .Col = enumBillDetail.EWAY_VALID_TO_DATE
                        .Lock = True
                        .Col = enumBillDetail.EWAY_VALID_TO_TIME
                        .Lock = True
                    Else
                        .Row = .ActiveRow
                        .Col = enumBillDetail.EWAY_IRN_NO
                        .Lock = False
                        .Col = enumBillDetail.EWAY_IRN_DATE
                        .Lock = False
                        .Col = enumBillDetail.IRN_BARCODE_STRING
                        .Lock = False
                        .Col = enumBillDetail.EWAY_IRN_CANCEL
                        .Lock = False
                        .Col = enumBillDetail.EWAY_IRN_CANCEL_DATE
                        .Lock = False
                        If cmbProcessType.Text.ToUpper() = Me.Invoice Then
                            .Col = enumBillDetail.EWAY_BILL_NO
                            .Lock = False
                            .Col = enumBillDetail.EWAY_TRANSPORT_MODE
                            .Lock = False
                            .Col = enumBillDetail.EWAY_TRANSPORTER_ID
                            .Lock = False
                            .Col = enumBillDetail.EWAY_TRANSPORTER_DOC_NO
                            .Lock = False
                            .Col = enumBillDetail.EWAY_TRANSPORTER_DOC_DATE
                            .Lock = False
                            .Col = enumBillDetail.EWAY_VEHICLE_NO
                            .Lock = False
                            .Col = enumBillDetail.EWAY_TYPE
                            .Lock = False
                            .Col = enumBillDetail.EWAY_VALID_FROM_DATE
                            .Lock = False
                            .Col = enumBillDetail.EWAY_VALID_FROM_TIME
                            .Lock = False
                            .Col = enumBillDetail.EWAY_VALID_TO_DATE
                            .Lock = False
                            .Col = enumBillDetail.EWAY_VALID_TO_TIME
                            .Lock = False
                        Else
                            .Col = enumBillDetail.EWAY_BILL_NO
                            .Lock = True
                            .Col = enumBillDetail.EWAY_TRANSPORT_MODE
                            .Lock = True
                            .Col = enumBillDetail.EWAY_TRANSPORTER_ID
                            .Lock = True
                            .Col = enumBillDetail.EWAY_TRANSPORTER_DOC_NO
                            .Lock = True
                            .Col = enumBillDetail.EWAY_TRANSPORTER_DOC_DATE
                            .Lock = True
                            .Col = enumBillDetail.EWAY_VEHICLE_NO
                            .Lock = True
                            .Col = enumBillDetail.EWAY_TYPE
                            .Lock = True
                            .Col = enumBillDetail.EWAY_VALID_FROM_DATE
                            .Lock = True
                            .Col = enumBillDetail.EWAY_VALID_FROM_TIME
                            .Lock = True
                            .Col = enumBillDetail.EWAY_VALID_TO_DATE
                            .Lock = True
                            .Col = enumBillDetail.EWAY_VALID_TO_TIME
                            .Lock = True
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub fspEWayBillDetail_ComboSelChange(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ComboSelChangeEvent) Handles fspEWayBillDetail.ComboSelChange
        Try
            If e.col = enumBillDetail.EWAY_TYPE Then
                With fspEWayBillDetail
                    .Row = e.row
                    .Col = enumBillDetail.EWAY_TYPE
                    If String.IsNullOrEmpty(.Text) OrElse .Text = "A" Then
                        .Col = enumBillDetail.EWAY_VALID_FROM_DATE
                        .Text = ""
                        .Lock = True
                        .Col = enumBillDetail.EWAY_VALID_FROM_TIME
                        .Text = ""
                        .Lock = True
                        .Col = enumBillDetail.EWAY_VALID_TO_DATE
                        .Text = ""
                        .Lock = True
                        .Col = enumBillDetail.EWAY_VALID_TO_TIME
                        .Text = ""
                        .Lock = True
                    ElseIf .Text = "B" Then
                        .Col = enumBillDetail.EWAY_VALID_FROM_DATE
                        .Text = ""
                        .Lock = False
                        .Col = enumBillDetail.EWAY_VALID_FROM_TIME
                        .Text = ""
                        .Lock = False
                        .Col = enumBillDetail.EWAY_VALID_TO_DATE
                        .Text = ""
                        .Lock = False
                        .Col = enumBillDetail.EWAY_VALID_TO_TIME
                        .Text = ""
                        .Lock = False
                    End If
                End With
            ElseIf e.col = enumBillDetail.EWAY_IRN_CANCEL Then
                With fspEWayBillDetail
                    .Row = e.row
                    .Col = enumBillDetail.EWAY_IRN_CANCEL
                    If String.IsNullOrEmpty(.Text) OrElse .Text = "NO" Then
                        .Col = enumBillDetail.EWAY_IRN_CANCEL_DATE
                        .Text = ""
                        .Lock = True
                    Else
                        .Col = enumBillDetail.EWAY_IRN_CANCEL_DATE
                        .Text = ""
                        .Lock = False
                    End If
                End With
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub fspEWayBillDetail_KeyDownEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles fspEWayBillDetail.KeyDownEvent
        Dim strQuery As String
        Dim strHelp() As String
        Dim intLoopCounter As Integer = 0
        Dim isChecked As String = String.Empty
        Try
            If e.keyCode = Keys.F1 Then
                fspEWayBillDetail.Row = fspEWayBillDetail.ActiveRow
                fspEWayBillDetail.Col = enumBillDetail.TICK
                isChecked = Convert.ToString(fspEWayBillDetail.Value)
                If isChecked = "0" Then Exit Sub
                If fspEWayBillDetail.ActiveCol = enumBillDetail.EWAY_TRANSPORT_MODE Then
                    With Me.fspEWayBillDetail
                        intLoopCounter = Me.fspEWayBillDetail.ActiveRow
                        strQuery = "Select Transport_Code,Transport_Mode From VW_EWAY_TRANSPORT_MODE"
                        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Transport Mode(s)")
                        If UBound(strHelp) > 0 Then
                            If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
                                MsgBox("Transport Mode Not Available.", MsgBoxStyle.Information, ResolveResString(100))
                                Exit Sub
                            End If
                            If IsNothing(strHelp) = False Then
                                intLoopCounter = Me.fspEWayBillDetail.ActiveRow
                                .SetText(enumBillDetail.EWAY_TRANSPORT_MODE, intLoopCounter, Trim(strHelp(0)).ToString())
                            End If
                        End If
                    End With
                End If

                If fspEWayBillDetail.ActiveCol = enumBillDetail.EWAY_TRANSPORTER_ID Then
                    With Me.fspEWayBillDetail
                        intLoopCounter = Me.fspEWayBillDetail.ActiveRow
                        strQuery = "SELECT DISTINCT ISNULL(TRANSPORTER_ID ,'') TRANSPORTER_ID,VENDOR_CODE,VENDOR_NAME FROM VENDOR_MST WHERE UNIT_CODE='" & gstrUnitId & "'  AND TRANSPORTER_FLAG=1 AND ACTIVE_FLAG='A' ORDER BY VENDOR_CODE"
                        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Transporter Id")
                        If UBound(strHelp) > 0 Then
                            If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
                                MsgBox("Transporter Id Not Available.", MsgBoxStyle.Information, ResolveResString(100))
                                Exit Sub
                            End If
                            If IsNothing(strHelp) = False Then
                                intLoopCounter = Me.fspEWayBillDetail.ActiveRow
                                .SetText(enumBillDetail.EWAY_TRANSPORTER_ID, intLoopCounter, Trim(strHelp(0)).ToString())
                            End If
                        End If
                    End With
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Try
            If cmbProcessType.Text.ToUpper() = Me.Invoice Then
                UpdateInvoices()
            Else
                UpdateSuppInvoices()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub UpdateInvoices()
        Dim dtInvoices As New DataTable
        Dim validFromDate As DateTime
        Dim validFromTime As DateTime
        Dim validToDate As DateTime
        Dim validToTime As DateTime
        Dim validFrom As DateTime
        Dim validTo As DateTime
        Dim eWayType As String = String.Empty
        Dim chkSelect As String = String.Empty
        Dim irnNo As String = String.Empty
        Dim irnBarcode As String = String.Empty
        Try
            If ValidateInvoices() Then
                Exit Sub
            End If

            With fspEWayBillDetail
                For i As Integer = 1 To .MaxRows
                    .Row = i
                    .Col = enumBillDetail.TICK
                    chkSelect = Convert.ToString(.Value)
                    If chkSelect = "1" Then
                        .Row = i
                        .Col = enumBillDetail.EWAY_IRN_NO
                        irnNo = Convert.ToString(.Text).Trim

                        .Row = i
                        .Col = enumBillDetail.IRN_BARCODE_STRING
                        irnBarcode = Convert.ToString(.Text).Trim

                        If irnNo.Length > 0 Then
                            If irnBarcode.Length = 0 Then
                                If MessageBox.Show("Are you sure to update IRN No. without scan/enter IRN barcode?", "eMPro", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = Windows.Forms.DialogResult.No Then
                                    .Row = i
                                    .Col = enumBillDetail.IRN_BARCODE_STRING
                                    .Text = String.Empty
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Exit Sub
                                Else
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next
            End With

            dtInvoices.Columns.Add("UNIT_CODE", GetType(System.String))
            dtInvoices.Columns.Add("DOC_NO", GetType(System.Int64))
            dtInvoices.Columns.Add("EWAY_BILL_NO", GetType(System.String))
            dtInvoices.Columns.Add("EWAY_TRANSPORT_MODE", GetType(System.String))
            dtInvoices.Columns.Add("EWAY_TRANSPORTER_ID", GetType(System.String))
            dtInvoices.Columns.Add("EWAY_TRANSPORTER_DOC_NO", GetType(System.String))
            dtInvoices.Columns.Add("EWAY_TRANSPORTER_DOC_DATE", GetType(System.DateTime))
            dtInvoices.Columns.Add("EWAY_VEHICLE_NO", GetType(System.String))
            dtInvoices.Columns.Add("EWAY_TYPE", GetType(System.String))
            dtInvoices.Columns.Add("EWAY_VALID_FROM", GetType(System.DateTime))
            dtInvoices.Columns.Add("EWAY_VALID_TO", GetType(System.DateTime))
            dtInvoices.Columns.Add("EWAY_IRN_NO", GetType(System.String))
            dtInvoices.Columns.Add("EWAY_IRN_DATE", GetType(System.DateTime))
            dtInvoices.Columns.Add("IRN_CANCEL", GetType(System.Boolean))
            dtInvoices.Columns.Add("IRN_CANCEL_DATE", GetType(System.DateTime))
            dtInvoices.Columns.Add("IRN_BARCODE_DATA", GetType(System.String))
            Dim drInvoices As DataRow
            Dim isChecked As String = String.Empty
            With fspEWayBillDetail
                For i As Integer = 1 To .MaxRows
                    .Row = i
                    .Col = enumBillDetail.TICK
                    isChecked = Convert.ToString(.Value)
                    If isChecked = "1" Then
                        drInvoices = dtInvoices.NewRow()
                        drInvoices("UNIT_CODE") = gstrUnitId
                        .Row = i
                        .Col = enumBillDetail.DOC_NO
                        drInvoices("DOC_NO") = .Text
                        .Row = i
                        .Col = enumBillDetail.EWAY_BILL_NO
                        drInvoices("EWAY_BILL_NO") = .Text
                        .Row = i
                        .Col = enumBillDetail.EWAY_TRANSPORT_MODE
                        drInvoices("EWAY_TRANSPORT_MODE") = .Text
                        .Row = i
                        .Col = enumBillDetail.EWAY_TRANSPORTER_ID
                        drInvoices("EWAY_TRANSPORTER_ID") = .Text
                        .Row = i
                        .Col = enumBillDetail.EWAY_TRANSPORTER_DOC_NO
                        drInvoices("EWAY_TRANSPORTER_DOC_NO") = .Text
                        .Row = i
                        .Col = enumBillDetail.EWAY_TRANSPORTER_DOC_DATE
                        If Convert.ToString(.Text) <> String.Empty Then
                            drInvoices("EWAY_TRANSPORTER_DOC_DATE") = .Text
                        Else
                            drInvoices("EWAY_TRANSPORTER_DOC_DATE") = DBNull.Value
                        End If
                        .Row = i
                        .Col = enumBillDetail.EWAY_VEHICLE_NO
                        drInvoices("EWAY_VEHICLE_NO") = .Text

                        .Row = i
                        .Col = enumBillDetail.EWAY_TYPE
                        eWayType = .Text
                        drInvoices("EWAY_TYPE") = eWayType

                        If eWayType = "B" Then
                            .Row = i
                            .Col = enumBillDetail.EWAY_VALID_FROM_DATE
                            validFromDate = Convert.ToDateTime(.Text)
                            .Row = i
                            .Col = enumBillDetail.EWAY_VALID_FROM_TIME
                            validFromTime = Convert.ToDateTime(.Text)
                            .Row = i
                            .Col = enumBillDetail.EWAY_VALID_TO_DATE
                            validToDate = Convert.ToDateTime(.Text)
                            .Row = i
                            .Col = enumBillDetail.EWAY_VALID_TO_TIME
                            validToTime = Convert.ToDateTime(.Text)

                            validFrom = New DateTime(validFromDate.Year, validFromDate.Month, validFromDate.Day, validFromTime.Hour, validFromTime.Minute, validFromTime.Second)
                            validTo = New DateTime(validToDate.Year, validToDate.Month, validToDate.Day, validToTime.Hour, validToTime.Minute, validToTime.Second)
                            drInvoices("EWAY_VALID_FROM") = validFrom
                            drInvoices("EWAY_VALID_TO") = validTo
                        Else
                            drInvoices("EWAY_VALID_FROM") = DBNull.Value
                            drInvoices("EWAY_VALID_TO") = DBNull.Value
                        End If

                        .Row = i
                        .Col = enumBillDetail.EWAY_IRN_NO
                        If Convert.ToString(.Text).Trim <> "" Then
                            drInvoices("EWAY_IRN_NO") = Convert.ToString(.Text).Trim

                            .Row = i
                            .Col = enumBillDetail.EWAY_IRN_DATE
                            drInvoices("EWAY_IRN_DATE") = Convert.ToDateTime(.Text)

                            .Row = i
                            .Col = enumBillDetail.IRN_BARCODE_STRING
                            drInvoices("IRN_BARCODE_DATA") = Convert.ToString(.Text).Trim

                            .Row = i
                            .Col = enumBillDetail.EWAY_IRN_CANCEL
                            drInvoices("IRN_CANCEL") = IIf(.Text.ToUpper().Trim = "YES", True, False)

                            If .Text.ToUpper().Trim = "YES" Then
                                .Row = i
                                .Col = enumBillDetail.EWAY_IRN_CANCEL_DATE
                                drInvoices("IRN_CANCEL_DATE") = Convert.ToDateTime(.Text)
                            Else
                                drInvoices("IRN_CANCEL_DATE") = DBNull.Value
                            End If

                        Else
                            drInvoices("EWAY_IRN_NO") = DBNull.Value
                            drInvoices("EWAY_IRN_DATE") = DBNull.Value
                            drInvoices("IRN_BARCODE_DATA") = DBNull.Value
                            drInvoices("IRN_CANCEL") = DBNull.Value
                            drInvoices("IRN_CANCEL_DATE") = DBNull.Value
                        End If

                        dtInvoices.Rows.Add(drInvoices)
                    End If
                Next
            End With
            If dtInvoices IsNot Nothing AndAlso dtInvoices.Rows.Count > 0 Then
                Dim sqlCmd As New SqlCommand
                With sqlCmd
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 300 ' 5 Minute
                    .CommandText = "USP_UPDATE_EWAY_DETAIL_IN_INVOICE"
                    .Parameters.Clear()
                    .Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
                    .Parameters.AddWithValue("@DATE_FROM", dtpDateFrom.Value)
                    .Parameters.AddWithValue("@DATE_To", dtpDateTo.Value)
                    .Parameters.AddWithValue("@OPERATION_FLAG", "UPDATE_INVOICE")
                    .Parameters.AddWithValue("@TYPE_EWAY_BILL_UPDATE", dtInvoices)
                    .Parameters.AddWithValue("@USER_ID", mP_User)
                    .Parameters.Add("@MSG", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                    SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                    If Convert.ToString(.Parameters("@MSG").Value) <> "" Then
                        MsgBox(Convert.ToString(.Parameters("@MSG").Value), MsgBoxStyle.Exclamation, ResolveResString(100))
                    Else
                        MsgBox("Selected Invoices updated successfully.")
                        btnClear.PerformClick()
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
        Dim chkSelectInner As String = String.Empty
        Dim countCheck As Integer = 0
        Dim transportMode As String = String.Empty
        Dim validFromDate As DateTime
        Dim validFromTime As DateTime
        Dim validToDate As DateTime
        Dim validToTime As DateTime
        Dim eWayBill_NO As String = String.Empty
        Dim eWayBill_IRN_No As String = String.Empty
        Dim IRN_DATE As DateTime
        Dim IRN_CANCEL_DATE As DateTime
        Dim IRN_CANCEL_FLAG As String = String.Empty
        Dim IRN_BARCODE_STRING As String = String.Empty
        Dim docNo As String = String.Empty
        Dim eWayBill_IRN_No_Inner As String = String.Empty
        Try
            If fspEWayBillDetail Is Nothing OrElse fspEWayBillDetail.MaxRows = 0 Then
                MessageBox.Show("No Invoice for updating.Please first search invoices.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                result = True
                Return result
            End If
            With fspEWayBillDetail
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
                    MessageBox.Show("Please select atleast one invoice for update.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    result = True
                    Return result
                End If

                For i As Integer = 1 To .MaxRows
                    .Row = i
                    .Col = enumBillDetail.TICK
                    chkSelect = Convert.ToString(.Value)
                    If chkSelect = "1" Then
                        .Row = i
                        .Col = enumBillDetail.EWAY_TRANSPORT_MODE
                        transportMode = Convert.ToString(.Text)
                        If String.IsNullOrEmpty(transportMode) Then
                            MessageBox.Show("Please select Transport Mode.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            result = True
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            Return result
                        End If
                        '.Row = i
                        '.Col = enumBillDetail.EWAY_TRANSPORTER_ID
                        'If String.IsNullOrEmpty(.Text) Then
                        '    MessageBox.Show("Please select Transporter ID.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        '    result = True
                        '    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        '    .Focus()
                        '    Return result
                        'End If
                        '.Row = i
                        '.Col = enumBillDetail.EWAY_VEHICLE_NO
                        'If transportMode.ToUpper() = "ROAD" Then
                        '    If String.IsNullOrEmpty(.Text) Then
                        '        MessageBox.Show("Please select Vehicle No.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        '        result = True
                        '        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        '        .Focus()
                        '        Return result
                        '    End If
                        'End If

                        .Row = i
                        .Col = enumBillDetail.DOC_NO
                        docNo = Convert.ToString(.Text).Trim

                        .Row = i
                        .Col = enumBillDetail.IRN_BARCODE_STRING
                        IRN_BARCODE_STRING = Convert.ToString(.Text).Trim

                        .Row = i
                        .Col = enumBillDetail.EWAY_BILL_NO
                        eWayBill_NO = Convert.ToString(.Text)

                        .Row = i
                        .Col = enumBillDetail.EWAY_IRN_NO
                        eWayBill_IRN_No = Convert.ToString(.Text)

                        If eWayBill_NO.Trim = "" And eWayBill_IRN_No.Trim = "" Then
                            MessageBox.Show("Atleast One eWayBill No. or IRN No. Required to Update!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            result = True
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            Return result
                        End If


                        .Row = i
                        .Col = enumBillDetail.EWAY_IRN_DATE
                        If eWayBill_IRN_No.Trim = "" AndAlso Not String.IsNullOrEmpty(.Text) Then
                            MessageBox.Show("IRN DATE Not required for NON IRN No.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            .Text = String.Empty
                            result = True
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            Return result
                        End If
                        If eWayBill_IRN_No.Trim = "" AndAlso Not String.IsNullOrEmpty(IRN_BARCODE_STRING) Then
                            MessageBox.Show("IRN Barcode Not required for NON IRN No.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            .Row = i
                            .Col = enumBillDetail.IRN_BARCODE_STRING
                            .Text = String.Empty
                            result = True
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            Return result
                        End If
                        If eWayBill_IRN_No.Trim <> "" Then
                            If eWayBill_IRN_No.Trim.Length <> 64 Then
                                MessageBox.Show("IRN No. must be of length 64.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                .Row = i
                                .Col = enumBillDetail.EWAY_IRN_NO
                                result = True
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                                Return result
                            End If
                            Dim strSql As String = "SELECT TOP 1 DOC_NO FROM SALESCHALLAN_DTL_IRN (NOLOCK) WHERE UNIT_CODE='" & gstrUnitId & "' AND RTRIM(LTRIM(IRN_NO))='" & eWayBill_IRN_No.Trim & "' AND DOC_NO<>'" & docNo & "'"
                            Dim res As String = Convert.ToString(SqlConnectionclass.ExecuteScalar(strSql))
                            If res.Length > 0 Then
                                MessageBox.Show("Enter IRN No. associated with document no. " & res & vbCrLf & " Please enter unique IRN No.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                .Row = i
                                .Col = enumBillDetail.EWAY_IRN_NO
                                result = True
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                                Return result
                            End If
                        End If
                        If String.IsNullOrEmpty(.Text) And eWayBill_IRN_No.Trim <> "" Then
                            MessageBox.Show("Please enter IRN DATE.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            result = True
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            Return result
                        End If
                        If Not IsDate(.Text) And eWayBill_IRN_No.Trim <> "" Then
                            MessageBox.Show("Please enter valid [IRN Date].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            result = True
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            Return result
                        End If
                        If Not String.IsNullOrEmpty(.Text) Then
                            IRN_DATE = Convert.ToDateTime(.Text)
                        End If

                        If IRN_BARCODE_STRING.Trim <> "" And eWayBill_IRN_No.Trim <> "" Then
                            If IRN_BARCODE_STRING.Trim.Length < 500 Then
                                MessageBox.Show("Please scan/enter valid IRN barcode.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                .Row = i
                                .Col = enumBillDetail.IRN_BARCODE_STRING
                                .Text = String.Empty
                                result = True
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                                Return result
                            End If
                        End If

                        .Row = i
                        .Col = enumBillDetail.EWAY_IRN_CANCEL
                        IRN_CANCEL_FLAG = Convert.ToString(.Text)

                        .Row = i
                        .Col = enumBillDetail.EWAY_IRN_CANCEL_DATE
                        If Not String.IsNullOrEmpty(.Text) And eWayBill_IRN_No.Trim <> "" Then
                            If IRN_CANCEL_FLAG = "NO" And Convert.ToString(.Text) <> "" Then
                                MessageBox.Show("IRN Cancel date not allowed for not Cancel IRN!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                .Text = String.Empty
                                result = True
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                                Return result

                            Else
                                If String.IsNullOrEmpty(.Text) Then
                                    MessageBox.Show("Please enter IRN CANCEL DATE.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    result = True
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Return result
                                End If
                                If Not IsDate(.Text) Then
                                    MessageBox.Show("Please enter valid [IRN CANCEL DATE].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    result = True
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Return result
                                End If

                                IRN_CANCEL_DATE = Convert.ToDateTime(.Text)
                            End If
                        ElseIf IRN_CANCEL_FLAG = "YES" And eWayBill_IRN_No.Trim() <> "" And .Text.ToString() = "" Then

                            MessageBox.Show("Please enter IRN CANCEL DATE.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            result = True
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            Return result


                        End If


                        If IRN_CANCEL_FLAG = "YES" And eWayBill_IRN_No.Trim() = "" Then
                            If Convert.ToString(.Text) = "" Then
                                MessageBox.Show("IRN Cancel date not blank for Canceled IRN!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                result = True
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                                Return result
                            Else
                                MessageBox.Show("No IRN Number found against Cancel!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                result = True
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                                Return result
                            End If
                        End If

                        If eWayBill_NO.Trim().Length > 0 Then
                            .Row = i
                            .Col = enumBillDetail.EWAY_TYPE
                            If String.IsNullOrEmpty(.Text) Then
                                MessageBox.Show("Please select Type.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                result = True
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                                Return result
                            End If
                            If .Text = "B" Then
                                .Col = enumBillDetail.EWAY_VALID_FROM_DATE

                                If String.IsNullOrEmpty(.Text) Then
                                    MessageBox.Show("Please enter Valid From Date.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    result = True
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Return result
                                End If
                                If Not IsDate(.Text) Then
                                    MessageBox.Show("Please enter valid [Valid From Date].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    result = True
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Return result
                                End If
                                validFromDate = Convert.ToDateTime(.Text)
                                .Col = enumBillDetail.EWAY_VALID_FROM_TIME

                                If String.IsNullOrEmpty(.Text) Then
                                    MessageBox.Show("Please enter Valid From Time.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    result = True
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Return result
                                End If
                                If Not IsDate(.Text) Then
                                    MessageBox.Show("Please enter valid [Valid From Time].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    result = True
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Return result
                                End If
                                If Not (.Text.ToUpper().Contains("AM") Or .Text.ToUpper().Contains("PM")) Then
                                    MessageBox.Show("Please enter Valid From Time with AM/PM.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    result = True
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Return result
                                End If
                                validFromTime = Convert.ToDateTime(.Text)
                                .Col = enumBillDetail.EWAY_VALID_TO_DATE

                                If String.IsNullOrEmpty(.Text) Then
                                    MessageBox.Show("Please enter Valid To Date.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    result = True
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Return result
                                End If
                                If Not IsDate(.Text) Then
                                    MessageBox.Show("Please enter valid [Valid To Date].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    result = True
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Return result
                                End If
                                validToDate = Convert.ToDateTime(.Text)
                                .Col = enumBillDetail.EWAY_VALID_TO_TIME

                                If String.IsNullOrEmpty(.Text) Then
                                    MessageBox.Show("Please enter Valid To Time.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    result = True
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Return result
                                End If
                                If Not IsDate(.Text) Then
                                    MessageBox.Show("Please enter valid [Valid To Time].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    result = True
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Return result
                                End If
                                If Not (.Text.ToUpper().Contains("AM") Or .Text.ToUpper().Contains("PM")) Then
                                    MessageBox.Show("Please enter Valid To Time with AM/PM.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    result = True
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Return result
                                End If
                                validToTime = Convert.ToDateTime(.Text)
                                Dim validFrom As DateTime = New DateTime(validFromDate.Year, validFromDate.Month, validFromDate.Day, validFromTime.Hour, validFromTime.Minute, validFromTime.Second)
                                Dim validTo As DateTime = New DateTime(validToDate.Year, validToDate.Month, validToDate.Day, validToTime.Hour, validToTime.Minute, validToTime.Second)
                                .Col = enumBillDetail.EWAY_VALID_TO_DATE
                                If validTo <= validFrom Then
                                    MessageBox.Show("[Valid To Date Time] should be greater than [Valid From Date Time].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    result = True
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Return result
                                End If
                            End If
                        End If
                    End If
                Next
                For i As Integer = 1 To .MaxRows
                    .Row = i
                    .Col = enumBillDetail.TICK
                    chkSelect = Convert.ToString(.Value)
                    If chkSelect = "1" Then
                        .Row = i
                        .Col = enumBillDetail.EWAY_IRN_NO
                        eWayBill_IRN_No = Convert.ToString(.Text)
                        If String.IsNullOrEmpty(eWayBill_IRN_No) Then Continue For

                        For j As Integer = 1 To .MaxRows
                            If i = j Then Continue For

                            .Row = j
                            .Col = enumBillDetail.TICK
                            chkSelectInner = Convert.ToString(.Value)
                            If chkSelectInner = "1" Then
                                .Row = j
                                .Col = enumBillDetail.EWAY_IRN_NO
                                eWayBill_IRN_No_Inner = Convert.ToString(.Text)

                                If String.IsNullOrEmpty(eWayBill_IRN_No_Inner) Then Continue For

                                If eWayBill_IRN_No.Trim = eWayBill_IRN_No_Inner.Trim Then
                                    MessageBox.Show("IRN No. on row no. " & i.ToString & " and row no. " & j.ToString & " is duplicate.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    .Row = i
                                    .Col = enumBillDetail.EWAY_IRN_NO
                                    result = True
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Return result
                                End If
                            End If
                        Next
                    End If
                Next
            End With
        Catch ex As Exception
            result = True
            RaiseException(ex)
        End Try
        Return result
    End Function

    Private Sub cmbProcessType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbProcessType.SelectedIndexChanged
        Try
            'dtpDateFrom.Value = GetServerDate()
            'dtpDateTo.Value = GetServerDate()
            txtCustomerCode.Text = String.Empty
            lblCustCodeDes.Text = String.Empty
            txtInvoiceFrom.Text = String.Empty
            txtInvoiceTo.Text = String.Empty
            fspEWayBillDetail.MaxRows = 0
            InitializeSpread()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub UpdateSuppInvoices()
        Dim dtInvoices As New DataTable
        Dim validFromDate As DateTime
        Dim validFromTime As DateTime
        Dim validToDate As DateTime
        Dim validToTime As DateTime
        Dim validFrom As DateTime
        Dim validTo As DateTime
        Dim eWayType As String = String.Empty
        Dim chkSelect As String = String.Empty
        Dim irnNo As String = String.Empty
        Dim irnBarcode As String = String.Empty
        Try
            If ValidateSuppInvoices() Then
                Exit Sub
            End If

            With fspEWayBillDetail
                For i As Integer = 1 To .MaxRows
                    .Row = i
                    .Col = enumBillDetail.TICK
                    chkSelect = Convert.ToString(.Value)
                    If chkSelect = "1" Then
                        .Row = i
                        .Col = enumBillDetail.EWAY_IRN_NO
                        irnNo = Convert.ToString(.Text).Trim

                        .Row = i
                        .Col = enumBillDetail.IRN_BARCODE_STRING
                        irnBarcode = Convert.ToString(.Text).Trim

                        If irnNo.Length > 0 Then
                            If irnBarcode.Length = 0 Then
                                If MessageBox.Show("Are you sure to update IRN No. without scan/enter IRN barcode?", "eMPro", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = Windows.Forms.DialogResult.No Then
                                    .Row = i
                                    .Col = enumBillDetail.IRN_BARCODE_STRING
                                    .Text = String.Empty
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Exit Sub
                                Else
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next
            End With

            dtInvoices.Columns.Add("UNIT_CODE", GetType(System.String))
            dtInvoices.Columns.Add("DOC_NO", GetType(System.Int64))
            dtInvoices.Columns.Add("VOUCHER_NO", GetType(System.String))
            dtInvoices.Columns.Add("DR_CR", GetType(System.String))
            dtInvoices.Columns.Add("EWAY_IRN_NO", GetType(System.String))
            dtInvoices.Columns.Add("EWAY_IRN_DATE", GetType(System.DateTime))
            dtInvoices.Columns.Add("IRN_CANCEL", GetType(System.Boolean))
            dtInvoices.Columns.Add("IRN_CANCEL_DATE", GetType(System.DateTime))
            dtInvoices.Columns.Add("IRN_BARCODE_DATA", GetType(System.String))
            Dim drInvoices As DataRow
            Dim isChecked As String = String.Empty
            With fspEWayBillDetail
                For i As Integer = 1 To .MaxRows
                    .Row = i
                    .Col = enumBillDetail.TICK
                    isChecked = Convert.ToString(.Value)
                    If isChecked = "1" Then
                        drInvoices = dtInvoices.NewRow()
                        drInvoices("UNIT_CODE") = gstrUnitId
                        .Row = i
                        .Col = enumBillDetail.DOC_NO
                        drInvoices("DOC_NO") = .Text
                        .Row = i
                        .Col = enumBillDetail.VOUCHER_NO
                        drInvoices("VOUCHER_NO") = .Text
                        .Row = i
                        .Col = enumBillDetail.DR_CR
                        drInvoices("DR_CR") = .Text
                        .Row = i
                        .Col = enumBillDetail.EWAY_IRN_NO
                        If Convert.ToString(.Text).Trim <> "" Then
                            drInvoices("EWAY_IRN_NO") = Convert.ToString(.Text).Trim
                            .Row = i
                            .Col = enumBillDetail.EWAY_IRN_DATE
                            drInvoices("EWAY_IRN_DATE") = Convert.ToDateTime(.Text)

                            .Row = i
                            .Col = enumBillDetail.IRN_BARCODE_STRING
                            drInvoices("IRN_BARCODE_DATA") = Convert.ToString(.Text).Trim

                            .Row = i
                            .Col = enumBillDetail.EWAY_IRN_CANCEL
                            drInvoices("IRN_CANCEL") = IIf(.Text.ToUpper().Trim = "YES", True, False)

                            If .Text.ToUpper().Trim = "YES" Then
                                .Row = i
                                .Col = enumBillDetail.EWAY_IRN_CANCEL_DATE
                                drInvoices("IRN_CANCEL_DATE") = Convert.ToDateTime(.Text)
                            Else
                                drInvoices("IRN_CANCEL_DATE") = DBNull.Value
                            End If

                        Else
                            drInvoices("EWAY_IRN_NO") = DBNull.Value
                            drInvoices("EWAY_IRN_DATE") = DBNull.Value
                            drInvoices("IRN_BARCODE_DATA") = DBNull.Value
                            drInvoices("IRN_CANCEL") = DBNull.Value
                            drInvoices("IRN_CANCEL_DATE") = DBNull.Value
                        End If

                        dtInvoices.Rows.Add(drInvoices)
                    End If
                Next
            End With
            If dtInvoices IsNot Nothing AndAlso dtInvoices.Rows.Count > 0 Then
                Dim sqlCmd As New SqlCommand
                With sqlCmd
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 300 ' 5 Minute
                    .CommandText = "USP_UPDATE_EWAY_DETAIL_IN_INVOICE"
                    .Parameters.Clear()
                    .Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
                    .Parameters.AddWithValue("@DATE_FROM", dtpDateFrom.Value)
                    .Parameters.AddWithValue("@DATE_To", dtpDateTo.Value)
                    .Parameters.AddWithValue("@OPERATION_FLAG", "UPDATE_SUPP_INVOICE")
                    .Parameters.AddWithValue("@TYPE_SUPP_EWAY_BILL_UPDATE", dtInvoices)
                    .Parameters.AddWithValue("@USER_ID", mP_User)
                    .Parameters.Add("@MSG", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                    SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                    If Convert.ToString(.Parameters("@MSG").Value) <> "" Then
                        MsgBox(Convert.ToString(.Parameters("@MSG").Value), MsgBoxStyle.Exclamation, ResolveResString(100))
                    Else
                        MsgBox("Selected Invoices updated successfully.")
                        btnClear.PerformClick()
                    End If
                End With
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            dtInvoices.Dispose()
        End Try
    End Sub

    Private Function ValidateSuppInvoices() As Boolean
        Dim result As Boolean = False
        Dim chkSelect As String = String.Empty
        Dim chkSelectInner As String = String.Empty
        Dim countCheck As Integer = 0
        Dim eWayBill_IRN_No As String = String.Empty
        Dim IRN_DATE As DateTime
        Dim IRN_CANCEL_DATE As DateTime
        Dim IRN_CANCEL_FLAG As String = String.Empty
        Dim IRN_BARCODE_STRING As String = String.Empty
        Dim voNO As String = String.Empty
        Dim eWayBill_IRN_No_Inner As String = String.Empty
        Try
            If fspEWayBillDetail Is Nothing OrElse fspEWayBillDetail.MaxRows = 0 Then
                MessageBox.Show("No Invoice for updating.Please first search invoices.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                result = True
                Return result
            End If
            With fspEWayBillDetail
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
                    MessageBox.Show("Please select atleast one invoice for update.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    result = True
                    Return result
                End If

                For i As Integer = 1 To .MaxRows
                    .Row = i
                    .Col = enumBillDetail.TICK
                    chkSelect = Convert.ToString(.Value)
                    If chkSelect = "1" Then
                        .Row = i
                        .Col = enumBillDetail.VOUCHER_NO
                        voNO = Convert.ToString(.Text).Trim

                        .Row = i
                        .Col = enumBillDetail.IRN_BARCODE_STRING
                        IRN_BARCODE_STRING = Convert.ToString(.Text).Trim

                        .Row = i
                        .Col = enumBillDetail.EWAY_IRN_NO
                        eWayBill_IRN_No = Convert.ToString(.Text)

                        If eWayBill_IRN_No.Trim = "" Then
                            MessageBox.Show("IRN No. Required to Update!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            result = True
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            Return result
                        End If

                        If eWayBill_IRN_No.Trim <> "" Then
                            If eWayBill_IRN_No.Trim.Length <> 64 Then
                                MessageBox.Show("IRN No. must be of length 64.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                .Row = i
                                .Col = enumBillDetail.EWAY_IRN_NO
                                result = True
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                                Return result
                            End If
                            Dim strSql As String = "SELECT TOP 1 ISNULL(VO_NO,'') FROM SUPPLEMENTARY_IRN (NOLOCK) WHERE UNIT_CODE='" & gstrUnitId & "' AND RTRIM(LTRIM(IRN_NO))='" & eWayBill_IRN_No.Trim & "' AND VO_NO<>'" & voNO & "'"
                            Dim res As String = Convert.ToString(SqlConnectionclass.ExecuteScalar(strSql))
                            If res.Length > 0 Then
                                MessageBox.Show("Enter IRN No. associated with voucher no. " & res & vbCrLf & " Please enter unique IRN No.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                .Row = i
                                .Col = enumBillDetail.EWAY_IRN_NO
                                result = True
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                                Return result
                            End If
                        End If

                        .Row = i
                        .Col = enumBillDetail.EWAY_IRN_DATE
                        If String.IsNullOrEmpty(.Text) And eWayBill_IRN_No.Trim <> "" Then
                            MessageBox.Show("Please enter IRN DATE.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            result = True
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            Return result
                        End If
                        If Not IsDate(.Text) And eWayBill_IRN_No.Trim <> "" Then
                            MessageBox.Show("Please enter valid [IRN Date].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            result = True
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            Return result
                        End If
                        If Not String.IsNullOrEmpty(.Text) Then
                            IRN_DATE = Convert.ToDateTime(.Text)
                        End If

                        If IRN_BARCODE_STRING.Trim <> "" And eWayBill_IRN_No.Trim <> "" Then
                            If IRN_BARCODE_STRING.Trim.Length < 500 Then
                                MessageBox.Show("Please scan/enter valid IRN barcode.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                .Row = i
                                .Col = enumBillDetail.IRN_BARCODE_STRING
                                .Text = String.Empty
                                result = True
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                                Return result
                            End If
                        End If

                        .Row = i
                        .Col = enumBillDetail.EWAY_IRN_CANCEL
                        IRN_CANCEL_FLAG = Convert.ToString(.Text)

                        .Row = i
                        .Col = enumBillDetail.EWAY_IRN_CANCEL_DATE
                        If Not String.IsNullOrEmpty(.Text) And eWayBill_IRN_No.Trim <> "" Then
                            If IRN_CANCEL_FLAG = "NO" And Convert.ToString(.Text) <> "" Then
                                MessageBox.Show("IRN Cancel date not allowed for not Cancel IRN!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                .Text = String.Empty
                                result = True
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                                Return result

                            Else
                                If String.IsNullOrEmpty(.Text) Then
                                    MessageBox.Show("Please enter IRN CANCEL DATE.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    result = True
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Return result
                                End If
                                If Not IsDate(.Text) Then
                                    MessageBox.Show("Please enter valid [IRN CANCEL DATE].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    result = True
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Return result
                                End If

                                IRN_CANCEL_DATE = Convert.ToDateTime(.Text)
                            End If
                        ElseIf IRN_CANCEL_FLAG = "YES" And eWayBill_IRN_No.Trim() <> "" And .Text.ToString() = "" Then

                            MessageBox.Show("Please enter IRN CANCEL DATE.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            result = True
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            Return result


                        End If


                        If IRN_CANCEL_FLAG = "YES" And eWayBill_IRN_No.Trim() = "" Then
                            If Convert.ToString(.Text) = "" Then
                                MessageBox.Show("IRN Cancel date not blank for Canceled IRN!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                result = True
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                                Return result
                            Else
                                MessageBox.Show("No IRN Number found against Cancel!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                result = True
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                                Return result
                            End If
                        End If
                    End If
                Next
                For i As Integer = 1 To .MaxRows
                    .Row = i
                    .Col = enumBillDetail.TICK
                    chkSelect = Convert.ToString(.Value)
                    If chkSelect = "1" Then
                        .Row = i
                        .Col = enumBillDetail.EWAY_IRN_NO
                        eWayBill_IRN_No = Convert.ToString(.Text)
                        If String.IsNullOrEmpty(eWayBill_IRN_No) Then Continue For

                        For j As Integer = 1 To .MaxRows
                            If i = j Then Continue For

                            .Row = j
                            .Col = enumBillDetail.TICK
                            chkSelectInner = Convert.ToString(.Value)
                            If chkSelectInner = "1" Then
                                .Row = j
                                .Col = enumBillDetail.EWAY_IRN_NO
                                eWayBill_IRN_No_Inner = Convert.ToString(.Text)
                                If String.IsNullOrEmpty(eWayBill_IRN_No_Inner) Then Continue For

                                If eWayBill_IRN_No.Trim = eWayBill_IRN_No_Inner.Trim Then
                                    MessageBox.Show("IRN No. on row no. " & i.ToString & " and row no. " & j.ToString & " is duplicate.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                    .Row = i
                                    .Col = enumBillDetail.EWAY_IRN_NO
                                    result = True
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Return result
                                End If
                            End If
                        Next
                    End If
                Next
            End With
        Catch ex As Exception
            result = True
            RaiseException(ex)
        End Try
        Return result
    End Function

    Private Sub fspEWayBillDetail_KeyUpEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyUpEvent) Handles fspEWayBillDetail.KeyUpEvent
        Dim varDocNo As String
        Dim vardate As Date
        Dim varchk As Object
        Try
            With fspEWayBillDetail
                .Col = enumBillDetail.DOC_NO
                .Row = .ActiveRow
                varDocNo = .Text
                .Col = enumBillDetail.TICK
                varchk = .Value
                If varchk = 1 Then
                    If e.keyCode = Keys.F1 Then
                        With FRMMKTTRN0102A
                            .Customer_code = txtCustomerCode.Text
                            .Customer_name = lblCustCodeDes.Text
                            .Invoice_number = varDocNo
                            .ShowDialog()
                            If varDocNo = .InvoiceNo Then
                                vardate = .IrnDate
                                fspEWayBillDetail.Col = enumBillDetail.EWAY_IRN_NO
                                fspEWayBillDetail.Text = .Irnnumber
                                fspEWayBillDetail.SetText(enumBillDetail.EWAY_IRN_DATE, fspEWayBillDetail.Row, vardate)
                                fspEWayBillDetail.Col = enumBillDetail.IRN_BARCODE_STRING
                                fspEWayBillDetail.Text = .IrnBarcodeString
                            Else
                                MsgBox("Invoice number is not correct.", MsgBoxStyle.Exclamation, ResolveResString(100))
                            End If


                        End With
                    End If
                End If

            End With
        Catch ex As Exception

        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class