Imports System.Data.SqlClient
Imports System.Text.RegularExpressions

Public Class FRMMKTTRN0084A
    Public gItem_Code As String = String.Empty
    Public gProvDocNo As String = String.Empty
    Public gInvoiceRate As String = String.Empty
    Public gProvDocNo_Type As String = String.Empty
    Public gNewRate As String = String.Empty
    Public gfilemode As Boolean = False

#Region "Form Events"

    Private Sub FRMMKTTRN0084A_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            lblItemValue.Text = gItem_Code
            lblProvDocNoVal.Text = gProvDocNo
            AddColumnInvoiceWiseDtlGrid()
            getInvoiceWiseDetail()
            Me.BringToFront()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

#End Region

#Region "Functions & Procedures"

    Private Sub getInvoiceWiseDetail()
        Dim strQry As String = String.Empty
        Dim odt As New DataTable

        Try
            If gfilemode = False Then
                If Not gProvDocNo_Type.Equals("AUTH") Then
                    strQry = " SELECT ITEM_CODE Part_Code, VARMODEL, INVOICE_NO, INVOICEDATE,CUSTREFNO, CUSTREFDATE, AMENDNO, case when AMENDDATE='1 Jan 1900' then NULL else AMENDDATE end AMENDDATE, cast(isnull(INVOICE_QTY,0.0) as numeric(18,2)) INVOICE_QTY" & _
                         " , CAST(ROUND(RATE,2,1) AS VARCHAR(20)) Rate, cast(isnull(REJECTIONQTY,0.0) as numeric(18,2)) REJECTIONQTY, cast(isnull(CDRQTY,0.0) as numeric(18,2)) CDRQTY, cast(isnull(ACTUALINVOICEQTY,0.0) as numeric(18,2)) ACTUALINVOICEQTY" & _
                         " FROM SALES_PROV_TMPPARTDETAIL(NOLOCK) TMP WHERE TMP.ITEM_CODE='" + gItem_Code.Trim() + "' AND IPADDRESS='" + gstrIpaddressWinSck + "' AND UNIT_CODE='" + gstrUNITID + "' and convert(varchar(20),Rate)='" + gInvoiceRate + "'"
                Else
                    strQry = " SELECT ITEM_CODE PART_CODE, VARMODEL, INVOICE_NO, INVOICEDATE,CUSTREFNO, CUSTREFDATE, AMENDNO, CASE WHEN AMENDDATE='1 JAN 1900' THEN NULL ELSE AMENDDATE END AMENDDATE, CAST(ISNULL(INVOICE_QTY,0.0) AS NUMERIC(18,2)) INVOICE_QTY" & _
                            " , CAST(ROUND(RATE,2,1) AS VARCHAR(20)) RATE, CAST(ISNULL(REJECTIONQTY,0.0) AS NUMERIC(18,2)) REJECTIONQTY, CAST(ISNULL(CDRQTY,0.0) AS NUMERIC(18,2)) CDRQTY, CAST(ISNULL(ACTUALINVOICEQTY,0.0) AS NUMERIC(18,2)) ACTUALINVOICEQTY" & _
                            " FROM SALES_PROV_TMPPARTDETAIL_AUTH(NOLOCK) TMP WHERE TMP.ITEM_CODE='" + gItem_Code.Trim() + "' AND IPADDRESS='" + gstrIpaddressWinSck + "' AND UNIT_CODE='" + gstrUNITID + "' and convert(varchar(20),Rate)='" + gInvoiceRate + "'"
                End If
            End If
            If lblProvDocNoVal.Text.Trim = "" Then
                If gfilemode = True Then
                    strQry = "SELECT  ITEM_CODE PART_CODE, INVOICENO INVOICE_NO,STRBILLDATE INVOICEDATE,'' CUSTREFNO, '' CUSTREFDATE, '' AMENDNO, 'N/A' AMENDDATE,STRSHP INVOICE_QTY,STROLDRATE Rate,0.0 REJECTIONQTY,0.0 CDRQTY,STRACP ACTUALINVOICEQTY " & _
                    " FROM TMP_SALESFILEDATA WHERE ITEM_CODE='" + gItem_Code.Trim() + "' AND IPADDRESS='" + gstrIpaddressWinSck + "' AND UNIT_CODE='" + gstrUNITID + "' and convert(varchar(20),STROLDRATE)='" + gInvoiceRate + "'" & _
                    " and convert(varchar(20),STRNEWRATE)='" + gNewRate + "' ORDER BY STRBILLDATE"
                End If
            Else
                If gfilemode = True Then
                    strQry = "SELECT  ITEM_CODE PART_CODE, INVOICENO INVOICE_NO,STRBILLDATE INVOICEDATE,'' CUSTREFNO, '' CUSTREFDATE,'' AMENDNO, 'N/A' AMENDDATE,STRSHP INVOICE_QTY,STROLDRATE Rate,0.0 REJECTIONQTY,0.0 CDRQTY,STRACP ACTUALINVOICEQTY " & _
                    " FROM SALESFILEDATA WHERE  UNIT_CODE='" + gstrUNITID + "' and PROV_DOCNO ='" + lblProvDocNoVal.Text.Trim + "'" & _
                    " ORDER BY STRBILLDATE"
                End If
            End If

            odt = SqlConnectionclass.GetDataTable(strQry)
            If Not IsNothing(odt) Then
                dgvInvWiseDtl.DataSource = odt
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    '''<summary>
    ''' Column addition in Invoice Wise Detail Grid
    ''' </summary>
    Private Sub AddColumnInvoiceWiseDtlGrid()
        Try
            If dgvInvWiseDtl.Columns.Count = 0 Then
                dgvInvWiseDtl.ColumnHeadersHeight = 35
                Dim dgvC As New DataGridViewTextBoxColumn
                dgvC.DataPropertyName = "Part_Code"
                dgvC.Name = "Part_Code"
                dgvC.HeaderText = "Part Code"
                dgvC.Width = 115
                dgvC.ReadOnly = True
                dgvInvWiseDtl.Columns.Add(dgvC)

                Dim dgvD As New DataGridViewTextBoxColumn
                dgvD.DataPropertyName = "VarModel"
                dgvD.Name = "VarModel"
                dgvD.HeaderText = "Model Code"
                dgvD.Width = 70
                dgvD.ReadOnly = True
                dgvInvWiseDtl.Columns.Add(dgvD)

                Dim dgvIN As New DataGridViewTextBoxColumn
                dgvIN.DataPropertyName = "Invoice_No"
                dgvIN.Name = "Invoice_No"
                dgvIN.HeaderText = "Invoice No"
                dgvIN.Width = 65
                dgvIN.ReadOnly = True
                dgvInvWiseDtl.Columns.Add(dgvIN)

                Dim dgvdt As New DataGridViewTextBoxColumn
                dgvdt.DataPropertyName = "InvoiceDate"
                dgvdt.Name = "InvoiceDate"
                dgvdt.HeaderText = "Invoice Date"
                dgvdt.Width = 80
                dgvdt.ReadOnly = True
                dgvInvWiseDtl.Columns.Add(dgvdt)

                Dim dgvCRF As New DataGridViewTextBoxColumn
                dgvCRF.DataPropertyName = "CustRefNo"
                dgvCRF.Name = "CustRefNo"
                dgvCRF.HeaderText = "Customer Ref. No"
                dgvCRF.Width = 115
                dgvCRF.ReadOnly = True
                dgvInvWiseDtl.Columns.Add(dgvCRF)

                Dim dgvRC As New DataGridViewTextBoxColumn
                dgvRC.DataPropertyName = "CustRefDate"
                dgvRC.Name = "CustRefDate"
                dgvRC.HeaderText = "Ref. Date"
                dgvRC.Width = 70
                dgvRC.ReadOnly = True
                dgvInvWiseDtl.Columns.Add(dgvRC)

                Dim dgvANO As New DataGridViewTextBoxColumn
                dgvANO.DataPropertyName = "AmendNo"
                dgvANO.Name = "AmendNo"
                dgvANO.HeaderText = "Amend No."
                dgvANO.Width = 70
                dgvANO.ReadOnly = True
                dgvInvWiseDtl.Columns.Add(dgvANO)

                Dim dgvADT As New DataGridViewTextBoxColumn
                dgvADT.DataPropertyName = "AmendDate"
                dgvADT.Name = "AmendDate"
                dgvADT.HeaderText = "Amend Date"
                dgvADT.Width = 80
                dgvADT.ReadOnly = True
                dgvInvWiseDtl.Columns.Add(dgvADT)

                Dim dgvIR As New DataGridViewTextBoxColumn
                dgvIR.DataPropertyName = "Rate"
                dgvIR.Name = "Rate"
                dgvIR.HeaderText = "Invoice Rate"
                dgvIR.Width = 60
                dgvIR.ReadOnly = True
                dgvIR.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgvInvWiseDtl.Columns.Add(dgvIR)

                Dim dgvQ As New DataGridViewTextBoxColumn
                dgvQ.DataPropertyName = "Invoice_Qty"
                dgvQ.Name = "Invoice_Qty"
                dgvQ.HeaderText = "Invoice Qty"
                dgvQ.Width = 68
                dgvQ.ReadOnly = True
                dgvQ.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgvInvWiseDtl.Columns.Add(dgvQ)

                Dim dgvRQ As New DataGridViewTextBoxColumn
                dgvRQ.DataPropertyName = "RejectionQty"
                dgvRQ.Name = "RejectionQty"
                dgvRQ.HeaderText = "Rejection Qty"
                dgvRQ.Width = 80
                dgvRQ.ReadOnly = True
                dgvRQ.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgvInvWiseDtl.Columns.Add(dgvRQ)

                Dim dgvCDR As New DataGridViewTextBoxColumn
                dgvCDR.DataPropertyName = "CDRQty"
                dgvCDR.Name = "CDRQty"
                dgvCDR.HeaderText = "CDR Qty"
                dgvCDR.Width = 60
                dgvCDR.ReadOnly = True
                dgvCDR.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgvInvWiseDtl.Columns.Add(dgvCDR)

                Dim dgvN As New DataGridViewTextBoxColumn
                dgvN.DataPropertyName = "ActualInvoiceQty"
                dgvN.Name = "ActualInvoiceQty"
                dgvN.HeaderText = "Actual Invoice Qty"
                dgvN.Width = 98
                dgvN.ReadOnly = True
                dgvN.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgvInvWiseDtl.Columns.Add(dgvN)

                dgvInvWiseDtl.AutoGenerateColumns = False
                dgvInvWiseDtl.RowsDefaultCellStyle.BackColor = Color.Lavender
                dgvInvWiseDtl.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(190, 200, 255)
                dgvInvWiseDtl.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

#End Region

#Region "Control events"

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Try
            Me.Dispose()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

#End Region

End Class