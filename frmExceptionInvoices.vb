Imports System.Data
Imports System.Data.SqlClient
'(c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
'File Name    -  frmExceptionInvoices.frm
'CREATED BY   -  ASHISH SHARMA
'CREATED ON   -  25 AUG 2020
'PURPOSE      -  102027599 - IRN CHANGES (Display Exception Invoice (s) list)
Public Class frmExceptionInvoices
    Dim invoiceType As String = String.Empty
    Dim DispLotNo As String = String.Empty


    Public WriteOnly Property SetInvoiceType() As String
        Set(ByVal value As String)
            invoiceType = value
        End Set
    End Property
    Public WriteOnly Property SetDispLotNo() As String
        Set(ByVal value As String)
            DispLotNo = value
        End Set
    End Property

    Private Enum GridColumn
        Doc_No = 0
        Cust_Code
        Cust_Name
        Cust_Type
        IRN_Required
        Eway_Required
        IRN_Received
        Eway_Received
    End Enum
    Private Sub frmExceptionInvoices_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            SetBackGroundColorNew(Me, True)
            ConfigureGridColumn()
            FillGrid()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub ConfigureGridColumn()
        Try
            dgvExceptionInvoices.Columns.Clear()

            dgvExceptionInvoices.Columns.Add("DocNo", "Doc No.")
            dgvExceptionInvoices.Columns.Add("CustCode", "Customer Code")
            dgvExceptionInvoices.Columns.Add("CustName", "Customer Name")
            dgvExceptionInvoices.Columns.Add("CustType", "Customer Type")
            dgvExceptionInvoices.Columns.Add("IRNRequired", "IRN Required")
            dgvExceptionInvoices.Columns.Add("EWAYRequired", "eWAY Required")
            dgvExceptionInvoices.Columns.Add("IRNReceived", "IRN Received")
            dgvExceptionInvoices.Columns.Add("EWAYReceived", "eWAY Received")


            dgvExceptionInvoices.Columns(GridColumn.Doc_No).Width = 100
            dgvExceptionInvoices.Columns(GridColumn.Cust_Code).Width = 75
            dgvExceptionInvoices.Columns(GridColumn.Cust_Name).Width = 200
            dgvExceptionInvoices.Columns(GridColumn.Cust_Type).Width = 75
            dgvExceptionInvoices.Columns(GridColumn.IRN_Required).Width = 65
            dgvExceptionInvoices.Columns(GridColumn.Eway_Required).Width = 70
            dgvExceptionInvoices.Columns(GridColumn.IRN_Received).Width = 65
            dgvExceptionInvoices.Columns(GridColumn.Eway_Received).Width = 70


            dgvExceptionInvoices.Columns(GridColumn.Doc_No).HeaderCell.Style.Font = New Font(dgvExceptionInvoices.Font, FontStyle.Bold)
            dgvExceptionInvoices.Columns(GridColumn.Cust_Code).HeaderCell.Style.Font = New Font(dgvExceptionInvoices.Font, FontStyle.Bold)
            dgvExceptionInvoices.Columns(GridColumn.Cust_Name).HeaderCell.Style.Font = New Font(dgvExceptionInvoices.Font, FontStyle.Bold)
            dgvExceptionInvoices.Columns(GridColumn.Cust_Type).HeaderCell.Style.Font = New Font(dgvExceptionInvoices.Font, FontStyle.Bold)
            dgvExceptionInvoices.Columns(GridColumn.IRN_Required).HeaderCell.Style.Font = New Font(dgvExceptionInvoices.Font, FontStyle.Bold)
            dgvExceptionInvoices.Columns(GridColumn.Eway_Required).HeaderCell.Style.Font = New Font(dgvExceptionInvoices.Font, FontStyle.Bold)
            dgvExceptionInvoices.Columns(GridColumn.IRN_Received).HeaderCell.Style.Font = New Font(dgvExceptionInvoices.Font, FontStyle.Bold)
            dgvExceptionInvoices.Columns(GridColumn.Eway_Received).HeaderCell.Style.Font = New Font(dgvExceptionInvoices.Font, FontStyle.Bold)

            dgvExceptionInvoices.Columns(GridColumn.Doc_No).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvExceptionInvoices.Columns(GridColumn.Cust_Code).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvExceptionInvoices.Columns(GridColumn.Cust_Name).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvExceptionInvoices.Columns(GridColumn.Cust_Type).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvExceptionInvoices.Columns(GridColumn.IRN_Required).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvExceptionInvoices.Columns(GridColumn.Eway_Required).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvExceptionInvoices.Columns(GridColumn.IRN_Received).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvExceptionInvoices.Columns(GridColumn.Eway_Received).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            dgvExceptionInvoices.Columns(GridColumn.Doc_No).ReadOnly = True
            dgvExceptionInvoices.Columns(GridColumn.Cust_Code).ReadOnly = True
            dgvExceptionInvoices.Columns(GridColumn.Cust_Name).ReadOnly = True
            dgvExceptionInvoices.Columns(GridColumn.Cust_Type).ReadOnly = True
            dgvExceptionInvoices.Columns(GridColumn.IRN_Required).ReadOnly = True
            dgvExceptionInvoices.Columns(GridColumn.Eway_Required).ReadOnly = True
            dgvExceptionInvoices.Columns(GridColumn.IRN_Received).ReadOnly = True
            dgvExceptionInvoices.Columns(GridColumn.Eway_Received).ReadOnly = True
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FillGrid()
        Dim dtExceptionInvoices As New DataTable
        Dim cmd As SqlCommand = New SqlCommand()
        Try
            cmd.CommandText = "USP_GET_LOCKED_EXCEPTION_INVOICES"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
            If Not String.IsNullOrEmpty(invoiceType) Then
                cmd.Parameters.AddWithValue("@INVOICE_TYPE", invoiceType)
            End If
            If Not String.IsNullOrEmpty(DispLotNo) Then
                cmd.Parameters.AddWithValue("@DISP_LOT_NO", DispLotNo)
            End If
            cmd.Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
            dtExceptionInvoices.Load(SqlConnectionclass.ExecuteReader(cmd))
            If Convert.ToString(cmd.Parameters("@MESSAGE").Value).Length > 0 Then
                MsgBox(Convert.ToString(cmd.Parameters("@MESSAGE").Value), MsgBoxStyle.Critical, "eMPro")
                Exit Sub
            End If
            dgvExceptionInvoices.Rows.Clear()
            If dtExceptionInvoices IsNot Nothing AndAlso dtExceptionInvoices.Rows.Count > 0 Then
                dgvExceptionInvoices.Rows.Add(dtExceptionInvoices.Rows.Count)
                Dim i As Integer = 0
                For Each dr As DataRow In dtExceptionInvoices.Rows
                    dgvExceptionInvoices.Rows(i).Cells(GridColumn.Doc_No).Value = dr("DOC_NO")
                    dgvExceptionInvoices.Rows(i).Cells(GridColumn.Cust_Code).Value = dr("ACCOUNT_CODE")
                    dgvExceptionInvoices.Rows(i).Cells(GridColumn.Cust_Name).Value = dr("CUST_NAME")
                    dgvExceptionInvoices.Rows(i).Cells(GridColumn.Cust_Type).Value = dr("CLASSIFICATIONTYPE")
                    dgvExceptionInvoices.Rows(i).Cells(GridColumn.IRN_Required).Value = dr("IRN_REQUIRED")
                    dgvExceptionInvoices.Rows(i).Cells(GridColumn.Eway_Required).Value = dr("EWAY_REQUIRED")
                    dgvExceptionInvoices.Rows(i).Cells(GridColumn.IRN_Received).Value = dr("IRN_RECEIVED")
                    dgvExceptionInvoices.Rows(i).Cells(GridColumn.Eway_Received).Value = dr("EWAY_RECEIVED")
                    i += 1
                Next
            Else
                MsgBox("No Record(s) found.", MsgBoxStyle.Critical, "eMPro")
                Exit Sub
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            If dtExceptionInvoices IsNot Nothing Then
                dtExceptionInvoices.Dispose()
            End If
            If cmd IsNot Nothing Then
                cmd.Dispose()
            End If
        End Try
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        Try
            FillGrid()
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
End Class