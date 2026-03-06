Imports System.Data
Imports System.Data.SqlClient
Public Class frmPaletteHelp
    Private _customerCode As String = String.Empty
    Dim _formMode As UCActXCtl.clsDeclares.ModeEnum
    Private _invoiceNo As Long = 0
    Private _itemCodes As DataTable
    Private _selectItemQty As DataTable
    Private _setItemQty As DataTable
    Public ReadOnly Property GetPalette() As DataTable
        Get
            Return _selectItemQty
        End Get
    End Property
    Public WriteOnly Property SetPalette() As DataTable
        Set(ByVal value As DataTable)
            _setItemQty = value
        End Set
    End Property
    Private Enum GridColumn
        Selection = 0
        PalleteLabel
        ItemCode
        ItemDescription
        Qty
    End Enum
    Private Enum GridSummaryColumn
        ItemCode
        ItemDescription
        Qty
    End Enum

    Public Sub New(ByVal formMode As UCActXCtl.clsDeclares.ModeEnum, ByVal customerCode As String, ByVal invoiceNo As Long, ByVal itemCodes As DataTable)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        _formMode = formMode
        _customerCode = customerCode
        _itemCodes = itemCodes
        _invoiceNo = invoiceNo
    End Sub
    Private Sub frmPaletteHelp_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetBackGroundColorNew(Me, True)
        If _formMode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            btnOK.Enabled = False
        Else
            btnOK.Enabled = True
        End If
        ConfigureGridColumn()
        FillGrid()
        FillSummaryGrid()
        SetPaletteItem()
    End Sub
    Private Sub ConfigureGridColumn()
        Try
            dgvPalette.Columns.Clear()

            Dim objChkBox As New DataGridViewCheckBoxColumn
            objChkBox.Name = "Selection"
            objChkBox.HeaderText = " "

            dgvPalette.Columns.Add(objChkBox)
            dgvPalette.Columns.Add("PalleteLabel", "Palette Label")
            dgvPalette.Columns.Add("ItemCode", "Item Code")
            dgvPalette.Columns.Add("ItemDescription", "Item Description")
            dgvPalette.Columns.Add("Qty", "Qty.")

            dgvPalette.Columns(GridColumn.Selection).Width = 35
            dgvPalette.Columns(GridColumn.PalleteLabel).Width = 110
            dgvPalette.Columns(GridColumn.ItemCode).Width = 120
            dgvPalette.Columns(GridColumn.ItemDescription).Width = 400
            dgvPalette.Columns(GridColumn.Qty).Width = 50


            dgvPalette.Columns(GridColumn.Selection).HeaderCell.Style.Font = New Font(dgvPalette.Font, FontStyle.Bold)
            dgvPalette.Columns(GridColumn.PalleteLabel).HeaderCell.Style.Font = New Font(dgvPalette.Font, FontStyle.Bold)
            dgvPalette.Columns(GridColumn.ItemCode).HeaderCell.Style.Font = New Font(dgvPalette.Font, FontStyle.Bold)
            dgvPalette.Columns(GridColumn.ItemDescription).HeaderCell.Style.Font = New Font(dgvPalette.Font, FontStyle.Bold)
            dgvPalette.Columns(GridColumn.Qty).HeaderCell.Style.Font = New Font(dgvPalette.Font, FontStyle.Bold)

            dgvPalette.Columns(GridColumn.Qty).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            dgvPalette.Columns(GridColumn.Selection).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvPalette.Columns(GridColumn.PalleteLabel).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvPalette.Columns(GridColumn.ItemCode).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvPalette.Columns(GridColumn.ItemDescription).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvPalette.Columns(GridColumn.Qty).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            If _formMode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                dgvPalette.Columns(GridColumn.Selection).ReadOnly = True
            Else
                dgvPalette.Columns(GridColumn.Selection).ReadOnly = False
            End If
            dgvPalette.Columns(GridColumn.PalleteLabel).ReadOnly = True
            dgvPalette.Columns(GridColumn.ItemCode).ReadOnly = True
            dgvPalette.Columns(GridColumn.ItemDescription).ReadOnly = True
            dgvPalette.Columns(GridColumn.Qty).ReadOnly = True

            dgvPalette.Columns(GridColumn.Selection).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvPalette.Columns(GridColumn.PalleteLabel).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvPalette.Columns(GridColumn.ItemCode).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvPalette.Columns(GridColumn.ItemDescription).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvPalette.Columns(GridColumn.Qty).SortMode = DataGridViewColumnSortMode.NotSortable

            dgvSummary.Columns.Clear()
            dgvSummary.Columns.Add("ItemCode", "Item Code")
            dgvSummary.Columns.Add("ItemDescription", "Item Description")
            dgvSummary.Columns.Add("Qty", "Total Qty.")

            dgvSummary.Columns(GridSummaryColumn.ItemCode).Width = 145
            dgvSummary.Columns(GridSummaryColumn.ItemDescription).Width = 450
            dgvSummary.Columns(GridSummaryColumn.Qty).Width = 100

            dgvSummary.Columns(GridSummaryColumn.ItemCode).HeaderCell.Style.Font = New Font(dgvSummary.Font, FontStyle.Bold)
            dgvSummary.Columns(GridSummaryColumn.ItemDescription).HeaderCell.Style.Font = New Font(dgvSummary.Font, FontStyle.Bold)
            dgvSummary.Columns(GridSummaryColumn.Qty).HeaderCell.Style.Font = New Font(dgvSummary.Font, FontStyle.Bold)

            dgvSummary.Columns(GridSummaryColumn.Qty).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            dgvSummary.Columns(GridSummaryColumn.ItemCode).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvSummary.Columns(GridSummaryColumn.ItemDescription).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvSummary.Columns(GridSummaryColumn.Qty).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            dgvSummary.Columns(GridSummaryColumn.ItemCode).ReadOnly = True
            dgvSummary.Columns(GridSummaryColumn.ItemDescription).ReadOnly = True
            dgvSummary.Columns(GridSummaryColumn.Qty).ReadOnly = True

            dgvSummary.Columns(GridSummaryColumn.ItemCode).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvSummary.Columns(GridSummaryColumn.ItemDescription).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvSummary.Columns(GridSummaryColumn.Qty).SortMode = DataGridViewColumnSortMode.NotSortable
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearch.TextChanged
        Try
            If dgvPalette Is Nothing OrElse dgvPalette.Rows.Count = 0 Then Exit Sub
            For i As Integer = 0 To dgvPalette.Rows.Count - 1
                dgvPalette.Rows(i).DefaultCellStyle.ForeColor = Color.Black
                dgvPalette.Rows(i).Cells(GridColumn.PalleteLabel).Style.Font = New Font(dgvPalette.Font, FontStyle.Regular)
                dgvPalette.Rows(i).Cells(GridColumn.ItemCode).Style.Font = New Font(dgvPalette.Font, FontStyle.Regular)
                dgvPalette.Rows(i).Cells(GridColumn.ItemDescription).Style.Font = New Font(dgvPalette.Font, FontStyle.Regular)
            Next
            If String.IsNullOrEmpty(txtSearch.Text) Then Exit Sub
            For i As Integer = 0 To dgvPalette.Rows.Count - 1
                If rdbItemCode.Checked Then
                    If Trim(UCase(Mid(dgvPalette.Rows(i).Cells(GridColumn.ItemCode).Value.ToString, 1, Len(txtSearch.Text)))) = Trim(UCase(txtSearch.Text)) Then
                        dgvPalette.Rows(i).DefaultCellStyle.ForeColor = Color.DarkBlue
                        dgvPalette.Rows(i).Cells(GridColumn.ItemCode).Style.Font = New Font(dgvPalette.Font, FontStyle.Bold)
                        dgvPalette.FirstDisplayedScrollingRowIndex = i
                        Exit For
                    End If
                ElseIf rdbItemDescription.Checked Then
                    If Trim(UCase(Mid(dgvPalette.Rows(i).Cells(GridColumn.ItemDescription).Value.ToString, 1, Len(txtSearch.Text)))) = Trim(UCase(txtSearch.Text)) Then
                        dgvPalette.Rows(i).DefaultCellStyle.ForeColor = Color.DarkBlue
                        dgvPalette.Rows(i).Cells(GridColumn.ItemDescription).Style.Font = New Font(dgvPalette.Font, FontStyle.Bold)
                        dgvPalette.FirstDisplayedScrollingRowIndex = i
                        Exit For
                    End If
                ElseIf rdbPaletteLabel.Checked Then
                    If Trim(UCase(Mid(dgvPalette.Rows(i).Cells(GridColumn.PalleteLabel).Value.ToString, 1, Len(txtSearch.Text)))) = Trim(UCase(txtSearch.Text)) Then
                        dgvPalette.Rows(i).DefaultCellStyle.ForeColor = Color.DarkBlue
                        dgvPalette.Rows(i).Cells(GridColumn.PalleteLabel).Style.Font = New Font(dgvPalette.Font, FontStyle.Bold)
                        dgvPalette.FirstDisplayedScrollingRowIndex = i
                        Exit For
                    End If
                End If
            Next
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FillGrid()
        Dim dtPalette As New DataTable
        Dim mode As String = String.Empty
        Try
            If String.IsNullOrEmpty(_customerCode) Then
                MsgBox("Please Select Customer Code")
                Exit Sub
            End If
            If _itemCodes Is Nothing OrElse _itemCodes.Rows.Count = 0 Then
                MsgBox("Please Select Item")
                Exit Sub
            End If
            If _formMode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                If Val(_invoiceNo) = 0 Then
                    MsgBox("Please select Invoice No.")
                    Exit Sub
                End If
            End If
            If _formMode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                mode = "EDIT"
            ElseIf _formMode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                mode = "ADD"
            Else
                mode = "VIEW"
            End If
            Dim cmd As SqlCommand = New SqlCommand()
            cmd.CommandText = "USP_NORMAL_FG_INVOICE_BARCODE"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
            cmd.Parameters.AddWithValue("@CUSTOMER_CODE", _customerCode)
            cmd.Parameters.AddWithValue("@PR_ITEM_TYPE", _itemCodes)
            cmd.Parameters.AddWithValue("@OPERATION_CODE", "GET_PALETTE")
            cmd.Parameters.AddWithValue("@MODE", UCase(mode))
            If _formMode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Or _formMode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                cmd.Parameters.AddWithValue("@INVOICE_NO", _invoiceNo)
            End If
            cmd.Parameters.AddWithValue("@RESULT", "")
            cmd.Parameters("@RESULT").Direction = ParameterDirection.InputOutput
            cmd.Parameters("@RESULT").SqlDbType = SqlDbType.VarChar
            cmd.Parameters("@RESULT").Size = 500
            cmd.Parameters.AddWithValue("@MESSAGE", "")
            cmd.Parameters("@MESSAGE").Direction = ParameterDirection.InputOutput
            cmd.Parameters("@MESSAGE").SqlDbType = SqlDbType.VarChar
            cmd.Parameters("@MESSAGE").Size = 500
            dtPalette.Load(SqlConnectionclass.ExecuteReader(cmd))
            If Convert.ToString(cmd.Parameters("@MESSAGE").Value).Length > 0 Then
                MsgBox(Convert.ToString(cmd.Parameters("@MESSAGE").Value), MsgBoxStyle.Critical, "eMPro")
                Exit Sub
            End If
            If Convert.ToString(cmd.Parameters("@RESULT").Value) = "Y" Then
                If dtPalette IsNot Nothing AndAlso dtPalette.Rows.Count > 0 Then
                    dgvPalette.Rows.Clear()
                    dgvPalette.Rows.Add(dtPalette.Rows.Count)
                    Dim i As Integer = 0
                    For Each dr As DataRow In dtPalette.Rows
                        dgvPalette.Rows(i).Cells(GridColumn.Selection).Value = False
                        dgvPalette.Rows(i).Cells(GridColumn.PalleteLabel).Value = dr("PALETTE_LABEL")
                        dgvPalette.Rows(i).Cells(GridColumn.ItemCode).Value = dr("ITEM_CODE")
                        dgvPalette.Rows(i).Cells(GridColumn.ItemDescription).Value = dr("ITEM_NAME")
                        dgvPalette.Rows(i).Cells(GridColumn.Qty).Value = dr("QUANTITY")
                        i += 1
                    Next
                End If
            Else
                MsgBox("Normal Invoice FG Barcoding functionality is not active", MsgBoxStyle.Information, "eMPro")
                Exit Sub
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        _selectItemQty = _setItemQty
        Me.Close()
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Try
            If Not ValidateData() Then Exit Sub
            CreateReturnDataTable()
            Dim drSelectedItem As DataRow
            For i As Integer = 0 To dgvPalette.Rows.Count - 1
                If Convert.ToBoolean(dgvPalette.Rows(i).Cells(GridColumn.Selection).Value) Then
                    drSelectedItem = _selectItemQty.NewRow()
                    drSelectedItem("PALETTE_LABEL") = Convert.ToString(dgvPalette.Rows(i).Cells(GridColumn.PalleteLabel).Value)
                    drSelectedItem("ITEM_CODE") = Convert.ToString(dgvPalette.Rows(i).Cells(GridColumn.ItemCode).Value)
                    drSelectedItem("QTY") = Convert.ToInt32(dgvPalette.Rows(i).Cells(GridColumn.Qty).Value)
                    _selectItemQty.Rows.Add(drSelectedItem)
                End If
            Next
            Me.Close()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub CreateReturnDataTable()
        _selectItemQty = New DataTable()
        _selectItemQty.Columns.Add("PALETTE_LABEL", GetType(System.String))
        _selectItemQty.Columns.Add("ITEM_CODE", GetType(System.String))
        _selectItemQty.Columns.Add("QTY", GetType(System.Int32))
    End Sub
    Private Function ValidateData() As Boolean
        Dim result As Boolean = True
        If dgvPalette.Rows.Count = 0 Then
            MsgBox("Palette Data not found", MsgBoxStyle.Information, "eMPro")
            result = False
            Return result
        End If

        Dim countRow As Integer = 0
        For j As Integer = 0 To _itemCodes.Rows.Count - 1
            countRow = 0
            For i As Integer = 0 To dgvPalette.Rows.Count - 1
                If Convert.ToString(_itemCodes.Rows(j)(0)) <> Convert.ToString(dgvPalette.Rows(i).Cells(GridColumn.ItemCode).Value) Then
                    countRow += 1
                End If
            Next
            If countRow = dgvPalette.Rows.Count Then
                MsgBox("Palette not found for item code : " & Convert.ToString(_itemCodes.Rows(j)(0)), MsgBoxStyle.Information, "eMPro")
                dgvPalette.Focus()
                dgvPalette.CurrentCell = dgvPalette.Rows(0).Cells(GridColumn.Selection)
                result = False
                Return result
            End If
        Next
        
        Dim flag As Boolean = False
        For j As Integer = 0 To _itemCodes.Rows.Count - 1
            flag = False
            For i As Integer = 0 To dgvPalette.Rows.Count - 1
                If Convert.ToString(_itemCodes.Rows(j)(0)) = Convert.ToString(dgvPalette.Rows(i).Cells(GridColumn.ItemCode).Value) Then
                    If Convert.ToBoolean(dgvPalette.Rows(i).Cells(GridColumn.Selection).Value) Then
                        flag = True
                        Exit For
                    End If
                End If
            Next
            If Not flag Then
                MsgBox("Please select atleast one palette for item code : " & Convert.ToString(_itemCodes.Rows(j)(0)), MsgBoxStyle.Information, "eMPro")
                dgvPalette.Focus()
                dgvPalette.CurrentCell = dgvPalette.Rows(i).Cells(GridColumn.Selection)
                result = False
                Return result
            End If
        Next
       
        Dim itemCode As String = String.Empty
        Dim paletteCode As String = String.Empty
        Dim itemCode1 As String = String.Empty
        Dim paletteCode1 As String = String.Empty
        For i As Integer = 0 To dgvPalette.Rows.Count - 1
            If Convert.ToBoolean(dgvPalette.Rows(i).Cells(GridColumn.Selection).Value) Then
                paletteCode = Convert.ToString(dgvPalette.Rows(i).Cells(GridColumn.PalleteLabel).Value)
                itemCode = Convert.ToString(dgvPalette.Rows(i).Cells(GridColumn.ItemCode).Value)
                If i > 0 Then
                    For j As Integer = 0 To i - 1
                        paletteCode1 = Convert.ToString(dgvPalette.Rows(j).Cells(GridColumn.PalleteLabel).Value)
                        itemCode1 = Convert.ToString(dgvPalette.Rows(j).Cells(GridColumn.ItemCode).Value)
                        If itemCode = itemCode1 Then
                            If Not Convert.ToBoolean(dgvPalette.Rows(j).Cells(GridColumn.Selection).Value) Then
                                MsgBox("Please select Palette in FIFO manner for Item Code : " & itemCode1, MsgBoxStyle.Information, "eMPro")
                                dgvPalette.Focus()
                                dgvPalette.CurrentCell = dgvPalette.Rows(j).Cells(GridColumn.Selection)
                                result = False
                                Return result
                            End If
                        End If
                    Next
                End If
            End If
        Next
        Return result
    End Function

    Private Sub rdbItemCode_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbItemCode.CheckedChanged, rdbItemDescription.CheckedChanged, rdbPaletteLabel.CheckedChanged
        Try
            txtSearch.Text = ""
            txtSearch.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FillSummaryGrid()
        Dim dt As New DataTable
        Try
            dgvSummary.Rows.Clear()

            If _itemCodes IsNot Nothing AndAlso _itemCodes.Rows.Count > 0 Then
                dt = _itemCodes.DefaultView.ToTable(True, "ITEM_CODE")
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim dv As DataView
                    dv = dt.DefaultView
                    dv.Sort = "ITEM_CODE"
                    dt = dv.ToTable()
                    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                        dgvSummary.Rows.Add(dt.Rows.Count)
                        Dim i As Integer = 0
                        For Each dr As DataRow In dt.Rows
                            dgvSummary.Rows(i).Cells(GridSummaryColumn.ItemCode).Value = dr("ITEM_CODE")
                            dgvSummary.Rows(i).Cells(GridSummaryColumn.ItemDescription).Value = Convert.ToString(SqlConnectionclass.ExecuteScalar("SELECT DESCRIPTION FROM ITEM_MST WHERE UNIT_CODE='" & gstrUnitId & "' AND ITEM_CODE='" & Convert.ToString(dr("ITEM_CODE")) & "'"))
                            dgvSummary.Rows(i).Cells(GridSummaryColumn.Qty).Value = 0
                            i += 1
                        Next
                    End If
                End If
            End If
            If _formMode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                For i As Integer = 0 To dgvPalette.Rows.Count - 1
                    dgvPalette.Rows(i).Cells(GridColumn.Selection).Value = True
                Next
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            dt.Dispose()
        End Try
    End Sub

    Private Sub CalculateSummary(ByVal itemCode As String, ByVal qty As Integer, ByVal operation As String)
        Try
            If dgvSummary.Rows.Count > 0 Then
                For i As Integer = 0 To dgvSummary.Rows.Count - 1
                    If itemCode = Convert.ToString(dgvSummary.Rows(i).Cells(GridSummaryColumn.ItemCode).Value) Then
                        If operation = "ADD" Then
                            dgvSummary.Rows(i).Cells(GridSummaryColumn.Qty).Value = Convert.ToInt32(dgvSummary.Rows(i).Cells(GridSummaryColumn.Qty).Value) + qty
                        Else
                            dgvSummary.Rows(i).Cells(GridSummaryColumn.Qty).Value = Convert.ToInt32(dgvSummary.Rows(i).Cells(GridSummaryColumn.Qty).Value) - qty
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub dgvPalette_CurrentCellDirtyStateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvPalette.CurrentCellDirtyStateChanged
        If dgvPalette.IsCurrentCellDirty Then
            dgvPalette.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub dgvPalette_CellValueChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPalette.CellValueChanged
        Try
            If e.RowIndex < 0 Then Exit Sub
            If e.ColumnIndex = GridColumn.Selection Then
                dgvPalette.CommitEdit(DataGridViewDataErrorContexts.Commit)
                If Convert.ToBoolean(dgvPalette.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) Then
                    CalculateSummary(Convert.ToString(dgvPalette.Rows(e.RowIndex).Cells(GridColumn.ItemCode).Value), Convert.ToInt32(dgvPalette.Rows(e.RowIndex).Cells(GridColumn.Qty).Value), "ADD")
                Else
                    CalculateSummary(Convert.ToString(dgvPalette.Rows(e.RowIndex).Cells(GridColumn.ItemCode).Value), Convert.ToInt32(dgvPalette.Rows(e.RowIndex).Cells(GridColumn.Qty).Value), "SUBTRACT")
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub SetPaletteItem()
        Try
            If _setItemQty IsNot Nothing AndAlso _setItemQty.Rows.Count > 0 Then
                For i As Integer = 0 To _setItemQty.Rows.Count - 1
                    For j As Integer = 0 To dgvPalette.Rows.Count - 1
                        If Convert.ToString(_setItemQty.Rows(i)("PALETTE_LABEL")) = Convert.ToString(dgvPalette.Rows(j).Cells(GridColumn.PalleteLabel).Value) And Convert.ToString(_setItemQty.Rows(i)("ITEM_CODE")) = Convert.ToString(dgvPalette.Rows(j).Cells(GridColumn.ItemCode).Value) Then
                            dgvPalette.Rows(j).Cells(GridColumn.Selection).Value = True
                            Exit For
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub CalculateQty()
        Try
            For i As Integer = 0 To dgvPalette.Rows.Count - 1
                If Convert.ToBoolean(dgvPalette.Rows(i).Cells(GridColumn.Selection).Value) Then
                    CalculateSummary(Convert.ToString(dgvPalette.Rows(i).Cells(GridColumn.ItemCode).Value), Convert.ToString(dgvPalette.Rows(i).Cells(GridColumn.Qty).Value), "ADD")
                End If
            Next
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
End Class