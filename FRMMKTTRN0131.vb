Imports System.IO
Imports System.Data.SqlClient
Imports ExcelAlias = Microsoft.Office.Interop.Excel

'*********************************************************************************************************************
'Copyright(c)       - MIND
'Name of Module     - HMIL Shortage & Picklist Generation
'Name of Form       - FRMMKTTRN0131  , HMIL Shortage & Picklist Generation
'Created by         - Ashish sharma
'Created Date       - 05 Feb 2024
'description        - To upload HMIL sub-daily requirements and GR report to create HMIL Shortage & Picklist Generation (New Development)
'*********************************************************************************************************************
'Modified by         - Ashish sharma
'Modified Date       - 06 May 2024
'description        - To upload HMIL sub-daily requirements and GR report for 24 hours from uploaded file time to create HMIL Shortage & Picklist Generation
'*********************************************************************************************************************

Public Class FRMMKTTRN0131
    Private Const SubDailyReqFileTotalColumns As Integer = 125
    Private Const SubDailyReqFileColumns As Integer = 43 ' Total column is 125 but needed only initial 43 columns
    Private Const GRFileColumns As Integer = 19
    Private Const SubDailyReqHeaderRowIndex As Byte = 1
    Private Const SubDailyReqDataRowIndex As Byte = 2
    Private Const GRHeaderRowIndex As Byte = 4
    Private Const GRDataRowIndex As Byte = 6
    Dim SubDailyReqExcelColumnName As String() = {"Base date", "Plant", "MRP Controller", "Material No.", _
                                        "ALC Code", "Materail Desc.", "ALT/SEL", "Material Status", "Del. Order", _
                                        "Final GR", "Closing Stock", "Open PO", "R/Point", _
                                        "Vendor", "Vendor Name", "Last Day G/I Qty", "Today.GI.Qty", "", "WIP Qty."}
    Dim GRExcelColumnName As String() = {"", "Plant", "Invoice No", "Status", "Material", _
                                        "GR Type", "Document H", "PO", "Inv.Qty", " GR Qty", _
                                        "Short Qty", "MRIR Rej.Q", "Gate In Dt", "Gate In Tm", _
                                        "Unit", "Description", "GR Documen", "GR Date", "GR Time"}
    Private Enum GridErrorsSubDailyReq
        Error_Desc = 0
        Source
    End Enum
    Private Enum GridErrorsGR
        Error_Desc = 0
        Source
    End Enum

    Private Enum GridShortage
        ItemCode = 0
        ItemDesc
        CustDrgNo
        UOM
        BinQty
        Qty
        PicklistQty
        DeliveryDate
    End Enum
    Private Sub FRMMKTTRN0131_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Try
            txtCustomerCode.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FRMMKTTRN0131_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Call FitToClient(Me, GrpMain, ctlHeader, GrpBoxButtons, 500)
            Me.MdiParent = mdifrmMain
            PictureBox1.Visible = False
            ConfigureGridColumn()
            btnPicklist.Enabled = False
            btnExport.Enabled = False
            txtCustomerCode.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub ConfigureGridColumn()
        Try
            dgvShortage.Columns.Clear()

            dgvShortage.Columns.Add("ItemCode", "Item Code")
            dgvShortage.Columns.Add("ItemDesc", "Description")
            dgvShortage.Columns.Add("CustDrgNo", "Cust.Drg.No.")
            dgvShortage.Columns.Add("UOM", "UOM")
            dgvShortage.Columns.Add("BinQty", "BinQty")
            dgvShortage.Columns.Add("Qty", "Qty")
            dgvShortage.Columns.Add("PicklistQty", "PicklistQty")
            dgvShortage.Columns.Add("DeliveryDate", "Delivery Date")

            dgvShortage.Columns(GridShortage.ItemCode).Width = 125
            dgvShortage.Columns(GridShortage.ItemDesc).Width = 250
            dgvShortage.Columns(GridShortage.CustDrgNo).Width = 125
            dgvShortage.Columns(GridShortage.UOM).Width = 40
            dgvShortage.Columns(GridShortage.BinQty).Width = 60
            dgvShortage.Columns(GridShortage.Qty).Width = 60
            dgvShortage.Columns(GridShortage.PicklistQty).Width = 80
            dgvShortage.Columns(GridShortage.DeliveryDate).Width = 100

            dgvShortage.Columns(GridShortage.ItemCode).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvShortage.Columns(GridShortage.ItemDesc).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvShortage.Columns(GridShortage.CustDrgNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvShortage.Columns(GridShortage.UOM).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvShortage.Columns(GridShortage.BinQty).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvShortage.Columns(GridShortage.Qty).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvShortage.Columns(GridShortage.PicklistQty).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvShortage.Columns(GridShortage.DeliveryDate).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            dgvShortage.Columns(GridShortage.ItemCode).ReadOnly = True
            dgvShortage.Columns(GridShortage.ItemDesc).ReadOnly = True
            dgvShortage.Columns(GridShortage.CustDrgNo).ReadOnly = True
            dgvShortage.Columns(GridShortage.UOM).ReadOnly = True
            dgvShortage.Columns(GridShortage.BinQty).ReadOnly = True
            dgvShortage.Columns(GridShortage.Qty).ReadOnly = True
            dgvShortage.Columns(GridShortage.PicklistQty).ReadOnly = True
            dgvShortage.Columns(GridShortage.DeliveryDate).ReadOnly = True

            dgvErrorsDetailSubDailyReq.DefaultCellStyle.WrapMode = DataGridViewTriState.True
            dgvErrorsDetailSubDailyReq.RowTemplate.Height = 30

            dgvErrorsDetailSubDailyReq.Columns.Clear()

            dgvErrorsDetailSubDailyReq.Columns.Add("ErrorDescription", "Error Description")
            dgvErrorsDetailSubDailyReq.Columns.Add("Source", "Source")

            dgvErrorsDetailSubDailyReq.Columns(GridErrorsSubDailyReq.Error_Desc).Width = 350
            dgvErrorsDetailSubDailyReq.Columns(GridErrorsSubDailyReq.Source).Width = 450

            dgvErrorsDetailSubDailyReq.Columns(GridErrorsSubDailyReq.Error_Desc).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvErrorsDetailSubDailyReq.Columns(GridErrorsSubDailyReq.Source).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            dgvErrorsDetailSubDailyReq.Columns(GridErrorsSubDailyReq.Error_Desc).ReadOnly = True
            dgvErrorsDetailSubDailyReq.Columns(GridErrorsSubDailyReq.Source).ReadOnly = True

            dgvErrorsDetailSubDailyReq.Columns(GridErrorsSubDailyReq.Error_Desc).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvErrorsDetailSubDailyReq.Columns(GridErrorsSubDailyReq.Source).SortMode = DataGridViewColumnSortMode.NotSortable

            dgvErrorsDetailSubDailyReq.DefaultCellStyle.WrapMode = DataGridViewTriState.True
            dgvErrorsDetailSubDailyReq.RowTemplate.Height = 30

            dgvErrorDetailGR.Columns.Clear()

            dgvErrorDetailGR.Columns.Add("ErrorDescription", "Error Description")
            dgvErrorDetailGR.Columns.Add("Source", "Source")

            dgvErrorDetailGR.Columns(GridErrorsGR.Error_Desc).Width = 350
            dgvErrorDetailGR.Columns(GridErrorsGR.Source).Width = 450

            dgvErrorDetailGR.Columns(GridErrorsGR.Error_Desc).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvErrorDetailGR.Columns(GridErrorsGR.Source).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            dgvErrorDetailGR.Columns(GridErrorsGR.Error_Desc).ReadOnly = True
            dgvErrorDetailGR.Columns(GridErrorsGR.Source).ReadOnly = True

            dgvErrorDetailGR.Columns(GridErrorsGR.Error_Desc).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvErrorDetailGR.Columns(GridErrorsGR.Source).SortMode = DataGridViewColumnSortMode.NotSortable

            dgvErrorDetailGR.DefaultCellStyle.WrapMode = DataGridViewTriState.True
            dgvErrorDetailGR.RowTemplate.Height = 30


        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub txtCustomerCode_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustomerCode.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Try
            Select Case KeyAscii
                Case System.Windows.Forms.Keys.Return
                    If btnBrowseSubDailyReq.Enabled = True Then btnBrowseSubDailyReq.Focus()
                Case 39, 34, 96
                    KeyAscii = 0
                Case 13
                    SendKeys.Send("{TAB}")
            End Select
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            e.KeyChar = Chr(KeyAscii)
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        End Try
    End Sub

    Private Sub txtCustomerCode_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyUp
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000
        Try
            If KeyCode = System.Windows.Forms.Keys.F1 Then
                If CmdCustCodeHelp.Enabled Then Call CmdCustCodeHelp_Click(CmdCustCodeHelp, New System.EventArgs())
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub txtCustomerCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustomerCode.TextChanged
        Try
            If String.IsNullOrEmpty(txtCustomerCode.Text) Then
                lblCustCodeDes.Text = String.Empty
                lblSubDailyRequirements.Tag = String.Empty
                lblSubDailyRequirements.Text = String.Empty
                lblGRReport.Tag = String.Empty
                lblGRReport.Text = String.Empty
                dgvShortage.Rows.Clear()
                dgvErrorsDetailSubDailyReq.Rows.Clear()
                dgvErrorDetailGR.Rows.Clear()
                btnPicklist.Enabled = False
                btnExport.Enabled = False
                btnUpload.Enabled = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtCustomerCode_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
        Dim dtCustomer As New DataTable
        Try
            If txtCustomerCode.Text.Trim.Length = 0 Then
                lblCustCodeDes.Text = String.Empty
            Else
                dtCustomer = SqlConnectionclass.GetDataTable("SELECT CUSTOMER_CODE,CUST_NAME FROM VW_GET_HMIL_CUSTOMER WHERE UNIT_CODE = '" & gstrUnitId & "' AND CUSTOMER_CODE='" + txtCustomerCode.Text.Trim() + "'")
                If dtCustomer IsNot Nothing AndAlso dtCustomer.Rows.Count > 0 Then
                    txtCustomerCode.Text = Convert.ToString(dtCustomer.Rows(0)("CUSTOMER_CODE"))
                    lblCustCodeDes.Text = Convert.ToString(dtCustomer.Rows(0)("CUST_NAME"))
                Else
                    MessageBox.Show("Customer code doesn't exist.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    txtCustomerCode.Text = String.Empty
                    lblCustCodeDes.Text = String.Empty
                    txtCustomerCode.Focus()
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            If dtCustomer IsNot Nothing Then
                dtCustomer.Dispose()
                dtCustomer = Nothing
            End If
        End Try
    End Sub

    Private Sub CmdCustCodeHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdCustCodeHelp.Click
        Dim strsql() As String
        Try
            strsql = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT CUSTOMER_CODE,CUST_NAME FROM VW_GET_HMIL_CUSTOMER WHERE UNIT_CODE ='" & gstrUnitId & "' ORDER BY CUSTOMER_CODE", "Customer(s) Help", 1)
            If Not (UBound(strsql) <= 0) Then
                If Not (UBound(strsql) = 0) Then
                    If (Len(strsql(0)) >= 1) And strsql(0) = "0" Then
                        MsgBox("No customer(s) record found.", MsgBoxStyle.Information, ResolveResString(100))
                        txtCustomerCode.Text = String.Empty
                        lblCustCodeDes.Text = String.Empty
                        txtCustomerCode.Focus()
                        Exit Sub
                    Else
                        txtCustomerCode.Text = strsql(0).Trim
                        lblCustCodeDes.Text = strsql(1).Trim
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            If MessageBox.Show("Are you sure to close?", ResolveResString(100), MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = Windows.Forms.DialogResult.Yes Then
                Me.Close()
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            txtCustomerCode.Text = String.Empty
            lblCustCodeDes.Text = String.Empty
            lblSubDailyRequirements.Tag = String.Empty
            lblSubDailyRequirements.Text = String.Empty
            lblGRReport.Tag = String.Empty
            lblGRReport.Text = String.Empty
            dgvErrorsDetailSubDailyReq.Rows.Clear()
            dgvErrorDetailGR.Rows.Clear()
            dgvShortage.Rows.Clear()
            btnPicklist.Enabled = False
            btnExport.Enabled = False
            btnUpload.Enabled = True
            PictureBox1.Visible = False
            tabControlShortage.SelectedTab = tabControlShortage.TabPages("tabPageShortage")
            txtCustomerCode.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnBrowseSubDailyReq_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowseSubDailyReq.Click
        Dim oFDialog As New OpenFileDialog()
        Try
            If String.IsNullOrEmpty(txtCustomerCode.Text.Trim()) Then
                MessageBox.Show("Please select a customer.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtCustomerCode.Text = String.Empty
                lblCustCodeDes.Text = String.Empty
                txtCustomerCode.Focus()
                Exit Sub
            End If

            dgvErrorsDetailSubDailyReq.DataSource = Nothing
            dgvErrorDetailGR.DataSource = Nothing
            dgvShortage.DataSource = Nothing
            Dim fileExtension As String = String.Empty
            Dim fileName As String = String.Empty

            oFDialog.Filter = "Excel Files|*.xls;*.xlsx"
            oFDialog.FilterIndex = 3
            oFDialog.RestoreDirectory = True
            oFDialog.Title = "Select a file to upload"
            If oFDialog.ShowDialog() = DialogResult.OK Then
                fileExtension = Path.GetExtension(oFDialog.FileName)
                fileName = Path.GetFileName(oFDialog.FileName)
                fileName = fileName.Replace(fileExtension, "")
                If (fileName.Length) > 4 Then
                    fileName = fileName.Substring(fileName.Length - 3, 2)
                Else
                    MessageBox.Show("Please select valid sub-daily requirements file, Invalid file name.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    lblSubDailyRequirements.Tag = String.Empty
                    lblSubDailyRequirements.Text = String.Empty
                    Exit Sub
                End If
                If fileName <> "00" And fileName <> "07" And fileName <> "09" And fileName <> "15" And fileName <> "17" Then
                    MessageBox.Show("Please select valid sub-daily requirements file, Invalid file name.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    lblSubDailyRequirements.Tag = String.Empty
                    lblSubDailyRequirements.Text = String.Empty
                    Exit Sub
                End If
                If String.IsNullOrEmpty(fileExtension) Then
                    MessageBox.Show("Please select valid extension file.Valid file extension is : .xls;xlsx", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    lblSubDailyRequirements.Tag = String.Empty
                    lblSubDailyRequirements.Text = String.Empty
                    Exit Sub
                End If
                If fileExtension.ToUpper() <> ".XLS" And fileExtension.ToUpper() <> ".XLSX" Then
                    MessageBox.Show("Please select valid extension file.Valid file extension is : .xls;xlsx", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    lblSubDailyRequirements.Tag = String.Empty
                    lblSubDailyRequirements.Text = String.Empty
                    Exit Sub
                End If
                lblSubDailyRequirements.Tag = oFDialog.SafeFileName
                lblSubDailyRequirements.Text = oFDialog.FileName
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub btnBrowseGRReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowseGRReport.Click
        Dim oFDialog As New OpenFileDialog()
        Try
            If String.IsNullOrEmpty(txtCustomerCode.Text.Trim()) Then
                MessageBox.Show("Please select a customer.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtCustomerCode.Text = String.Empty
                lblCustCodeDes.Text = String.Empty
                txtCustomerCode.Focus()
                Exit Sub
            End If

            If String.IsNullOrEmpty(lblSubDailyRequirements.Text.Trim()) Then
                MessageBox.Show("Please select a sub-daily requirements file to upload.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                lblSubDailyRequirements.Text = String.Empty
                lblSubDailyRequirements.Tag = String.Empty
                btnBrowseSubDailyReq.Focus()
                Exit Sub
            End If

            dgvErrorsDetailSubDailyReq.DataSource = Nothing
            dgvErrorDetailGR.DataSource = Nothing
            dgvShortage.DataSource = Nothing
            Dim fileExtension As String = String.Empty

            oFDialog.Filter = "Excel Files|*.xls;*.xlsx"
            oFDialog.FilterIndex = 3
            oFDialog.RestoreDirectory = True
            oFDialog.Title = "Select a file to upload"
            If oFDialog.ShowDialog() = DialogResult.OK Then
                fileExtension = Path.GetExtension(oFDialog.FileName)
                If String.IsNullOrEmpty(fileExtension) Then
                    MessageBox.Show("Please select valid extension file.Valid file extension is : .xls;xlsx", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    lblGRReport.Tag = String.Empty
                    lblGRReport.Text = String.Empty
                    Exit Sub
                End If
                If fileExtension.ToUpper() <> ".XLS" And fileExtension.ToUpper() <> ".XLSX" Then
                    MessageBox.Show("Please select valid extension file.Valid file extension is : .xls;xlsx", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    lblGRReport.Tag = String.Empty
                    lblGRReport.Text = String.Empty
                    Exit Sub
                End If
                lblGRReport.Tag = oFDialog.SafeFileName
                lblGRReport.Text = oFDialog.FileName
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub btnUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        Dim dtSubDailyReqFileData As New DataTable
        Dim dtGRFileData As New DataTable
        Dim dtShortage As New DataTable
        Dim sqlCmd As New SqlCommand
        Dim i As Integer = 0
        Try
            If String.IsNullOrEmpty(txtCustomerCode.Text.Trim()) Then
                MessageBox.Show("Please select a customer.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtCustomerCode.Text = String.Empty
                lblCustCodeDes.Text = String.Empty
                txtCustomerCode.Focus()
                Exit Sub
            End If

            If String.IsNullOrEmpty(lblSubDailyRequirements.Text.Trim()) Then
                MessageBox.Show("Please select a sub-daily requirements file to upload.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                lblSubDailyRequirements.Text = String.Empty
                lblSubDailyRequirements.Tag = String.Empty
                btnBrowseSubDailyReq.Focus()
                Exit Sub
            End If

            If String.IsNullOrEmpty(lblGRReport.Text.Trim()) Then
                MessageBox.Show("Please select a GR report file to upload.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                lblGRReport.Text = String.Empty
                lblGRReport.Tag = String.Empty
                btnBrowseGRReport.Focus()
                Exit Sub
            End If

            If MessageBox.Show("Are you sure to upload?", ResolveResString(100), MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            dgvErrorsDetailSubDailyReq.Rows.Clear()
            dgvErrorDetailGR.Rows.Clear()
            dgvShortage.Rows.Clear()

            PictureBox1.Visible = True

            dtSubDailyReqFileData = GetSubDailyReqExcelFileDataIntoDatatable()
            dtGRFileData = GetGRExcelFileDataIntoDatatable()

            Dim strSlot As String = String.Empty
            Dim fileExtension As String = Path.GetExtension(lblSubDailyRequirements.Text)
            strSlot = Path.GetFileName(lblSubDailyRequirements.Text)
            strSlot = strSlot.Replace(fileExtension, "")
            If (strSlot.Length) > 4 Then
                strSlot = strSlot.Substring(strSlot.Length - 3, 2)
            End If

            If dtSubDailyReqFileData IsNot Nothing AndAlso dtSubDailyReqFileData.Rows.Count > 0 Then
                If dtGRFileData IsNot Nothing AndAlso dtGRFileData.Rows.Count > 0 Then
                    With sqlCmd
                        .CommandType = CommandType.StoredProcedure
                        .CommandTimeout = 600 ' 10 Minute
                        .CommandText = "USP_HMIL_DISPATCH_SHORTAGE_PLAN"
                        .Parameters.Clear()
                        .Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
                        .Parameters.AddWithValue("@CUSTOMER_CODE", txtCustomerCode.Text.Trim())
                        .Parameters.AddWithValue("@SLOT", strSlot)
                        .Parameters.AddWithValue("@SUB_DAILY_REQ_FILE_NAME", lblSubDailyRequirements.Tag)
                        .Parameters.AddWithValue("@GR_FILE_NAME", lblGRReport.Tag)
                        .Parameters.AddWithValue("@IP_ADDRESS", gstrIpaddressWinSck)
                        .Parameters.AddWithValue("@USER_ID", mP_User)
                        .Parameters.AddWithValue("@UDT_HMIL_SUB_DAILY_REQUIREMENTS", dtSubDailyReqFileData)
                        .Parameters.AddWithValue("@UDT_HMIL_GR_REPORT", dtGRFileData)
                        .Parameters.Add("@PLAN_ID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                        .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                        dtShortage.Load(SqlConnectionclass.ExecuteReader(sqlCmd))
                        If Convert.ToString(.Parameters("@MESSAGE").Value) <> String.Empty Then
                            PopulateLog()
                            PictureBox1.Visible = False
                            MessageBox.Show(Convert.ToString(.Parameters("@MESSAGE").Value), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        Else
                            i = 0
                            If dtShortage IsNot Nothing AndAlso dtShortage.Rows.Count > 0 Then
                                dgvShortage.Rows.Clear()
                                dgvShortage.Rows.Add(dtShortage.Rows.Count)
                                For Each dr As DataRow In dtShortage.Rows
                                    dgvShortage.Rows(i).Cells(GridShortage.ItemCode).Value = Convert.ToString(dr("ITEM_CODE"))
                                    dgvShortage.Rows(i).Cells(GridShortage.CustDrgNo).Value = Convert.ToString(dr("CUST_DRGNO"))
                                    dgvShortage.Rows(i).Cells(GridShortage.ItemDesc).Value = Convert.ToString(dr("PART_NAME"))
                                    dgvShortage.Rows(i).Cells(GridShortage.UOM).Value = Convert.ToString(dr("UOM"))
                                    dgvShortage.Rows(i).Cells(GridShortage.BinQty).Value = Convert.ToString(dr("SNP"))
                                    dgvShortage.Rows(i).Cells(GridShortage.Qty).Value = Convert.ToString(dr("QTY"))
                                    dgvShortage.Rows(i).Cells(GridShortage.PicklistQty).Value = Convert.ToString(dr("PICKLIST_QTY"))
                                    dgvShortage.Rows(i).Cells(GridShortage.DeliveryDate).Value = Convert.ToString(dr("DELIVERY_DATE"))
                                    i += 1
                                Next
                                btnPicklist.Enabled = True
                                btnExport.Enabled = True
                            End If
                            tabControlShortage.SelectedTab = tabControlShortage.TabPages("tabPageShortage")

                            PictureBox1.Visible = False
                            lblSubDailyRequirements.Text = String.Empty
                            lblSubDailyRequirements.Tag = String.Empty
                            lblGRReport.Text = String.Empty
                            lblGRReport.Tag = String.Empty
                            'txtCustomerCode.Text = String.Empty
                            'lblCustCodeDes.Text = String.Empty
                            btnUpload.Enabled = False
                            MessageBox.Show("HMIL File(s) have been Uploaded Successfully !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    End With
                Else
                    MessageBox.Show("No record(s) found in GR file.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    lblGRReport.Text = String.Empty
                    lblGRReport.Tag = String.Empty
                    btnBrowseGRReport.Focus()
                End If
            Else
                MessageBox.Show("No record(s) found in Sub-Daily Requirement file.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                lblSubDailyRequirements.Text = String.Empty
                lblSubDailyRequirements.Tag = String.Empty
                lblGRReport.Text = String.Empty
                lblGRReport.Tag = String.Empty
                btnBrowseSubDailyReq.Focus()
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            PictureBox1.Visible = False
            If sqlCmd IsNot Nothing Then
                sqlCmd.Dispose()
                sqlCmd = Nothing
            End If
            If dtSubDailyReqFileData IsNot Nothing Then
                dtSubDailyReqFileData.Dispose()
                dtSubDailyReqFileData = Nothing
            End If
            If dtGRFileData IsNot Nothing Then
                dtGRFileData.Dispose()
                dtGRFileData = Nothing
            End If
            If dtShortage IsNot Nothing Then
                dtShortage.Dispose()
                dtShortage = Nothing
            End If
        End Try
    End Sub
    Private Function GetSubDailyReqExcelFileDataIntoDatatable() As DataTable
        Dim dtFileData As New DataTable
        Dim xlApp As ExcelAlias.Application
        Dim xlWorkBook As ExcelAlias.Workbook
        Dim xlWorkSheet As ExcelAlias.Worksheet
        Try
            xlApp = New ExcelAlias.ApplicationClass
            xlWorkBook = xlApp.Workbooks.Open(lblSubDailyRequirements.Text)
            xlWorkSheet = xlWorkBook.ActiveSheet

            Dim data As Object(,) = DirectCast(xlWorkSheet.UsedRange.Value2, Object(,))

            If Not ValidateSubDailyReqExcelFileColumn(data) Then
                lblSubDailyRequirements.Text = String.Empty
                lblSubDailyRequirements.Tag = String.Empty
                Return dtFileData
            End If

            For col As Integer = 1 To SubDailyReqFileColumns
                If data(SubDailyReqHeaderRowIndex, col).ToString.Trim.ToUpper = "BASE DATE" _
                                    OrElse data(SubDailyReqHeaderRowIndex, col).ToString.Trim.ToUpper = "PLANT" _
                                    OrElse data(SubDailyReqHeaderRowIndex, col).ToString.Trim.ToUpper = "MRP CONTROLLER" _
                                    OrElse data(SubDailyReqHeaderRowIndex, col).ToString.Trim.ToUpper = "MATERIAL NO." _
                                    OrElse data(SubDailyReqHeaderRowIndex, col).ToString.Trim.ToUpper = "ALC CODE" _
                                    OrElse data(SubDailyReqHeaderRowIndex, col).ToString.Trim.ToUpper = "MATERAIL DESC." _
                                    OrElse data(SubDailyReqHeaderRowIndex, col).ToString.Trim.ToUpper = "ALT/SEL" _
                                    OrElse data(SubDailyReqHeaderRowIndex, col).ToString.Trim.ToUpper = "MATERIAL STATUS" _
                                    OrElse data(SubDailyReqHeaderRowIndex, col).ToString.Trim.ToUpper = "DEL. ORDER" _
                                    OrElse data(SubDailyReqHeaderRowIndex, col).ToString.Trim.ToUpper = "R/POINT" _
                                    OrElse data(SubDailyReqHeaderRowIndex, col).ToString.Trim.ToUpper = "VENDOR" _
                                    OrElse data(SubDailyReqHeaderRowIndex, col).ToString.Trim.ToUpper = "VENDOR NAME" _
                                    OrElse data(SubDailyReqHeaderRowIndex, col).ToString.Trim.ToUpper = "UNIT" Then
                    dtFileData.Columns.Add(data(SubDailyReqHeaderRowIndex, col).ToString.Trim.ToUpper, GetType(System.String))
                Else
                    If col = 18 Then
                        dtFileData.Columns.Add("BLANK_HEADER_1", GetType(System.Decimal))
                    ElseIf col = 20 Then
                        dtFileData.Columns.Add("00", GetType(System.Decimal))
                    ElseIf col = 32 Then
                        dtFileData.Columns.Add("100", GetType(System.Decimal))
                    ElseIf col > 32 AndAlso col < 44 Then
                        dtFileData.Columns.Add("1" & data(SubDailyReqHeaderRowIndex, col).ToString.Trim.ToUpper, GetType(System.Decimal))
                    Else
                        dtFileData.Columns.Add(data(SubDailyReqHeaderRowIndex, col).ToString.Trim.ToUpper, GetType(System.Decimal))
                    End If
                End If
            Next

            For row As Integer = SubDailyReqDataRowIndex To data.GetUpperBound(0)
                Dim newDataRow As DataRow = dtFileData.NewRow()
                For col As Integer = 1 To SubDailyReqFileColumns
                    newDataRow(col - 1) = data(row, col)
                Next
                dtFileData.Rows.Add(newDataRow)
            Next
        Catch ex As Exception
            Throw
        Finally
            xlWorkBook.Close()
            xlApp.Quit()

            ReleaseComObject(xlApp)
            ReleaseComObject(xlWorkBook)
            ReleaseComObject(xlWorkSheet)
        End Try
        Return dtFileData
    End Function

    Private Function GetGRExcelFileDataIntoDatatable() As DataTable
        Dim dtFileData As New DataTable
        Dim xlApp As ExcelAlias.Application
        Dim xlWorkBook As ExcelAlias.Workbook
        Dim xlWorkSheet As ExcelAlias.Worksheet
        Try
            xlApp = New ExcelAlias.ApplicationClass
            xlWorkBook = xlApp.Workbooks.Open(lblGRReport.Text)
            xlWorkSheet = xlWorkBook.ActiveSheet

            Dim data As Object(,) = DirectCast(xlWorkSheet.UsedRange.Value2, Object(,))

            If Not ValidateGRExcelFileColumn(data) Then
                lblGRReport.Text = String.Empty
                lblGRReport.Tag = String.Empty
                Return dtFileData
            End If

            For col As Integer = 2 To GRFileColumns
                If data(GRHeaderRowIndex, col).ToString.Trim.ToUpper = "PLANT" _
                                    OrElse data(GRHeaderRowIndex, col).ToString.Trim.ToUpper = "INVOICE NO" _
                                    OrElse data(GRHeaderRowIndex, col).ToString.Trim.ToUpper = "STATUS" _
                                    OrElse data(GRHeaderRowIndex, col).ToString.Trim.ToUpper = "MATERIAL" _
                                    OrElse data(GRHeaderRowIndex, col).ToString.Trim.ToUpper = "GR TYPE" _
                                    OrElse data(GRHeaderRowIndex, col).ToString.Trim.ToUpper = "DOCUMENT H" _
                                    OrElse data(GRHeaderRowIndex, col).ToString.Trim.ToUpper = "PO" _
                                    OrElse data(GRHeaderRowIndex, col).ToString.Trim.ToUpper = "GATE IN DT" _
                                    OrElse data(GRHeaderRowIndex, col).ToString.Trim.ToUpper = "GATE IN TM" _
                                    OrElse data(GRHeaderRowIndex, col).ToString.Trim.ToUpper = "UNIT" _
                                    OrElse data(GRHeaderRowIndex, col).ToString.Trim.ToUpper = "DESCRIPTION" _
                                    OrElse data(GRHeaderRowIndex, col).ToString.Trim.ToUpper = "GR DOCUMEN" _
                                    OrElse data(GRHeaderRowIndex, col).ToString.Trim.ToUpper = "GR DATE" _
                                    OrElse data(GRHeaderRowIndex, col).ToString.Trim.ToUpper = "GR TIME" Then
                    dtFileData.Columns.Add(data(GRHeaderRowIndex, col).ToString.Trim.ToUpper, GetType(System.String))
                Else
                    dtFileData.Columns.Add(data(GRHeaderRowIndex, col).ToString.Trim.ToUpper, GetType(System.Decimal))
                End If
            Next

            For row As Integer = GRDataRowIndex To data.GetUpperBound(0)
                Dim newDataRow As DataRow = dtFileData.NewRow()
                For col As Integer = 2 To GRFileColumns
                    newDataRow(col - 2) = data(row, col)
                Next
                dtFileData.Rows.Add(newDataRow)
            Next
        Catch ex As Exception
            Throw
        Finally
            xlWorkBook.Close()
            xlApp.Quit()

            ReleaseComObject(xlApp)
            ReleaseComObject(xlWorkBook)
            ReleaseComObject(xlWorkSheet)
        End Try
        Return dtFileData
    End Function

    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Function ValidateSubDailyReqExcelFileColumn(ByVal fileData As Object(,)) As Boolean
        Dim result As Boolean = True
        If fileData Is Nothing Then
            MsgBox("Browse Sub-Daily Requirement file does not contain data.")
            result = False
            Return result
        End If
        If fileData.GetUpperBound(0) < SubDailyReqDataRowIndex Then
            MsgBox("Browse Sub-Daily Requirement file data is not in HMIL Sub-Daily Requirement Excel format, Please upload correct formatted file.")
            result = False
            Return result
        End If
        If fileData.GetUpperBound(1) < SubDailyReqFileTotalColumns Then
            MsgBox("Browse Sub-Daily Requirement file data is not in HMIL Sub-Daily Requirement Excel format, Please upload correct formatted file.")
            result = False
            Return result
        End If
        For col As Integer = 1 To SubDailyReqFileColumns
            If col > 19 Then Exit For

            If fileData(SubDailyReqHeaderRowIndex, col).ToString.ToUpper.Trim <> SubDailyReqExcelColumnName(col - 1).ToString.ToUpper.Trim Then
                MsgBox("Browse Sub-Daily Requirement file data is not in HMIL Sub-Daily Requirement Excel format, Please upload correct formatted file." & vbCrLf & "Excel file does not contain column : " & SubDailyReqExcelColumnName(col - 1).ToString.ToUpper.Trim)
                result = False
                Return result
            End If
        Next
        Return result
    End Function

    Private Function ValidateGRExcelFileColumn(ByVal fileData As Object(,)) As Boolean
        Dim result As Boolean = True

        If fileData Is Nothing Then
            MsgBox("Browse GR file does not contain data.")
            result = False
            Return result
        End If
        If fileData.GetUpperBound(0) < GRDataRowIndex Then
            MsgBox("Browse GR file data is not in HMIL GR Excel format, Please upload correct formatted file.")
            result = False
            Return result
        End If
        If fileData.GetUpperBound(1) < GRFileColumns Then
            MsgBox("Browse GR file data is not in HMIL GR Excel format, Please upload correct formatted file.")
            result = False
            Return result
        End If
        For col As Integer = 1 To GRFileColumns
            If If(fileData(GRHeaderRowIndex, col) Is Nothing, "", fileData(GRHeaderRowIndex, col).ToString.ToUpper.Trim) <> GRExcelColumnName(col - 1).ToString.ToUpper.Trim Then
                MsgBox("Browse GR file data is not in HMIL GR Excel format, Please upload correct formatted file." & vbCrLf & "Excel file does not contain column : " & GRExcelColumnName(col - 1).ToString.ToUpper.Trim)
                result = False
                Return result
            End If
        Next

        Return result
    End Function
    Private Sub PopulateLog()
        Dim dtErrors As New DataTable
        Try
            Dim i As Integer = 0
            dtErrors = SqlConnectionclass.GetDataTable("SELECT ERR_DESC,SOURCE_DESC FROM HMIL_DISPATCH_PLAN_ERROR_LOG WITH (NOLOCK) WHERE UNIT_CODE = '" & gstrUnitId & "' AND  IP_ADDRESS='" & gstrIpaddressWinSck & "' AND LOG_TYPE='Sub-Daily' ORDER BY LOG_ID ")
            If dtErrors IsNot Nothing AndAlso dtErrors.Rows.Count > 0 Then
                dgvErrorsDetailSubDailyReq.Rows.Clear()
                dgvErrorsDetailSubDailyReq.Rows.Add(dtErrors.Rows.Count)
                For Each dr As DataRow In dtErrors.Rows
                    dgvErrorsDetailSubDailyReq.Rows(i).Cells(GridErrorsSubDailyReq.Error_Desc).Value = Convert.ToString(dr("ERR_DESC"))
                    dgvErrorsDetailSubDailyReq.Rows(i).Cells(GridErrorsSubDailyReq.Source).Value = Convert.ToString(dr("SOURCE_DESC"))
                    i += 1
                Next
                btnExport.Enabled = True
                tabControlShortage.SelectedTab = tabControlShortage.TabPages("tabPageErrorsSubDailyReq")
            End If
            i = 0
            dtErrors = SqlConnectionclass.GetDataTable("SELECT ERR_DESC,SOURCE_DESC FROM HMIL_DISPATCH_PLAN_ERROR_LOG WITH (NOLOCK) WHERE UNIT_CODE = '" & gstrUnitId & "' AND  IP_ADDRESS='" & gstrIpaddressWinSck & "' AND LOG_TYPE='GR' ORDER BY LOG_ID ")
            If dtErrors IsNot Nothing AndAlso dtErrors.Rows.Count > 0 Then
                dgvErrorDetailGR.Rows.Clear()
                dgvErrorDetailGR.Rows.Add(dtErrors.Rows.Count)
                For Each dr As DataRow In dtErrors.Rows
                    dgvErrorDetailGR.Rows(i).Cells(GridErrorsGR.Error_Desc).Value = Convert.ToString(dr("ERR_DESC"))
                    dgvErrorDetailGR.Rows(i).Cells(GridErrorsGR.Source).Value = Convert.ToString(dr("SOURCE_DESC"))
                    i += 1
                Next
                btnExport.Enabled = True
                tabControlShortage.SelectedTab = tabControlShortage.TabPages("tabPageErrorsGR")
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            If (dtErrors IsNot Nothing) Then
                dtErrors.Dispose()
                dtErrors = Nothing
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub btnPicklist_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPicklist.Click
        Dim sqlCmd As New SqlCommand
        Try
            If dgvErrorsDetailSubDailyReq IsNot Nothing AndAlso dgvErrorsDetailSubDailyReq.Rows.Count > 0 Then
                MessageBox.Show("Please check sub-daily requirements error(s) detail tab.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                tabControlShortage.SelectedTab = tabControlShortage.TabPages("tabPageErrorsSubDailyReq")
                Exit Sub
            End If
            If dgvErrorDetailGR IsNot Nothing AndAlso dgvErrorDetailGR.Rows.Count > 0 Then
                MessageBox.Show("Please check GR error(s) detail tab.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                tabControlShortage.SelectedTab = tabControlShortage.TabPages("tabPageErrorsGR")
                Exit Sub
            End If
            If dgvShortage Is Nothing OrElse dgvShortage.Rows.Count = 0 Then
                MessageBox.Show("No record(s) found in shortage tab.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                tabControlShortage.SelectedTab = tabControlShortage.TabPages("tabPageShortage")
                Exit Sub
            End If
            If String.IsNullOrEmpty(txtCustomerCode.Text.Trim()) Then
                MessageBox.Show("Please select a customer.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtCustomerCode.Text = String.Empty
                lblCustCodeDes.Text = String.Empty
                txtCustomerCode.Focus()
                Exit Sub
            End If

            If MessageBox.Show("Are you sure to create picklist?", ResolveResString(100), MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            PictureBox1.Visible = True

            With sqlCmd
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 600 ' 10 Minute
                .CommandText = "USP_HMIL_DISPATCH_PICKLIST"
                .Parameters.Clear()
                .Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
                .Parameters.AddWithValue("@CUSTOMER_CODE", txtCustomerCode.Text.Trim())
                .Parameters.AddWithValue("@IP_ADDRESS", gstrIpaddressWinSck)
                .Parameters.AddWithValue("@USER_ID", mP_User)
                .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                If Convert.ToString(.Parameters("@MESSAGE").Value) <> String.Empty Then
                    PictureBox1.Visible = False
                    MessageBox.Show(Convert.ToString(.Parameters("@MESSAGE").Value), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                Else
                    PictureBox1.Visible = False
                    lblSubDailyRequirements.Tag = String.Empty
                    lblSubDailyRequirements.Text = String.Empty
                    lblGRReport.Tag = String.Empty
                    lblGRReport.Text = String.Empty
                    dgvErrorsDetailSubDailyReq.Rows.Clear()
                    dgvErrorDetailGR.Rows.Clear()
                    dgvShortage.Rows.Clear()
                    btnPicklist.Enabled = False
                    btnExport.Enabled = False
                    btnUpload.Enabled = True
                    PictureBox1.Visible = False
                    tabControlShortage.SelectedTab = tabControlShortage.TabPages("tabPageShortage")
                    txtCustomerCode.Focus()
                    MessageBox.Show("Picklist created successfully !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            PictureBox1.Visible = False
            If sqlCmd IsNot Nothing Then
                sqlCmd.Dispose()
                sqlCmd = Nothing
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
        Dim xlApp As ExcelAlias.Application
        Dim xlWorkBook As ExcelAlias.Workbook
        Dim xlWorkSheet As ExcelAlias.Worksheet
        Dim r As Int32, c As Int32
        Try
            If (dgvShortage Is Nothing OrElse dgvShortage.Rows.Count = 0) And (dgvErrorsDetailSubDailyReq Is Nothing OrElse dgvErrorsDetailSubDailyReq.Rows.Count = 0) And (dgvErrorDetailGR Is Nothing OrElse dgvErrorDetailGR.Rows.Count = 0) Then
                MessageBox.Show("No record(s) found.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                tabControlShortage.SelectedTab = tabControlShortage.TabPages("tabPageShortage")
                Exit Sub
            End If
            xlApp = New ExcelAlias.ApplicationClass

            If dgvShortage IsNot Nothing AndAlso dgvShortage.Rows.Count > 0 Then
                xlWorkBook = xlApp.Workbooks.Add()
                Dim arr(dgvShortage.Rows.Count, dgvShortage.Columns.Count) As Object
                xlWorkSheet = xlWorkBook.Worksheets(1)
                xlWorkSheet.Name = tabControlShortage.TabPages("tabPageShortage").Text
                For r = 0 To dgvShortage.Rows.Count - 1
                    For c = 0 To dgvShortage.Columns.Count - 1
                        arr(r, c) = dgvShortage.Rows(r).Cells(c).Value
                    Next
                Next
                c = 0
                For Each column As DataGridViewColumn In dgvShortage.Columns
                    xlWorkSheet.Cells(1, c + 1) = column.Name
                    c += 1
                Next
                'add the data starting in cell A2
                xlWorkSheet.Range(xlWorkSheet.Cells(2, 1), xlWorkSheet.Cells(dgvShortage.Rows.Count + 1, dgvShortage.Columns.Count)).Value = arr
                xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(1, dgvShortage.Columns.Count)).Font.Bold = True
                xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(dgvShortage.Rows.Count + 1, dgvShortage.Columns.Count)).EntireColumn.AutoFit()
                xlApp.Visible = True
            Else
                If dgvErrorsDetailSubDailyReq IsNot Nothing AndAlso dgvErrorsDetailSubDailyReq.Rows.Count > 0 Then
                    xlWorkBook = xlApp.Workbooks.Add()
                    Dim arrSub(dgvErrorsDetailSubDailyReq.Rows.Count, dgvErrorsDetailSubDailyReq.Columns.Count) As Object
                    xlWorkSheet = xlWorkBook.Worksheets(1)
                    xlWorkSheet.Name = "SubDailyReqError(s)"
                    For r = 0 To dgvErrorsDetailSubDailyReq.Rows.Count - 1
                        For c = 0 To dgvErrorsDetailSubDailyReq.Columns.Count - 1
                            arrSub(r, c) = dgvErrorsDetailSubDailyReq.Rows(r).Cells(c).Value
                        Next
                    Next
                    c = 0
                    For Each column As DataGridViewColumn In dgvErrorsDetailSubDailyReq.Columns
                        xlWorkSheet.Cells(1, c + 1) = column.Name
                        c += 1
                    Next

                    xlWorkSheet.Range(xlWorkSheet.Cells(2, 1), xlWorkSheet.Cells(dgvErrorsDetailSubDailyReq.Rows.Count + 1, dgvErrorsDetailSubDailyReq.Columns.Count)).Value = arrSub
                    xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(1, dgvErrorsDetailSubDailyReq.Columns.Count)).Font.Bold = True
                    xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(dgvErrorsDetailSubDailyReq.Rows.Count + 1, dgvErrorsDetailSubDailyReq.Columns.Count)).EntireColumn.AutoFit()
                End If
                If dgvErrorDetailGR IsNot Nothing AndAlso dgvErrorDetailGR.Rows.Count > 0 Then
                    If xlWorkBook Is Nothing Then
                        xlWorkBook = xlApp.Workbooks.Add()
                        xlWorkSheet = xlWorkBook.Worksheets(1)
                    Else
                        xlWorkBook.Sheets.Add(After:=xlWorkBook.Worksheets(xlWorkBook.Worksheets.Count))
                    End If
                    Dim arrGR(dgvErrorDetailGR.Rows.Count, dgvErrorDetailGR.Columns.Count) As Object
                    xlWorkSheet = xlWorkBook.Worksheets.Add()
                    xlWorkSheet.Name = "GRError(s)"
                    For r = 0 To dgvErrorDetailGR.Rows.Count - 1
                        For c = 0 To dgvErrorDetailGR.Columns.Count - 1
                            arrGR(r, c) = dgvErrorDetailGR.Rows(r).Cells(c).Value
                        Next
                    Next
                    c = 0
                    For Each column As DataGridViewColumn In dgvErrorDetailGR.Columns
                        xlWorkSheet.Cells(1, c + 1) = column.Name
                        c += 1
                    Next

                    xlWorkSheet.Range(xlWorkSheet.Cells(2, 1), xlWorkSheet.Cells(dgvErrorDetailGR.Rows.Count + 1, dgvErrorDetailGR.Columns.Count)).Value = arrGR
                    xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(1, dgvErrorsDetailSubDailyReq.Columns.Count)).Font.Bold = True
                    xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(dgvErrorDetailGR.Rows.Count + 1, dgvErrorDetailGR.Columns.Count)).EntireColumn.AutoFit()
                End If
                xlApp.Visible = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            ReleaseComObject(xlApp)
            ReleaseComObject(xlWorkBook)
            ReleaseComObject(xlWorkSheet)

            xlWorkSheet = Nothing
            xlWorkBook = Nothing
            xlApp = Nothing
        End Try
    End Sub

    Private Sub btnPicklistReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPicklistReport.Click
        Dim dtPicklistData As New DataTable
        Dim sqlCmd As New SqlCommand
        Try
            If String.IsNullOrEmpty(txtCustomerCode.Text.Trim()) Then
                MessageBox.Show("Please select a customer.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtCustomerCode.Text = String.Empty
                lblCustCodeDes.Text = String.Empty
                txtCustomerCode.Focus()
                Exit Sub
            End If

            PictureBox1.Visible = True

            With sqlCmd
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 600 ' 10 Minute
                .CommandText = "USP_HMIL_DISPATCH_PICKLIST_EXCEL_RPT"
                .Parameters.Clear()
                .Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
                .Parameters.AddWithValue("@CUSTOMER_CODE", txtCustomerCode.Text.Trim())
                .Parameters.AddWithValue("@IP_ADDRESS", gstrIpaddressWinSck)
                .Parameters.AddWithValue("@USER_ID", mP_User)
                .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                dtPicklistData.Load(SqlConnectionclass.ExecuteReader(sqlCmd))
                If Convert.ToString(.Parameters("@MESSAGE").Value) <> String.Empty Then
                    PictureBox1.Visible = False
                    MessageBox.Show(Convert.ToString(.Parameters("@MESSAGE").Value), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                Else
                    If dtPicklistData IsNot Nothing AndAlso dtPicklistData.Rows.Count > 0 Then
                        GeneratePicklistReport(dtPicklistData)
                    Else
                        PictureBox1.Visible = False
                        MessageBox.Show("No record(s) found.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Exit Sub
                    End If
                End If
            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            PictureBox1.Visible = False
            If sqlCmd IsNot Nothing Then
                sqlCmd.Dispose()
                sqlCmd = Nothing
            End If
            If dtPicklistData IsNot Nothing Then
                dtPicklistData.Dispose()
                dtPicklistData = Nothing
            End If
        End Try
    End Sub

    Private Sub GeneratePicklistReport(ByVal dtPicklistData As DataTable)
        Dim xlApp As ExcelAlias.Application
        Dim xlWorkBook As ExcelAlias.Workbook
        Dim xlWorkSheet As ExcelAlias.Worksheet
        Dim r As Int32, c As Int32
        Try
            xlApp = New ExcelAlias.ApplicationClass
            If dtPicklistData IsNot Nothing AndAlso dtPicklistData.Rows.Count > 0 Then
                xlWorkBook = xlApp.Workbooks.Add()

                Dim arr(dtPicklistData.Rows.Count, dtPicklistData.Columns.Count) As Object
                xlWorkSheet = xlWorkBook.Worksheets(1)
                xlWorkSheet.Name = "PicklistReport"
                For r = 0 To dtPicklistData.Rows.Count - 1
                    For c = 0 To dtPicklistData.Columns.Count - 1
                        arr(r, c) = dtPicklistData.Rows(r)(c)
                    Next
                Next
                c = 0
                For Each column As DataColumn In dtPicklistData.Columns
                    If c > 5 And c < 19 Then
                        xlWorkSheet.Cells(1, c + 1) = "'" & column.ColumnName.Replace("A", String.Empty)
                    Else
                        xlWorkSheet.Cells(1, c + 1) = "'" & column.ColumnName
                    End If
                    c += 1
                Next
                'add the data starting in cell A2
                xlWorkSheet.Range(xlWorkSheet.Cells(2, 1), xlWorkSheet.Cells(dtPicklistData.Rows.Count + 1, dtPicklistData.Columns.Count)).Value = arr
                xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(1, dtPicklistData.Columns.Count)).Font.Bold = True
                xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(dtPicklistData.Rows.Count + 1, dtPicklistData.Columns.Count)).EntireColumn.AutoFit()
                xlApp.Visible = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            PictureBox1.Visible = False
            ReleaseComObject(xlApp)
            ReleaseComObject(xlWorkBook)
            ReleaseComObject(xlWorkSheet)

            xlWorkSheet = Nothing
            xlWorkBook = Nothing
            xlApp = Nothing
        End Try
    End Sub
End Class