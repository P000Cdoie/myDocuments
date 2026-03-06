Imports System.IO
Imports System.Data.SqlClient

'*********************************************************************************************************************
'Copyright(c)       - MIND
'Name of Module     - TKML 560B PDS Uploading & Picklist Generation
'Name of Form       - FRMMKTTRN0126  , TKML 560B PDS Uploading & Picklist Generation
'Created by         - Ashish sharma
'Created Date       - 25 Nov 2022
'description        - To upload TKML 560B PDS file uploading & Picklist Generation (New Development)
'*********************************************************************************************************************

Public Class FRMMKTTRN0126
    Private Enum GridErrors
        Error_Desc = 0
        Source
    End Enum
    Private Sub FRMMKTTRN0126_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            txtCustomerCode.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FRMMKTTRN0126_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Call FitToClient(Me, GrpMain, ctlHeader, GrpBoxButtons, 500)
            Me.MdiParent = mdifrmMain
            PictureBox1.Visible = False
            ConfigureGridColumn()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub ConfigureGridColumn()
        Try
            dgvErrorsDetail.Columns.Clear()

            dgvErrorsDetail.Columns.Add("ErrorDescription", "Error Description")
            dgvErrorsDetail.Columns.Add("Source", "Source")

            dgvErrorsDetail.Columns(GridErrors.Error_Desc).Width = 200
            dgvErrorsDetail.Columns(GridErrors.Source).Width = 600

            dgvErrorsDetail.Columns(GridErrors.Error_Desc).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvErrorsDetail.Columns(GridErrors.Source).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            dgvErrorsDetail.Columns(GridErrors.Error_Desc).ReadOnly = True
            dgvErrorsDetail.Columns(GridErrors.Source).ReadOnly = True

            dgvErrorsDetail.Columns(GridErrors.Error_Desc).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvErrorsDetail.Columns(GridErrors.Source).SortMode = DataGridViewColumnSortMode.NotSortable

            dgvErrorsDetail.DefaultCellStyle.WrapMode = DataGridViewTriState.True
            dgvErrorsDetail.RowTemplate.Height = 30
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtCustomerCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustomerCode.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Try
            Select Case KeyAscii
                Case System.Windows.Forms.Keys.Return
                    If CmdBrowse.Enabled = True Then CmdBrowse.Focus()
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

    Private Sub CmdCustCodeHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdCustCodeHelp.Click
        Dim strsql() As String
        Try
            strsql = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT CUSTOMER_CODE,CUST_NAME FROM VW_TKML_560B_PDS_GET_CUSTOMER WHERE UNIT_CODE ='" & gstrUnitId & "' ORDER BY CUSTOMER_CODE", "Customer(s) Help", 1)
            If Not (UBound(strsql) <= 0) Then
                If Not (UBound(strsql) = 0) Then
                    If (Len(strsql(0)) >= 1) And strsql(0) = "0" Then
                        MsgBox("No customer(s) record found.", MsgBoxStyle.Information, ResolveResString(100))
                        txtCustomerCode.Text = ""
                        lblCustCodeDes.Text = ""
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

    Private Sub txtCustomerCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyUp
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

    Private Sub txtCustomerCode_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
        Dim dtCustomer As New DataTable
        Try
            If txtCustomerCode.Text.Trim.Length = 0 Then
                lblCustCodeDes.Text = String.Empty
            Else
                dtCustomer = SqlConnectionclass.GetDataTable("SELECT CUSTOMER_CODE,CUST_NAME FROM VW_TKML_560B_PDS_GET_CUSTOMER WHERE UNIT_CODE = '" & gstrUnitId & "' AND CUSTOMER_CODE='" + txtCustomerCode.Text.Trim() + "'")
                If dtCustomer IsNot Nothing AndAlso dtCustomer.Rows.Count > 0 Then
                    txtCustomerCode.Text = Convert.ToString(dtCustomer.Rows(0)("CUSTOMER_CODE"))
                    lblCustCodeDes.Text = Convert.ToString(dtCustomer.Rows(0)("CUST_NAME"))
                Else
                    MessageBox.Show("Customer code doesn't exist.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    txtCustomerCode.Text = String.Empty
                    lblCustCodeDes.Text = String.Empty
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

    Private Sub CmdBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdBrowse.Click
        Dim oFDialog As New OpenFileDialog()
        Try
            If String.IsNullOrEmpty(txtCustomerCode.Text.Trim()) Then
                MessageBox.Show("Please select a customer.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtCustomerCode.Text = String.Empty
                lblCustCodeDes.Text = String.Empty
                txtCustomerCode.Focus()
                Exit Sub
            End If

            dgvErrorsDetail.DataSource = Nothing
            Dim fileExtension As String = String.Empty

            oFDialog.Filter = "Text files (*.txt)|*.txt"
            oFDialog.FilterIndex = 3
            oFDialog.RestoreDirectory = True
            oFDialog.Title = "Select a file to upload"
            If oFDialog.ShowDialog() = DialogResult.OK Then
                fileExtension = Path.GetExtension(oFDialog.FileName)
                If String.IsNullOrEmpty(fileExtension) Then
                    MessageBox.Show("Please select valid extension file.Valid file extension is : .TXT", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    lblFileName.Tag = String.Empty
                    lblFileName.Text = String.Empty
                    Exit Sub
                End If
                If fileExtension.ToUpper() <> ".TXT" Then
                    MessageBox.Show("Please select valid extension file.Valid file extension is : .TXT", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    lblFileName.Tag = String.Empty
                    lblFileName.Text = String.Empty
                    Exit Sub
                End If
                lblFileName.Tag = oFDialog.SafeFileName
                lblFileName.Text = oFDialog.FileName
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub btnUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        Try
            If String.IsNullOrEmpty(txtCustomerCode.Text.Trim()) Then
                MessageBox.Show("Please select a customer.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtCustomerCode.Text = String.Empty
                lblCustCodeDes.Text = String.Empty
                txtCustomerCode.Focus()
                Exit Sub
            End If

            If String.IsNullOrEmpty(lblFileName.Text.Trim()) Then
                MessageBox.Show("Please select a file to upload.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                lblFileName.Text = String.Empty
                lblFileName.Tag = String.Empty
                CmdBrowse.Focus()
                Exit Sub
            End If

            If MessageBox.Show("Are you sure to upload?", ResolveResString(100), MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

            dgvErrorsDetail.Rows.Clear()
            If ReadSequencePDSFile() = False Then
                Exit Sub
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Function ReadSequencePDSFile() As Boolean
        Dim currentRow As String()
        Dim dtSeqPDS As New DataTable
        Dim sqlCmd As New SqlCommand
        Try
            ReadSequencePDSFile = False
            Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(lblFileName.Text.Trim)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters("|")
                PictureBox1.Visible = True

                dtSeqPDS.Columns.Add("LOADID", GetType(System.Int32))
                dtSeqPDS.Columns.Add("SUPPLIER", GetType(System.String))
                dtSeqPDS.Columns.Add("SUPPLIER_PLANT", GetType(System.String))
                dtSeqPDS.Columns.Add("SUPPLIER_NAME", GetType(System.String))
                dtSeqPDS.Columns.Add("PLANT", GetType(System.String))
                dtSeqPDS.Columns.Add("PLANT_NAME", GetType(System.String))
                dtSeqPDS.Columns.Add("RECEIVING_PLACE", GetType(System.String))
                dtSeqPDS.Columns.Add("ORDER_TYPE", GetType(System.Int32))
                dtSeqPDS.Columns.Add("PDS_NUMBER", GetType(System.String))
                dtSeqPDS.Columns.Add("EKB_ORDER_NO", GetType(System.String))
                dtSeqPDS.Columns.Add("COLLECT_DATE", GetType(System.DateTime))
                dtSeqPDS.Columns.Add("COLLECT_TIME", GetType(System.String))
                dtSeqPDS.Columns.Add("ARRIVAL_DATE", GetType(System.DateTime))
                dtSeqPDS.Columns.Add("ARRIVAL_TIME", GetType(System.String))
                dtSeqPDS.Columns.Add("MAIN_ROUTE_GRP_CODE", GetType(System.String))
                dtSeqPDS.Columns.Add("MAIN_ROUTE_ORDER_SEQ", GetType(System.String))
                dtSeqPDS.Columns.Add("SUB_ROUTE_GRP_CODE", GetType(System.String))
                dtSeqPDS.Columns.Add("SUB_ROUTE_ORDER_SEQ", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS1_ROUTE", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS1_DOCK", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS1_ARV_DATE", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS1_ARV_TIME", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS1_DPT_DATE", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS1_DPT_TIME", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS2_ROUTE", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS2_DOCK", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS2_ARV_DATE", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS2_ARV_TIME", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS2_DPT_DATE", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS2_DPT_TIME", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS3_ROUTE", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS3_DOCK", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS3_ARV_DATE", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS3_ARV_TIME", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS3_DPT_DATE", GetType(System.String))
                dtSeqPDS.Columns.Add("CRS3_DPT_TIME", GetType(System.String))
                dtSeqPDS.Columns.Add("SUPPLIER_TYPE", GetType(System.String))
                dtSeqPDS.Columns.Add("LINE_NO", GetType(System.Int32))
                dtSeqPDS.Columns.Add("PART_NO", GetType(System.String))
                dtSeqPDS.Columns.Add("PART_NAME", GetType(System.String))
                dtSeqPDS.Columns.Add("KANBAN_NO", GetType(System.String))
                dtSeqPDS.Columns.Add("SEQ_NO", GetType(System.Int64))
                dtSeqPDS.Columns.Add("PACKING_SIZE", GetType(System.Int32))
                dtSeqPDS.Columns.Add("UNIT_QTY", GetType(System.Int32))
                dtSeqPDS.Columns.Add("PACK_QTY", GetType(System.Int32))
                dtSeqPDS.Columns.Add("ZERO_ORDER", GetType(System.String))
                dtSeqPDS.Columns.Add("SORT_LANE", GetType(System.String))
                dtSeqPDS.Columns.Add("SHIPPING_DATE", GetType(System.String))
                dtSeqPDS.Columns.Add("SHIPPING_TIME", GetType(System.String))
                dtSeqPDS.Columns.Add("KB_PRINT_DATE_P", GetType(System.String))
                dtSeqPDS.Columns.Add("KB_PRINT_TIME_P", GetType(System.String))
                dtSeqPDS.Columns.Add("KB_PRINT_DATE_L", GetType(System.String))
                dtSeqPDS.Columns.Add("KB_PRINT_TIME_L", GetType(System.String))
                dtSeqPDS.Columns.Add("REMARK", GetType(System.String))
                dtSeqPDS.Columns.Add("ORDER_RELEASE_DATE", GetType(System.String))
                dtSeqPDS.Columns.Add("ORDER_RELEASE_TIME", GetType(System.String))
                dtSeqPDS.Columns.Add("MAIN_ROUTE_DATE", GetType(System.String))
                dtSeqPDS.Columns.Add("BILL_OUT_FLAG", GetType(System.String))
                dtSeqPDS.Columns.Add("SHIPPING_DOCK", GetType(System.String))
                dtSeqPDS.Columns.Add("PACKING_TYPE", GetType(System.String))
                dtSeqPDS.Columns.Add("KANBAN_ORIENTATION", GetType(System.String))
                dtSeqPDS.Columns.Add("DOLLY_CODE", GetType(System.String))
                dtSeqPDS.Columns.Add("PICKER_ROUTE", GetType(System.String))
                dtSeqPDS.Columns.Add("CYCLE_SERIAL", GetType(System.String))
                dtSeqPDS.Columns.Add("PACKING_CODE", GetType(System.String))
                dtSeqPDS.Columns.Add("TKM_LINE_ADDRESS", GetType(System.String))
                dtSeqPDS.Columns.Add("BPA_NUMBER", GetType(System.String))
                dtSeqPDS.Columns.Add("SECONDARY_VENDOR_CODE", GetType(System.String))
                dtSeqPDS.Columns.Add("SECONDARY_PARTCODE", GetType(System.String))

                Dim drSeqPDS As DataRow
                While Not MyReader.EndOfData
                    Try
                        Application.DoEvents()
                        currentRow = MyReader.ReadFields()
                        If currentRow.Length < 68 Then
                            PictureBox1.Visible = False
                            MessageBox.Show("Invalid file format !" + vbCr + "No. of columns in a file can't be less then 68 !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Exit Function
                        End If

                        drSeqPDS = dtSeqPDS.NewRow()
                        drSeqPDS("LOADID") = Val(currentRow(0))
                        drSeqPDS("SUPPLIER") = currentRow(1)
                        drSeqPDS("SUPPLIER_PLANT") = currentRow(2)
                        drSeqPDS("SUPPLIER_NAME") = currentRow(3)
                        drSeqPDS("PLANT") = currentRow(4)
                        drSeqPDS("PLANT_NAME") = currentRow(5)
                        drSeqPDS("RECEIVING_PLACE") = currentRow(6)
                        drSeqPDS("ORDER_TYPE") = Val(currentRow(7))
                        drSeqPDS("PDS_NUMBER") = currentRow(8)
                        drSeqPDS("EKB_ORDER_NO") = currentRow(9)
                        drSeqPDS("COLLECT_DATE") = Convert.ToDateTime(currentRow(10)).ToString("dd MMM yyyy")
                        drSeqPDS("COLLECT_TIME") = currentRow(11)
                        drSeqPDS("ARRIVAL_DATE") = Convert.ToDateTime(currentRow(12)).ToString("dd MMM yyyy")
                        drSeqPDS("ARRIVAL_TIME") = currentRow(13)
                        drSeqPDS("MAIN_ROUTE_GRP_CODE") = currentRow(14)
                        drSeqPDS("MAIN_ROUTE_ORDER_SEQ") = currentRow(15)
                        drSeqPDS("SUB_ROUTE_GRP_CODE") = currentRow(16)
                        drSeqPDS("SUB_ROUTE_ORDER_SEQ") = currentRow(17)
                        drSeqPDS("CRS1_ROUTE") = currentRow(18)
                        drSeqPDS("CRS1_DOCK") = currentRow(19)
                        drSeqPDS("CRS1_ARV_DATE") = currentRow(20)
                        drSeqPDS("CRS1_ARV_TIME") = currentRow(21)
                        drSeqPDS("CRS1_DPT_DATE") = currentRow(22)
                        drSeqPDS("CRS1_DPT_TIME") = currentRow(23)
                        drSeqPDS("CRS2_ROUTE") = currentRow(24)
                        drSeqPDS("CRS2_DOCK") = currentRow(25)
                        drSeqPDS("CRS2_ARV_DATE") = currentRow(26)
                        drSeqPDS("CRS2_ARV_TIME") = currentRow(27)
                        drSeqPDS("CRS2_DPT_DATE") = currentRow(28)
                        drSeqPDS("CRS2_DPT_TIME") = currentRow(29)
                        drSeqPDS("CRS3_ROUTE") = currentRow(30)
                        drSeqPDS("CRS3_DOCK") = currentRow(31)
                        drSeqPDS("CRS3_ARV_DATE") = currentRow(32)
                        drSeqPDS("CRS3_ARV_TIME") = currentRow(33)
                        drSeqPDS("CRS3_DPT_DATE") = currentRow(34)
                        drSeqPDS("CRS3_DPT_TIME") = currentRow(35)
                        drSeqPDS("SUPPLIER_TYPE") = currentRow(36)
                        drSeqPDS("LINE_NO") = Val(currentRow(37))
                        drSeqPDS("PART_NO") = currentRow(38)
                        drSeqPDS("PART_NAME") = currentRow(39)
                        drSeqPDS("KANBAN_NO") = currentRow(40)
                        drSeqPDS("SEQ_NO") = Val(currentRow(41))
                        drSeqPDS("PACKING_SIZE") = Val(currentRow(42))
                        drSeqPDS("UNIT_QTY") = Val(currentRow(43))
                        drSeqPDS("PACK_QTY") = Val(currentRow(44))
                        drSeqPDS("ZERO_ORDER") = currentRow(45)
                        drSeqPDS("SORT_LANE") = currentRow(46)
                        drSeqPDS("SHIPPING_DATE") = currentRow(47)
                        drSeqPDS("SHIPPING_TIME") = currentRow(48)
                        drSeqPDS("KB_PRINT_DATE_P") = currentRow(49)
                        drSeqPDS("KB_PRINT_TIME_P") = currentRow(50)
                        drSeqPDS("KB_PRINT_DATE_L") = currentRow(51)
                        drSeqPDS("KB_PRINT_TIME_L") = currentRow(52)
                        drSeqPDS("REMARK") = currentRow(53)
                        drSeqPDS("ORDER_RELEASE_DATE") = currentRow(54)
                        drSeqPDS("ORDER_RELEASE_TIME") = ""
                        drSeqPDS("MAIN_ROUTE_DATE") = currentRow(55)
                        drSeqPDS("BILL_OUT_FLAG") = currentRow(56)
                        drSeqPDS("SHIPPING_DOCK") = currentRow(57)
                        drSeqPDS("PACKING_TYPE") = currentRow(58)
                        drSeqPDS("KANBAN_ORIENTATION") = currentRow(59)
                        drSeqPDS("DOLLY_CODE") = currentRow(60)
                        drSeqPDS("PICKER_ROUTE") = currentRow(61)
                        drSeqPDS("CYCLE_SERIAL") = currentRow(62)
                        drSeqPDS("PACKING_CODE") = currentRow(63)
                        drSeqPDS("TKM_LINE_ADDRESS") = currentRow(64)
                        drSeqPDS("BPA_NUMBER") = currentRow(65)
                        drSeqPDS("SECONDARY_VENDOR_CODE") = currentRow(66)
                        drSeqPDS("SECONDARY_PARTCODE") = currentRow(67)

                        dtSeqPDS.Rows.Add(drSeqPDS)
                    Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                        MessageBox.Show("Line " & ex.Message & "is not valid and will be skipped.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Exit Function
                    End Try
                End While

                If dtSeqPDS IsNot Nothing AndAlso dtSeqPDS.Rows.Count > 0 Then
                    With sqlCmd
                        .CommandType = CommandType.StoredProcedure
                        .CommandTimeout = 600 ' 10 Minute
                        If chkSequencePDS.Checked Then
                            .CommandText = "USP_UPLOAD_PDS_TKML_SEQUENCING_SCHEDULE_DOOR_TRIM"
                        Else
                            .CommandText = "USP_UPLOAD_PDS_TKML_SEQUENCING_SCHEDULE"
                        End If
                        .Parameters.Clear()
                        .Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
                        .Parameters.AddWithValue("@CUSTOMER_CODE", txtCustomerCode.Text.Trim())
                        .Parameters.AddWithValue("@FILE_TYPE", "OEM TKML SEQ")
                        .Parameters.AddWithValue("@FILE_PATH", lblFileName.Text.Trim)
                        .Parameters.AddWithValue("@USER_ID", mP_User)
                        .Parameters.AddWithValue("@IP_ADDRESS", gstrIpaddressWinSck)
                        .Parameters.AddWithValue("@UDT_PDS_TKML_SEQUENCING_FILE", dtSeqPDS)
                        .Parameters.AddWithValue("@ACTION", "UPLOAD_MANUALLY")
                        .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                        SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                        If Convert.ToString(.Parameters("@MESSAGE").Value) <> "" Then
                            PopulateScheduleLog()
                            PictureBox1.Visible = False
                            MessageBox.Show(Convert.ToString(.Parameters("@MESSAGE").Value), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Function
                        Else
                            ReadSequencePDSFile = True
                            PictureBox1.Visible = False
                            lblFileName.Text = String.Empty
                            lblFileName.Tag = String.Empty
                            MessageBox.Show("File Uploaded Successfully !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    End With
                End If
            End Using
        Catch ex As Exception
            RaiseException(ex)
        Finally
            PictureBox1.Visible = False
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            If dtSeqPDS IsNot Nothing Then
                dtSeqPDS.Dispose()
            End If
            If sqlCmd IsNot Nothing Then
                sqlCmd.Dispose()
            End If
        End Try
    End Function

    Private Function PopulateScheduleLog() As Boolean
        Dim dtErrors As New DataTable
        PopulateScheduleLog = False
        Try
            Dim i As Integer = 0
            dtErrors = SqlConnectionclass.GetDataTable("SELECT ERR_DESC,SOURCE FROM SCHEDULE_TKML_UPLOAD_LOG WITH (NOLOCK) WHERE UNIT_CODE = '" & gstrUnitId & "' AND  IP_ADDRESS='" & gstrIpaddressWinSck & "' ORDER BY LOG_ID ")
            If dtErrors IsNot Nothing AndAlso dtErrors.Rows.Count > 0 Then
                dgvErrorsDetail.Rows.Clear()
                dgvErrorsDetail.Rows.Add(dtErrors.Rows.Count)
                For Each dr As DataRow In dtErrors.Rows
                    dgvErrorsDetail.Rows(i).Cells(GridErrors.Error_Desc).Value = Convert.ToString(dr("ERR_DESC"))
                    dgvErrorsDetail.Rows(i).Cells(GridErrors.Source).Value = Convert.ToString(dr("SOURCE"))
                    i += 1
                Next
                PopulateScheduleLog = True
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
    End Function

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            txtCustomerCode.Text = String.Empty
            lblCustCodeDes.Text = String.Empty
            lblFileName.Tag = String.Empty
            lblFileName.Text = String.Empty
            dgvErrorsDetail.Rows.Clear()
            PictureBox1.Visible = False
            txtCustomerCode.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtCustomerCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustomerCode.TextChanged
        Try
            If String.IsNullOrEmpty(txtCustomerCode.Text) Then
                lblCustCodeDes.Text = String.Empty
                lblFileName.Tag = String.Empty
                lblFileName.Text = String.Empty
                dgvErrorsDetail.Rows.Clear()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub chkSequencePDS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSequencePDS.CheckedChanged
        Try
            txtCustomerCode.Text = String.Empty
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
End Class