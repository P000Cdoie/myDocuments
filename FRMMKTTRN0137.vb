Imports System.IO
Imports System.Data.SqlClient

'*********************************************************************************************************************
'Copyright(c)       - MIND
'Name of Module     - Invoice Against Acknowlegment
'Name of Form       - FRMMKTTRN0126  , TKML 560B PDS Uploading & Picklist Generation
'Created by         - Prashant Rajpal
'Created Date       - 24 Dec 2024
'description        - To upload file received from NISSAN and Marelli and Create Invoice 
'*********************************************************************************************************************

Public Class FRMMKTTRN0137
    Private Enum GridDetails
        ShipCode = 0
        Ship_to_code
        Supplier
        Site
        SchDate
        Partnumber
        Deliverynumber
        PoReceivedNo
        Quantity
        UnitPrice
        Amount
        SOPrice
    End Enum
    Dim mblnSaveStatus As Boolean
    Dim minttotalnoofrows As Integer

    Private Sub FRMMKTTRN0126_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            txtConsCode.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FRMMKTTRN0126_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Call FitToClient(Me, GrpMain, ctlHeader, GrpBoxButtons, 500)
            Me.MdiParent = mdifrmMain
            PictureBox1.Visible = False
            REFRESHFORM()
            'ConfigureGridColumn()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub ConfigureGridColumn()
        'Try
        '    dgvDetail.Columns.Clear()
        '    dgvDetail.RowHeadersVisible = True
        '    dgvDetail.Columns.Add("ShipCode", "Ship Code")
        '    dgvDetail.Columns.Add("ShiptoCode", "ShiptoCode")
        '    dgvDetail.Columns.Add("Site", "Site")
        '    dgvDetail.Columns.Add("Supplier", "Supplier")
        '    dgvDetail.Columns.Add("HorizonEndDate", "SchDate")
        '    dgvDetail.Columns.Add("Partnumber", "Partnumber")
        '    dgvDetail.Columns.Add("Deliverynumber", "Deliverynumber")
        '    dgvDetail.Columns.Add("PoReceivedNo", "PoReceivedNo")
        '    dgvDetail.Columns.Add("Quantity", "Quantity")
        '    dgvDetail.Columns.Add("UnitPrice", "UnitPrice")
        '    dgvDetail.Columns.Add("Amount", "Amount")
        '    dgvDetail.Columns.Add("SOPrice", "SOPrice")

        '    dgvDetail.Columns(GridDetails.ShipCode).Width = 100
        '    dgvDetail.Columns(GridDetails.Ship_to_code).Width = 100
        '    dgvDetail.Columns(GridDetails.Site).Width = 100
        '    dgvDetail.Columns(GridDetails.Supplier).Width = 100
        '    dgvDetail.Columns(GridDetails.SchDate).Width = 100
        '    dgvDetail.Columns(GridDetails.Partnumber).Width = 100
        '    dgvDetail.Columns(GridDetails.Deliverynumber).Width = 100
        '    dgvDetail.Columns(GridDetails.PoReceivedNo).Width = 100
        '    dgvDetail.Columns(GridDetails.Quantity).Width = 100
        '    dgvDetail.Columns(GridDetails.UnitPrice).Width = 100
        '    dgvDetail.Columns(GridDetails.Amount).Width = 100
        '    dgvDetail.Columns(GridDetails.soprice).Width = 100


        '    dgvDetail.Columns(GridDetails.ShipCode).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        '    dgvDetail.Columns(GridDetails.Ship_to_code).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        '    dgvDetail.Columns(GridDetails.Site).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        '    dgvDetail.Columns(GridDetails.Supplier).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        '    dgvDetail.Columns(GridDetails.SchDate).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        '    dgvDetail.Columns(GridDetails.Partnumber).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        '    dgvDetail.Columns(GridDetails.Deliverynumber).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        '    dgvDetail.Columns(GridDetails.PoReceivedNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        '    dgvDetail.Columns(GridDetails.Quantity).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        '    dgvDetail.Columns(GridDetails.UnitPrice).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        '    dgvDetail.Columns(GridDetails.Amount).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        '    dgvDetail.Columns(GridDetails.SOPrice).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        '    dgvDetail.Columns(GridDetails.ShipCode).ReadOnly = True
        '    dgvDetail.Columns(GridDetails.Ship_to_code).ReadOnly = True
        '    dgvDetail.Columns(GridDetails.Site).ReadOnly = True
        '    dgvDetail.Columns(GridDetails.Supplier).ReadOnly = True
        '    dgvDetail.Columns(GridDetails.SchDate).ReadOnly = True
        '    dgvDetail.Columns(GridDetails.Partnumber).ReadOnly = True
        '    dgvDetail.Columns(GridDetails.Deliverynumber).ReadOnly = True
        '    dgvDetail.Columns(GridDetails.PoReceivedNo).ReadOnly = True
        '    dgvDetail.Columns(GridDetails.Quantity).ReadOnly = True
        '    dgvDetail.Columns(GridDetails.UnitPrice).ReadOnly = True
        '    dgvDetail.Columns(GridDetails.Amount).ReadOnly = True
        '    dgvDetail.Columns(GridDetails.SOPrice).ReadOnly = True


        '    dgvDetail.Columns(GridDetails.ShipCode).SortMode = DataGridViewColumnSortMode.NotSortable
        '    dgvDetail.Columns(GridDetails.Ship_to_code).SortMode = DataGridViewColumnSortMode.NotSortable
        '    dgvDetail.Columns(GridDetails.Site).SortMode = DataGridViewColumnSortMode.NotSortable
        '    dgvDetail.Columns(GridDetails.Supplier).SortMode = DataGridViewColumnSortMode.NotSortable
        '    dgvDetail.Columns(GridDetails.SchDate).SortMode = DataGridViewColumnSortMode.NotSortable
        '    dgvDetail.Columns(GridDetails.Partnumber).SortMode = DataGridViewColumnSortMode.NotSortable
        '    dgvDetail.Columns(GridDetails.Deliverynumber).SortMode = DataGridViewColumnSortMode.NotSortable
        '    dgvDetail.Columns(GridDetails.PoReceivedNo).SortMode = DataGridViewColumnSortMode.NotSortable
        '    dgvDetail.Columns(GridDetails.Quantity).SortMode = DataGridViewColumnSortMode.NotSortable
        '    dgvDetail.Columns(GridDetails.UnitPrice).SortMode = DataGridViewColumnSortMode.NotSortable
        '    dgvDetail.Columns(GridDetails.Amount).SortMode = DataGridViewColumnSortMode.NotSortable
        '    dgvDetail.Columns(GridDetails.SOPrice).SortMode = DataGridViewColumnSortMode.NotSortable


        '    dgvDetail.DefaultCellStyle.WrapMode = DataGridViewTriState.True

        '    dgvDetail.RowTemplate.Height = 20
        'Catch ex As Exception
        '    RaiseException(ex)

        'End Try
    End Sub

    Private Sub txtConsCodeCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtConsCode.KeyPress
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

    Private Sub CmdConsCodeHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdConsCodeHelp.Click
        Dim strsql() As String
        Try
            strsql = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT CUSTOMER_CODE,CUST_NAME FROM CUSTOMER_MST WHERE UNIT_CODE ='" & gstrUNITID & "' AND INVOICEAGSTACKN=1 ORDER BY CUSTOMER_CODE", "Customer(s) Help", 1)
            If Not (UBound(strsql) <= 0) Then
                If Not (UBound(strsql) = 0) Then
                    If (Len(strsql(0)) >= 1) And strsql(0) = "0" Then
                        MsgBox("No customer(s) record found.", MsgBoxStyle.Information, ResolveResString(100))
                        txtConsCode.Text = ""
                        lblConsCodeDes.Text = ""
                        Exit Sub
                    Else
                        txtConsCode.Text = strsql(0).Trim
                        lblConsCodeDes.Text = strsql(1).Trim
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub txtConsCodeCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtConsCode.KeyUp
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000
        Try
            If KeyCode = System.Windows.Forms.Keys.F1 Then
                If CmdConsCodeHelp.Enabled Then Call CmdCustCodeHelp_Click(CmdConsCodeHelp, New System.EventArgs())
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub txtConsCodeCode_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtConsCode.Validating
        Dim dtCustomer As New DataTable
        Try
            If txtConsCode.Text.Trim.Length = 0 Then
                lblConsCodeDes.Text = String.Empty
            Else
                dtCustomer = SqlConnectionclass.GetDataTable("SELECT CUSTOMER_CODE,CUST_NAME FROM CUSTOMER_MST WHERE UNIT_CODE = '" & gstrUNITID & "' AND CUSTOMER_CODE='" + txtConsCode.Text.Trim() + "' AND INVOICEAGSTACKN=1 ")
                If dtCustomer IsNot Nothing AndAlso dtCustomer.Rows.Count > 0 Then
                    txtConsCode.Text = Convert.ToString(dtCustomer.Rows(0)("CUSTOMER_CODE"))
                    lblConsCodeDes.Text = Convert.ToString(dtCustomer.Rows(0)("CUST_NAME"))
                Else
                    MessageBox.Show("Customer code doesn't exist.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    txtConsCode.Text = String.Empty
                    lblConsCodeDes.Text = String.Empty
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
            If String.IsNullOrEmpty(txtConsCode.Text.Trim()) Then
                MessageBox.Show("Please select a customer.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtConsCode.Text = String.Empty
                lblConsCodeDes.Text = String.Empty
                txtConsCode.Focus()
                Exit Sub
            End If

            dgvDetail.DataSource = Nothing

            Dim fileExtension As String = String.Empty

            oFDialog.Filter = "Text files (*.csv)|*.csv"
            oFDialog.FilterIndex = 3
            oFDialog.RestoreDirectory = True
            oFDialog.Title = "Select a file to upload"
            If oFDialog.ShowDialog() = DialogResult.OK Then
                fileExtension = Path.GetExtension(oFDialog.FileName)
                If String.IsNullOrEmpty(fileExtension) Then
                    MessageBox.Show("Please select valid extension file.Valid file extension is : .CSV", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    lblFileName.Tag = String.Empty
                    lblFileName.Text = String.Empty
                    Exit Sub
                End If
                If fileExtension.ToUpper() <> ".CSV" Then
                    MessageBox.Show("Please select valid extension file.Valid file extension is : .CSV", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    lblFileName.Tag = String.Empty
                    lblFileName.Text = String.Empty
                    Exit Sub
                End If
                lblFileName.Tag = oFDialog.SafeFileName
                lblFileName.Text = oFDialog.FileName

            End If


            If String.IsNullOrEmpty(lblFileName.Text.Trim()) Then
                MessageBox.Show("Please select a file to upload.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                lblFileName.Text = String.Empty
                lblFileName.Tag = String.Empty
                CmdBrowse.Focus()
                Exit Sub
            End If

            'If MessageBox.Show("Are you sure to upload?", ResolveResString(100), MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = Windows.Forms.DialogResult.No Then
            'Exit Sub
            'End If

            dgvDetail.DataSource = Nothing
            If ReadInvoiceFile() = False Then
                Exit Sub
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

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If String.IsNullOrEmpty(txtCustCode.Text.Trim()) Then
                MessageBox.Show("Please select a Customer.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtCustCode.Text = String.Empty
                lblcustCodeDes.Text = String.Empty
                txtCustCode.Focus()
                Exit Sub
            End If


            If String.IsNullOrEmpty(txtConsCode.Text.Trim()) Then
                MessageBox.Show("Please select a Consignee .", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtConsCode.Text = String.Empty
                lblConsCodeDes.Text = String.Empty
                txtConsCode.Focus()
                Exit Sub
            End If

            If String.IsNullOrEmpty(lblFileName.Text.Trim()) Then
                MessageBox.Show("Please select a file to upload.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                lblFileName.Text = String.Empty
                lblFileName.Tag = String.Empty
                CmdBrowse.Focus()
                Exit Sub
            End If

            If mblnSaveStatus = False Then
                MessageBox.Show("Part No(s)(RED COLOUR) are not defined in Sales Order , please Define First !!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            If MessageBox.Show("Are you sure to Save the Invoice ?", ResolveResString(100), MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = Windows.Forms.DialogResult.No Then
                Exit Sub
            Else
                If validations() = True Then
                    savedata()
                End If
            End If

            'dgvDetail.Rows.Clear()

            
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Function validations() As Boolean

        Dim strStockLoc As String = String.Empty
        Dim STRQUERY As String = String.Empty
        Dim strPartnumber As String = String.Empty
        Dim strExist As String = String.Empty
        Dim ValidatePartcode As Object
        Dim INTQUANTITY As Integer
        Dim INTSORATE As Integer
        Dim INTUNITPRICE As Integer


        Try
            strStockLoc = SqlConnectionclass.ExecuteScalar("SELECT STOCK_LOCATION FROM SALECONF (NOLOCK) WHERE UNIT_CODE='" + gstrUNITID + "' AND  INVOICE_TYPE ='INV' AND SUB_TYPE ='F' AND LOCATION_CODE ='" & gstrUNITID & "' AND (FIN_START_DATE <= GETDATE() AND FIN_END_DATE >= GETDATE())")
            If strStockLoc.Length = 0 Then
                MessageBox.Show("Please Define Stock Location in Sales Conf First !!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            'For intRow As Integer = 0 To minttotalnoofrows - 1

            '    strPartnumber = dgvDetail.Rows(intRow).Cells(GridDetails.Partnumber).Value

            '    STRQUERY = "SELECT TOP 1 ITEM_CODE FROM VW_ACTIVE_SALESORDER  WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" & txtCustCode.Text.Trim & "' AND CONSIGNEE_CODE='" & txtConsCode.Text & "' AND CUST_DRGNO = '" & strPartnumber & "'"
            '    strExist = SqlConnectionclass.ExecuteScalar(STRQUERY)

            '    If strExist.Length = 0 Then
            '        MessageBox.Show("Sales order is not defined this part code :" + strPartnumber + " !!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
            '        Return False
            '    End If
            'Next

            'For intCount As Integer = 0 To minttotalnoofrows - 1

            '    With dgvDetail
            '        ValidatePartcode = dgvDetail.Rows(intCount).Cells(GridDetails.Partnumber).Value
            '        INTQUANTITY = dgvDetail.Rows(intCount).Cells(GridDetails.Quantity).Value
            '        INTUNITPRICE = dgvDetail.Rows(intCount).Cells(GridDetails.UnitPrice).Value
            '        INTSORATE = dgvDetail.Rows(intCount).Cells(GridDetails.SOPrice).Value


            '        If INTQUANTITY = 0 Then
            '            MessageBox.Show("QUANTITY CAN NOT BE ZERO FOR PARTCODE- " & ValidatePartcode, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
            '            Exit Function
            '        End If

            '        If INTUNITPRICE <> INTSORATE Then
            '            MessageBox.Show("UNIT PRICE SHOULD MATCH WITH SO PRICE - " & ValidatePartcode, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)

            '            Exit Function
            '        End If
            '    End With

            'Next
            Return True
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Function
    Private Function ReadInvoiceFile() As Boolean

        Dim currentRow As String()
        Dim dtSeqPDS As New DataTable
        Dim sqlCmd As New SqlCommand
        Dim strSql As String
        Dim dtcolor As New DataTable
        Dim stritemrate As String


        Dim i As Integer = 0
        mblnSaveStatus = True
        Try
            ReadInvoiceFile = False
            Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(lblFileName.Text.Trim)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(",")
                PictureBox1.Visible = True

                dtSeqPDS.Columns.Add("SIDECODE", GetType(System.String))
                dtSeqPDS.Columns.Add("SHIPTOCODE", GetType(System.String))
                dtSeqPDS.Columns.Add("SUPPLIER", GetType(System.String))
                dtSeqPDS.Columns.Add("SITE", GetType(System.String))
                dtSeqPDS.Columns.Add("ARRIVAL_DATE", GetType(System.String))
                dtSeqPDS.Columns.Add("PART_NUMBER", GetType(System.String))
                dtSeqPDS.Columns.Add("DELIVERYNUMBER", GetType(System.String))
                dtSeqPDS.Columns.Add("PO_RECEIVEDDNO", GetType(System.String))
                dtSeqPDS.Columns.Add("QUANTITY", GetType(System.Int32))
                dtSeqPDS.Columns.Add("UNIT_PRICE", GetType(System.Double))
                dtSeqPDS.Columns.Add("AMOUNT", GetType(System.String))
                dtSeqPDS.Columns.Add("SOPRICE", GetType(System.Double))

                Dim drSeqPDS As DataRow
                While Not MyReader.EndOfData
                    Try
                        Application.DoEvents()
                        currentRow = MyReader.ReadFields()
                        If currentRow.Length <> 11 Then
                            PictureBox1.Visible = False
                            MessageBox.Show("Invalid file format !" + vbCr + "No. of columns in a file can't be less or more than 11 !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Exit Function
                        End If

                        drSeqPDS = dtSeqPDS.NewRow()

                        drSeqPDS("SIDECODE") = currentRow(0)
                        drSeqPDS("SHIPTOCODE") = currentRow(1)
                        drSeqPDS("SUPPLIER") = currentRow(2)
                        drSeqPDS("SITE") = currentRow(3)
                        drSeqPDS("ARRIVAL_DATE") = currentRow(4)
                        drSeqPDS("PART_NUMBER") = currentRow(5)
                        drSeqPDS("DELIVERYNUMBER") = currentRow(6)
                        drSeqPDS("PO_RECEIVEDDNO") = currentRow(7)
                        drSeqPDS("QUANTITY") = Val(currentRow(8))
                        drSeqPDS("UNIT_PRICE") = currentRow(9)
                        drSeqPDS("AMOUNT") = currentRow(10)
                        stritemrate = SqlConnectionclass.ExecuteScalar("SELECT RATE FROM VW_GET_SORATE WHERE UNIT_CODE='" + gstrUNITID + "' AND  ACCOUNT_CODE='" & txtCustCode.Text.Trim & "' AND CONSIGNEE_CODE ='" & txtConsCode.Text & "' AND CUST_DRGNO='" & drSeqPDS("PART_NUMBER") & "'")
                        If Val(stritemrate) = 0 Then
                            stritemrate = 0
                        Else
                            drSeqPDS("SOPRICE") = stritemrate
                        End If

                        dtSeqPDS.Rows.Add(drSeqPDS)



                    Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                        MessageBox.Show("Line " & ex.Message & "is not valid and will be skipped.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Exit Function
                    End Try
                End While
                If currentRow(1) <> txtConsCode.Text Then
                    MsgBox("File is not associated with this Consignee Code ", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Function
                End If
                minttotalnoofrows = dtSeqPDS.Rows.Count

                If dtSeqPDS IsNot Nothing AndAlso dtSeqPDS.Rows.Count > 0 Then
                    'dgvDetail.DataSource = Nothing
                    dgvDetail.DataSource = dtSeqPDS
                    dgvDetail.ReadOnly = True
                    dgvDetail.AllowUserToAddRows = False
                    dgvDetail.AllowUserToDeleteRows = False
                    dgvDetail.DefaultCellStyle.BackColor = Color.LightGray

                    If dgvDetail.Columns.Contains("Serialno") = False Then
                        Dim serialcol As New DataGridViewTextBoxColumn()
                        serialcol.HeaderText = "S. No"
                        serialcol.Name = "SerialNo"
                        dgvDetail.Columns.Insert(0, serialcol)
                    End If

                    For i = 0 To dtSeqPDS.Rows.Count - 1
                        dgvDetail.Rows(i).Cells("SerialNo").Value = i + 1

                        dtcolor = SqlConnectionclass.GetDataTable("SELECT TOP 1 1 FROM VW_ACTIVE_SALESORDER WHERE UNIT_CODE='" & gstrUNITID & "' AND CONSIGNEE_CODE='" & txtConsCode.Text.Trim & "' AND CUST_dRGNO='" & dtSeqPDS.Rows(i)("PART_NUMBER").ToString & "'")
                        If dtcolor IsNot Nothing AndAlso dtcolor.Rows.Count > 0 Then
                            dgvDetail.Rows(i).DefaultCellStyle.BackColor = Color.LightBlue
                        Else
                            dgvDetail.Rows(i).DefaultCellStyle.BackColor = Color.Red
                            mblnSaveStatus = False
                        End If

                    Next

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

    Private Function DELETEDATA() As Boolean
        Dim sqlcommand As New SqlCommand()

        Try
            With SqlCommand
                .CommandText = "USP_DELETE_ACKNOWLEDGMENT_INVOICE"
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .Connection = SqlConnectionclass.GetConnection()
                .Parameters.Clear()
                .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                .Parameters.AddWithValue("@CUSTOMER_CODE", txtCustCode.Text)
                .Parameters.AddWithValue("@CONSIGNEE_CODE", txtConsCode.Text)
                .Parameters.AddWithValue("@INVOICENO", txtchallanno.Text)
                .Parameters.Add("@ERRMSG", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                .ExecuteNonQuery()
                '.Dispose()

                If .Parameters("@ERRMSG").Value.ToString() <> "" Then
                    MsgBox(.Parameters("@ERRMSG").Value.ToString(), MsgBoxStyle.Information, ResolveResString(100))
                    Return False
                End If

            End With

        Catch ex As Exception
            RaiseException(ex)
        Finally
            SqlCommand.Dispose()

        End Try
    End Function
    Private Function SaveData() As Boolean
        Dim DT As New DataTable()
        Dim sqlcommand As New SqlCommand()
        Dim intLoopCounter As Integer
        Dim VAL_Shipcode As Object = Nothing, VAL_ShiptoCode As Object = Nothing, VAL_Site As Object = Nothing, VAL_Supplier As Object = Nothing, VAL_Schdate As Object = Nothing
        Dim VAL_Partnumber As Object = Nothing, VAL_Deliverynumber As Object = Nothing, VAL_PoReceivedNo As Object = Nothing, VAL_Quantity As Object = Nothing, VAL_UnitPrice As Object = Nothing, VAL_Amount As Object = Nothing
        Dim DocumentNo As String = String.Empty

        Try
            SaveData = False

            'DT.Columns.Add("ShipCode", GetType(String))
            'DT.Columns.Add("ShiptoCode", GetType(String))
            'DT.Columns.Add("Site", GetType(String))
            'DT.Columns.Add("Supplier", GetType(String))
            'DT.Columns.Add("SchDate", GetType(String))
            'DT.Columns.Add("Partnumber", GetType(String))
            'DT.Columns.Add("Deliverynumber", GetType(String))
            'DT.Columns.Add("PoReceivedNo", GetType(String))
            'DT.Columns.Add("Quantity", GetType(Double))
            'DT.Columns.Add("UnitPrice", GetType(Double))
            'DT.Columns.Add("Amount", GetType(Double))
            'DT.Columns.Add("SOPrice", GetType(Double))

            DT = dgvDetail.DataSource

            Dim validationqtyrows() As DataRow = DT.Select("quantity= 0")

            For Each row As DataRow In validationqtyrows
                MessageBox.Show("QUANTITY CAN NOT BE ZERO FOR PARTCODE", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Function
            Next
            Dim validationrows() As DataRow = DT.Select("UNIT_PRICE - SOPRICE >0 or SOPRICE - UNIT_PRICE >0")

            For Each row As DataRow In validationrows
                MessageBox.Show("UNIT PRICE SHOULD MATCH WITH SO PRICE FOR THIS PARTCODE : " + row("part_number"), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                For i As Integer = 0 To dgvDetail.Rows.Count
                    If dgvDetail.Rows(i).Cells.Item("DELIVERYNUMBER").Value.ToString.ToUpper = row("DELIVERYNUMBER").ToString.ToUpper Then
                        dgvDetail.Rows(i).Selected = True
                        Exit Function
                    End If
                Next

                'If validationrows.Length > 0 Then
                '        Dim index As Integer = dgvDetail.Rows.IndexOf(row(0))
                '        dgvDetail.Rows(index).Selected = True

                '    End If
                Exit Function

                Next

                With sqlcommand
                .CommandText = "USP_GENERATE_INVOICE_AGAINST_ACKNOWLEDGMENT"
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .Connection = SqlConnectionclass.GetConnection()
                .Parameters.Clear()
                .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                .Parameters.AddWithValue("@CUSTOMER_CODE", txtCustCode.Text)
                .Parameters.AddWithValue("@CONSIGNEE_CODE", txtConsCode.Text)
                .Parameters.AddWithValue("@FILENAME", lblFileName.Text)
                .Parameters.AddWithValue("@FILEDATA", DT)
                .Parameters.AddWithValue("User_Id", mP_User)
                .Parameters.Add("@ERRMSG", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                .Parameters.Add("@CHALLAN_No", SqlDbType.VarChar, 10).Direction = ParameterDirection.Output


                .ExecuteNonQuery()
                '.Dispose()

                If .Parameters("@ERRMSG").Value.ToString() <> "" Then
                    MsgBox(.Parameters("@ERRMSG").Value.ToString(), MsgBoxStyle.Information, ResolveResString(100))
                    Return False
                End If
                DocumentNo = .Parameters("@CHALLAN_No").Value.ToString()
            End With

            MsgBox("Invoice Generated Successfully with Document No :" + DocumentNo, MsgBoxStyle.Information, ResolveResString(100))
            dgvDetail.DataSource = Nothing

            txtchallanno.Text = String.Empty
            txtCustCode.Text = String.Empty
            txtConsCode.Text = String.Empty


            Return True

        Catch ex As Exception
            RaiseException(ex)
        Finally

            sqlcommand.Dispose()
            DT.Dispose()
        End Try
    End Function

    Private Function PopulateScheduleLog() As Boolean
        Dim dtErrors As New DataTable
        PopulateScheduleLog = False
        Try
            Dim i As Integer = 0
            dtErrors = SqlConnectionclass.GetDataTable("SELECT ERR_DESC,SOURCE FROM SCHEDULE_TKML_UPLOAD_LOG WITH (NOLOCK) WHERE UNIT_CODE = '" & gstrUNITID & "' AND  IP_ADDRESS='" & gstrIpaddressWinSck & "' ORDER BY LOG_ID ")
            If dtErrors IsNot Nothing AndAlso dtErrors.Rows.Count > 0 Then
                dgvDetail.DataSource = Nothing
                dgvDetail.Rows.Add(dtErrors.Rows.Count)
                For Each dr As DataRow In dtErrors.Rows
                    dgvDetail.Rows(i).Cells(GridDetails.ShipCode).Value = Convert.ToString(dr("ERR_DESC"))
                    dgvDetail.Rows(i).Cells(GridDetails.Ship_to_code).Value = Convert.ToString(dr("SOURCE"))
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

    Private Function PopulateInvoiceData() As Boolean
        Dim dtinvoicedata As New DataTable

        PopulateInvoiceData = False
        Try


            dtinvoicedata = SqlConnectionclass.GetDataTable("SELECT ITEMCODE,PARTNUMBER,DELIVERYNUMBER,QUANTITY,RATE ,SHIPCODE, SHIPTOCODE , HORIZONENDDATE , SUPPLIER  FROM VW_DELIVERY_MKT_ACKN_HISTORY WITH (NOLOCK) WHERE UNIT_CODE = '" & gstrUNITID & "' AND  DOC_NO='" & txtchallanno.Text & "'")
            If dtinvoicedata IsNot Nothing AndAlso dtinvoicedata.Rows.Count > 0 Then
                dgvDetail.DataSource = dtinvoicedata
                PopulateInvoiceData = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            If (dtinvoicedata IsNot Nothing) Then
                dtinvoicedata.Dispose()
                dtinvoicedata = Nothing
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Function


    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then

                REFRESHFORM()
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub REFRESHFORM()
        Try
            txtCustCode.Text = String.Empty
            lblcustCodeDes.Text = String.Empty
            txtConsCode.Text = String.Empty
            lblConsCodeDes.Text = String.Empty
            lblFileName.Tag = String.Empty
            lblFileName.Text = String.Empty
            dgvDetail.DataSource = Nothing
            PictureBox1.Visible = False
            txtchallanno.Text = String.Empty
            txtCustCode.Focus()
            btnDelete.Enabled = False : btnDelete.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            btnSave.Enabled = True : btnSave.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)


        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtConsCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtConsCode.TextChanged
        Try
            If String.IsNullOrEmpty(txtConsCode.Text) Then
                lblConsCodeDes.Text = String.Empty
                lblFileName.Tag = String.Empty
                lblFileName.Text = String.Empty
                dgvDetail.DataSource = Nothing
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub chkSequencePDS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            txtConsCode.Text = String.Empty
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub CmdCustCodeHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdCustCodeHelp.Click
        Dim strsql() As String
        Try
            strsql = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT CUSTOMER_CODE,CUST_NAME ,BILL_ADDRESS1 FROM CUSTOMER_MST WHERE UNIT_CODE ='" & gstrUNITID & "' AND INVOICEAGSTACKN=1 ORDER BY CUSTOMER_CODE", "Customer(s) Help", 1)
            If Not (UBound(strsql) <= 0) Then
                If Not (UBound(strsql) = 0) Then
                    If (Len(strsql(0)) >= 1) And strsql(0) = "0" Then
                        MsgBox("No customer(s) record found.", MsgBoxStyle.Information, ResolveResString(100))
                        txtCustCode.Text = ""
                        lblcustCodeDes.Text = ""
                        Exit Sub
                    Else
                        txtCustCode.Text = strsql(0).Trim
                        lblcustCodeDes.Text = strsql(1).Trim + strsql(2).Trim

                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub txtCustCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustCode.TextChanged
        Try

            lblcustCodeDes.Text = String.Empty
            lblFileName.Tag = String.Empty
            lblFileName.Text = String.Empty
            txtConsCode.Text = String.Empty
            dgvDetail.DataSource = Nothing

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub cmddocno_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdChallanNo.Click
        Dim strQselect As String = String.Empty
        Dim INV_detail As String = String.Empty

        strQselect = " SELECT DISTINCT  A.DOC_NO AS INVOICE_NO,A.INVOICE_DATE AS INV_DATE "
        strQselect = strQselect + " ,A.ACCOUNT_CODE AS CUST_CODE ,A.CUST_NAME ,A.CONSIGNEE_CODE AS CONS_CODE "
        strQselect = strQselect + " FROM SALESCHALLAN_DTL(NOLOCK) AS A"
        strQselect = strQselect + " INNER JOIN SALES_DTL (NOLOCK) AS B"
        strQselect = strQselect + " ON A.UNIT_CODE=B.UNIT_CODE AND A.DOC_NO=B.DOC_NO "
        strQselect = strQselect + " INNER JOIN CUSTOMER_MST (NOLOCK) AS CM On CM.UNIT_CODE=A.UNIT_CODE AND A.ACCOUNT_CODE=CM.CUSTOMER_CODE "
        strQselect = strQselect + " WHERE A.UNIT_CODE='" & gstrUNITID & "' AND INVOICEAGSTACKN = 1 ORDER BY A.INVOICE_DATE DESC"

        INV_detail = GetDocumentNo(strQselect, "INV_No")


        If INV_detail = Nothing Then
            MessageBox.Show("NO INVOICE EXISTS", "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            Exit Sub
        End If

        Dim Split() As String = INV_detail.Split("~")
        If Not String.IsNullOrEmpty(Split(0)) Then
            txtchallanno.Text = Split(0).ToString

            If Not String.IsNullOrEmpty(Split(2)) Then
                txtCustCode.Text = Split(2).ToString
            End If
            If Not String.IsNullOrEmpty(Split(4)) Then
                txtConsCode.Text = Split(4).ToString
            End If

            PopulateInvoiceData()
            btnSave.Enabled = False : btnSave.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            btnDelete.Enabled = True : btnDelete.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        Else
            txtchallanno.Text = String.Empty
            Call ConfirmWindow(10225, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
        End If


    End Sub
    Private Function GetDocumentNo(ByVal Qselect As String, ByVal HelpFor As String) As String

        Dim Result As String = ""
        Dim strHelp() As String = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, Qselect)

        Try
            If UBound(strHelp) = -1 Then Exit Function
            If Not IsNothing(strHelp) AndAlso strHelp.Length = 0 Then
                Result = "No record found."
                Return Result
            ElseIf String.IsNullOrEmpty(strHelp(1)) Then
                Result = "No record found."
                Return Result
            Else
                If HelpFor = "INV_No" Then
                    If IsNothing(strHelp(0)) Or IsNothing(strHelp(1)) Or IsNothing(strHelp(2)) Or IsNothing(strHelp(3)) Or IsNothing(strHelp(4)) Then
                        Result = "No record found."
                        Return Result
                    End If
                    Result = strHelp(0).ToString + "~" + strHelp(1).ToString + "~" + strHelp(2).ToString + "~" + strHelp(3).ToString + "~" + strHelp(4).ToString '+ "~" + strHelp(5).ToString + "~" + strHelp(6).ToString
                End If
                Return Result
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Function

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim strSQL As String = String.Empty
        Try
            If txtchallanno.Text = "" Then
                MessageBox.Show("Please Select Challan No. first , Cannot Delete!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            Else


                strSQL = "SELECT TOP 1 1 FROM SALESCHALLAN_DTL WHERE UNIT_CODE ='" & gstrUNITID & "' AND DOC_NO = '" & txtchallanno.Text & "' AND BILL_FLAG = 1"
                If IsRecordExists(strSQL) Then
                    MessageBox.Show("Invoice Locked. Cannot Delete!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                strSQL = "Select top 1 1 from SALESCHALLAN_DTL where unit_code ='" & gstrUNITID & "' and doc_no = '" & txtchallanno.Text & "' and Cancel_flag = 1"
                If IsRecordExists(strSQL) Then
                    MessageBox.Show("This is a cancelled Invoice!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
                If MessageBox.Show("Are you sure, you want to delete invoice !!.", ResolveResString(100), MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                    Call DELETEDATA()
                    MessageBox.Show("Invoice Deleted Successfully ", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    REFRESHFORM()
                End If
            End If


        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Public Function IsRecordExists(ByVal strSql As String) As Boolean
        'satvir
        Dim oSqlDr As SqlDataReader = Nothing

        IsRecordExists = False

        Try
            oSqlDr = SqlConnectionclass.ExecuteReader(strSql)
            If oSqlDr.HasRows = True Then
                IsRecordExists = True
            End If
            If oSqlDr.IsClosed = False Then oSqlDr.Close()
        Catch ex As Exception
            RaiseException(ex)
        Finally
            If Not oSqlDr Is Nothing AndAlso oSqlDr.IsClosed = False Then
                oSqlDr.Close()
            End If
            oSqlDr = Nothing
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Function

End Class