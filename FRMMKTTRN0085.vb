Imports System
Imports System.Data.SqlClient
Imports VB = Microsoft.VisualBasic
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel

Public Class FRMMKTTRN0085
    '========================================================================================
    'COPYRIGHT          :   MOTHERSONSUMI INFOTECH & DESIGN LTD.
    'AUTHOR             :   GEETANJALI AGGARWAL
    'CREATION DATE      :   18 Nov 2014
    'DESCRIPTION        :   10688280-Sales Provisioning Authorisation 
    '========================================================================================

#Region "Global variables"
    Dim dtSelItems As System.Data.DataTable
    Dim gPDocNo_Authorized As Boolean = False
    Dim strFinDocNo As String
    Dim objCrDr As prj_DrCrNote.cls_DrCrNote
    Dim mobjExchangeRate As prj_ExchRateGetter.cls_ExchRateGetter
    'Dim blnEOUFlag As Boolean
    Dim blnECSSTax As Boolean
    Dim intECSRoundOffDecimal As Integer
    Dim intExciseRoundOffDecimal As Integer
    Dim dblBasicValue As Decimal = 0.0
    Dim strSupplInvNo As String = String.Empty
    Dim xbook As Workbook
    Dim xSheet As Worksheet
    Dim xApp As Application
    Dim DocFrm As Form
    Dim part_code As String = String.Empty
    Dim rate As Double = 0.0

    Enum RateWiseDtlGrid
        Col_Select = 1
        Col_Part_Code = 2
        Col_Model_Code = 3
        Col_Inv_Qty = 4
        Col_ActInvQty = 5
        Col_InvRate = 6
        Col_PriceChange = 7
        Col_NewRate = 8
        Col_ChangeEff = 9
        Col_Change = 10
        Col_NewInvRate = 11
        Col_TotEffVal = 12
        Col_ReasonChange = 13
        Col_CorrectionNature = 14
        Col_ShowInv = 15
    End Enum
#End Region

#Region "Form Events"
    Private Sub FRMMKTTRM0085_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Call FitToClient(Me, GrpMain, ctlFormHeader, GrpCmdBtn, 610)
            Me.MdiParent = mdifrmMain
            InitializeForm(1)
            Me.BringToFront()
            'If Not IsRecordExists("select 1 from Sales_Prov_Hdr where Unit_Code='" + gstrUNITID + "' and PERSONTO_AUTH in (select employee_code from User_Mst where User_id='" + mP_User + "' and Unit_Code='" + gstrUNITID + "')") Then
            '    MessageBox.Show("You have no Provisioning to authorize.", ResolveResString(1000), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            '    Me.Close()
            '    Return
            'End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
#End Region

#Region "Routines & Functions"
    '''<summary>Initialize form</summary>
    '''<param name="Form_Status_flag"> Form Initialization No </param>
    Private Sub InitializeForm(ByVal Form_Status_flag As Integer)
        Try
            If Form_Status_flag = 1 Then   'Page Load
                strFinDocNo = String.Empty
                strSupplInvNo = String.Empty
                dtSelItems = Nothing
                txtProvDocNo.Text = ""
                txtProvDocNo.Enabled = True
                dtpFrm.Enabled = False
                dtpToDt.Enabled = False
                lblKAMName.Text = ""
                txtCustCode.Text = String.Empty
                lblCustDesc.Text = String.Empty
                txtCustCode.Enabled = False
                txtRModelDesc.Text = String.Empty
                txtRPartDesc.Text = String.Empty
                txtPositiveVal.Text = String.Empty
                rbText.Enabled = False
                rbPartCode.Enabled = False
                txtMultilinetxt.Text = String.Empty
                txtMultilinetxt.Enabled = False
                txtPartCode.Text = String.Empty
                txtPartDesc.Text = String.Empty

                BtnHelpPartCode.Enabled = False
                BtnHelpProvDocNo.Enabled = True
                BtnAuthorize.Enabled = False
                BtnReject.Enabled = False
                BtnUploadDoc.Enabled = False
                DocUploadPanel.Visible = False
                AddColumnDispatchDtlGrid()
                AddRateWiseGridColumn()
                ClearAllTaxes()
                ClearDataGridView(dgvDispatchDtl)
                ClearDataGridView(dgvDocList)
            ElseIf Form_Status_flag = 2 Then    'Sales Prov View Mode for submiited Doc No
                dgvDispatchDtl.ReadOnly = True
                txtProvDocNo.Enabled = True
                dtpFrm.Enabled = False
                dtpToDt.Enabled = False
                txtCustCode.Enabled = False
                txtRModelDesc.Text = String.Empty
                txtRPartDesc.Text = String.Empty
                BtnHelpProvDocNo.Enabled = True
                LockUnlockRateWiseGrid(1, fpSpreadRateWiseDtl.MaxRows, True)
                rbText.Enabled = Not gPDocNo_Authorized
                rbPartCode.Enabled = Not gPDocNo_Authorized
                BtnAuthorize.Enabled = Not gPDocNo_Authorized
                BtnReject.Enabled = Not gPDocNo_Authorized

                bindDocList()
                BtnUploadDoc.Enabled = True
                If Not gPDocNo_Authorized Then
                    If rbText.Checked Then
                        txtMultilinetxt.Enabled = True
                    Else
                        BtnHelpPartCode.Enabled = True
                    End If
                Else
                    txtMultilinetxt.Enabled = False
                    BtnHelpPartCode.Enabled = False
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    '''<summary>To clear BOM exploded items from FAR grid</summary>
    Private Function clearFarGrid() As Boolean
        Dim flag As Boolean = True
        Try
            fpSpreadRateWiseDtl.MaxRows = 0
        Catch ex As Exception
            flag = False
        End Try
        Return flag
    End Function

    ''' <summary>
    ''' Save data in database.
    ''' </summary>
    ''' <returns>True for successful save.</returns>
    Private Function SaveData(Optional ByVal AuthStatus As String = "A") As Boolean
        Dim SqlCmd As New SqlCommand
        Try
            If AuthStatus = "A" Then
                If Not ValidateSave() Then
                    Return False
                End If
                If MessageBox.Show("Are You Sure , you want to authorise?", ResolveResString(100), MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.No Then
                    Return False
                End If
            Else
                If MessageBox.Show("Are You Sure , you want to reject?", ResolveResString(100), MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = System.Windows.Forms.DialogResult.No Then
                    Return False
                End If
            End If

            SqlCmd = New SqlCommand
            With SqlCmd
                .CommandType = CommandType.StoredProcedure
                .CommandText = "USP_SALES_PROV_Generation"
                .Connection = SqlConnectionclass.GetConnection()
                .Transaction = .Connection.BeginTransaction
                .CommandTimeout = 0
                If AuthStatus = "A" Then
                    If Val(txtNegativeVal.Text.Trim()) <> 0 Then
                        objCrDr = New prj_DrCrNote.cls_DrCrNote(GetServerDate.ToString("dd MMM yyyy"))
                        mobjExchangeRate = New prj_ExchRateGetter.cls_ExchRateGetter(gstrUNITID)
                        If Not generateCreditNote() Then
                            .Transaction.Rollback()
                            Return False
                        End If
                    End If
                    If Val(txtPositiveVal.Text.Trim()) > 0 Then
                        Dim mstrPurposeCode As String = String.Empty
                        If IsRecordExists(" SELECT inv_GLD_prpsCode FROM SaleConf WHERE Unit_code='" + gstrUNITID + "' and Invoice_Type='Inv' AND Sub_Type ='F'" & _
                                          " AND Location_Code='" + gstrUNITID + "' and datediff(dd,getdate(),fin_start_date)<=0 and datediff(dd,fin_end_date,getdate())<=0 ") Then
                            mstrPurposeCode = SqlConnectionclass.ExecuteScalar(" SELECT inv_GLD_prpsCode " & _
                                              " FROM SaleConf WHERE Unit_code='" + gstrUNITID + "' and Invoice_Type='Inv' AND Sub_Type ='F'" & _
                                              " AND Location_Code='" + gstrUNITID + "' and datediff(dd,getdate(),fin_start_date)<=0  " & _
                                              " and datediff(dd,fin_end_date,getdate())<=0 ")
                            If mstrPurposeCode = "" Then
                                .Transaction.Rollback()
                                MessageBox.Show("Please select a Purpose Code in Sales Configuration", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                                Return False
                            End If
                        Else
                            MessageBox.Show("No record found in Sales Configuration for the selected Location, Invoice Type and Sub-Category", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            .Transaction.Rollback()
                            Return False
                        End If
                        SelectChallanNoFromSupplementatryInvHdr()
                        If String.IsNullOrEmpty(strSupplInvNo) Then
                            .Transaction.Rollback()
                            MessageBox.Show("Supplemantory Invoice No cannot be generated.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            Return False
                        End If
                    End If
                End If
                .Parameters.Clear()
                .Parameters.Add(New SqlParameter("@p_ProvDocNo", SqlDbType.VarChar, 20, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                .Parameters("@p_ProvDocNo").Value = txtProvDocNo.Text.Trim
                .Parameters.AddWithValue("@p_TRANTYPE", "C")
                If rbText.Checked Then
                    .Parameters.AddWithValue("@p_IsText", 1)
                    .Parameters.AddWithValue("@p_TextPartDesc", txtMultilinetxt.Text.Trim())
                Else
                    .Parameters.AddWithValue("@p_IsText", 0)
                    .Parameters.AddWithValue("@p_TextPartDesc", txtPartCode.Text.Trim())
                End If
                .Parameters.AddWithValue("@p_AuthStatus", AuthStatus)
                .Parameters.AddWithValue("@p_CRDocNo", strFinDocNo)
                .Parameters.AddWithValue("@p_SupplInvDocNo", strSupplInvNo)
                .Parameters.AddWithValue("@p_UserId", mP_User)
                .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                .Parameters.AddWithValue("@p_IPAddress", gstrIpaddressWinSck)
                .Parameters.Add(New SqlParameter("@p_ERROR", SqlDbType.VarChar, 200, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                .ExecuteNonQuery()
                If String.IsNullOrEmpty(Convert.ToString(SqlCmd.Parameters("@p_ERROR").Value)) Then
                    .Transaction.Commit()
                    Dim msg As String = String.Empty
                    If AuthStatus = "A" Then
                        If Not String.IsNullOrEmpty(strFinDocNo) Then
                            msg = "Credit Note " + strFinDocNo + ","
                        End If
                        If Not String.IsNullOrEmpty(strSupplInvNo) Then
                            msg = msg + "Supplimentory Invoice No " + strSupplInvNo
                        End If
                        If Not String.IsNullOrEmpty(msg) Then
                            msg = msg + " generated successfully."
                        End If
                    ElseIf AuthStatus = "R" Then
                        msg = "Provision Doc No " + txtProvDocNo.Text.Trim() + " rejected successfully."
                    End If
                    If Not String.IsNullOrEmpty(msg) Then
                        MessageBox.Show(msg, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Return True
                    End If
                Else
                    .Transaction.Rollback()
                    MessageBox.Show(Convert.ToString(SqlCmd.Parameters("@p_ERROR").Value), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If
            End With
        Catch ex As Exception
            If (Not IsNothing(SqlCmd.Transaction)) Then
                SqlCmd.Transaction.Rollback()
            End If
            RaiseException(ex)
            Return False
        Finally
            SqlCmd.Dispose()
        End Try
        Return True
    End Function

    ''' <summary>Validate manadatory fields</summary>
    Private Function ValidateSave() As Boolean
        Dim strQry As String = String.Empty
        Dim isValid As Boolean = False
        Try
            If rbText.Checked And String.IsNullOrEmpty(txtMultilinetxt.Text.Trim()) Then
                MessageBox.Show("Matter to be printed on invoice can not be blank.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtMultilinetxt.Focus()
                Return False
            End If

            If rbPartCode.Checked And String.IsNullOrEmpty(txtPartCode.Text.Trim()) Then
                MessageBox.Show("Part Code can not be blank.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                txtPartCode.Focus()
                Return False
            End If
            Return True
        Catch ex As Exception
            RaiseException(ex)
        End Try
        Return True
    End Function

    ''' <summary>
    ''' Clear records from datagridview
    ''' </summary>
    ''' <param name="dgv">DatagridView which needs to clear.</param>
    Private Sub ClearDataGridView(ByRef dgv As DataGridView)
        Try
            dgv.DataSource = Nothing
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    '''<summary>Add Column in far spread grid.</summary>
    Private Sub AddRateWiseGridColumn()
        Try
            With fpSpreadRateWiseDtl
                .EditEnterAction = FPSpreadADO.EditEnterActionConstants.EditEnterActionNext
                .BackColorStyle = FPSpreadADO.BackColorStyleConstants.BackColorStyleUnderGrid

                .MaxRows = 0
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .set_RowHeight(.Row, 20)

                .MaxCols = RateWiseDtlGrid.Col_Select
                .Col = RateWiseDtlGrid.Col_Select
                .Value = " "
                .set_ColWidth(RateWiseDtlGrid.Col_Select, 2)

                .MaxCols = RateWiseDtlGrid.Col_Part_Code
                .Col = RateWiseDtlGrid.Col_Part_Code
                .Value = "Part Code"
                .set_ColWidth(RateWiseDtlGrid.Col_Part_Code, 15)

                .MaxCols = RateWiseDtlGrid.Col_Model_Code
                .Col = RateWiseDtlGrid.Col_Model_Code
                .Value = "Model Code"
                .set_ColWidth(RateWiseDtlGrid.Col_Model_Code, 10)

                .MaxCols = RateWiseDtlGrid.Col_Inv_Qty
                .Col = RateWiseDtlGrid.Col_Inv_Qty
                .Value = "Invoice Qty"
                .set_ColWidth(RateWiseDtlGrid.Col_Inv_Qty, 10)

                .MaxCols = RateWiseDtlGrid.Col_ActInvQty
                .Col = RateWiseDtlGrid.Col_ActInvQty
                .Value = "Actual Invoice Qty" + vbCrLf + "[A]"
                .set_ColWidth(RateWiseDtlGrid.Col_ActInvQty, 11)

                .MaxCols = RateWiseDtlGrid.Col_InvRate
                .Col = RateWiseDtlGrid.Col_InvRate
                .Value = "Invoice Rate" + vbCrLf + "[B]"
                .set_ColWidth(RateWiseDtlGrid.Col_InvRate, 10)

                .MaxCols = RateWiseDtlGrid.Col_PriceChange
                .Col = RateWiseDtlGrid.Col_PriceChange
                .Value = "Price Change"
                .set_ColWidth(RateWiseDtlGrid.Col_PriceChange, 8)

                .MaxCols = RateWiseDtlGrid.Col_NewRate
                .Col = RateWiseDtlGrid.Col_NewRate
                .Value = "New Rate" + vbCrLf + "[C](Press F1)"
                .set_ColWidth(RateWiseDtlGrid.Col_NewInvRate, 15)

                .MaxCols = RateWiseDtlGrid.Col_ChangeEff
                .Col = RateWiseDtlGrid.Col_ChangeEff
                .Value = "Change Effect(%)"
                .set_ColWidth(RateWiseDtlGrid.Col_ChangeEff, 6)

                .MaxCols = RateWiseDtlGrid.Col_Change
                .Col = RateWiseDtlGrid.Col_Change
                .Value = "Change" + vbCrLf + "(%)"
                .set_ColWidth(RateWiseDtlGrid.Col_Change, 6)

                .MaxCols = RateWiseDtlGrid.Col_NewInvRate
                .Col = RateWiseDtlGrid.Col_NewInvRate
                .Value = "New Invoice Rate" + vbCrLf + "[C]"
                .set_ColWidth(RateWiseDtlGrid.Col_NewInvRate, 13)

                .MaxCols = RateWiseDtlGrid.Col_TotEffVal
                .Col = RateWiseDtlGrid.Col_TotEffVal
                .Value = "Total Effect in Value" + vbCrLf + "[D]=[A]*([C]-[B])"
                .set_ColWidth(RateWiseDtlGrid.Col_TotEffVal, 15)

                .MaxCols = RateWiseDtlGrid.Col_ReasonChange
                .Col = RateWiseDtlGrid.Col_ReasonChange
                .Value = "Reason for Change"
                .set_ColWidth(RateWiseDtlGrid.Col_ReasonChange, 10)

                .MaxCols = RateWiseDtlGrid.Col_CorrectionNature
                .Col = RateWiseDtlGrid.Col_CorrectionNature
                .Value = "Nature of Correction" + vbCrLf + "(press F1)"
                .set_ColWidth(RateWiseDtlGrid.Col_CorrectionNature, 15)

                .MaxCols = RateWiseDtlGrid.Col_ShowInv
                .Col = RateWiseDtlGrid.Col_ShowInv
                .Value = "Show Invoice(s)"
                .set_ColWidth(RateWiseDtlGrid.Col_ShowInv, 10)

                .ClipboardOptions = FPSpreadADO.ClipboardOptionsConstants.ClipboardOptionsNoHeaders
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ''' <summary>Bind data in FAR grid.</summary>
    Private Sub bindRateItemfpGrid()
        Dim strQry As String = String.Empty
        Dim odt As New System.Data.DataTable
        Dim strUOM As String = String.Empty
        Try
            odt = GetRateWisePartDetail()
            If Not IsNothing(odt) Then
                With fpSpreadRateWiseDtl
                    .MaxRows = 0
                    For Each row As DataRow In odt.Rows
                        .MaxRows = .MaxRows + 1
                        .Row = .MaxRows
                        .set_RowHeight(.Row, 12)

                        .Col = RateWiseDtlGrid.Col_Part_Code
                        .Value = Convert.ToString(row("ITEM_CODE"))

                        .Col = RateWiseDtlGrid.Col_Model_Code
                        .Value = Convert.ToString(row("VarModel"))

                        .Col = RateWiseDtlGrid.Col_Inv_Qty
                        .Value = Convert.ToString(row("Invoice_Qty"))
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatMax = 9999999999.99
                        .TypeFloatMin = 0
                        .TypeFloatSeparator = False
                        .TypeFloatDecimalPlaces = 2

                        .Col = RateWiseDtlGrid.Col_ActInvQty
                        .Value = Convert.ToString(row("ActualInvoiceQty"))
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatMax = 9999999999.99
                        .TypeFloatMin = 0
                        .TypeFloatSeparator = False
                        .TypeFloatDecimalPlaces = 2

                        .Col = RateWiseDtlGrid.Col_InvRate
                        .Value = Convert.ToString(row("Rate"))
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatMax = 9999999999.99
                        .TypeFloatMin = 0
                        .TypeFloatSeparator = False
                        .TypeFloatDecimalPlaces = 2

                        .Col = RateWiseDtlGrid.Col_PriceChange
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
                        .TypeComboBoxList = "Value" + Chr(9) + "Percentage(%)"
                        If Convert.ToBoolean(row("PriceChangeType")) Then     '0 for Value and 1 for %
                            .TypeComboBoxCurSel = 1
                        Else
                            .TypeComboBoxCurSel = 0
                        End If

                        .Col = RateWiseDtlGrid.Col_ChangeEff
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
                        .TypeComboBoxList = "-" + Chr(9) + "+"
                        If Convert.ToBoolean(row("ChangeEffect")) Then   '0 for Subtraction and 1 for Addition;
                            .TypeComboBoxCurSel = 1
                        Else
                            .TypeComboBoxCurSel = 0
                        End If

                        'If Convert.ToBoolean(row("SELECT")) Then
                        .Col = RateWiseDtlGrid.Col_NewRate
                        If Not Convert.ToBoolean(row("PriceChangeType")) Then     '0 for Value and 1 for %
                            .Value = Convert.ToString(row("NewRate"))
                        End If
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatMax = 9999999999.99
                        .TypeFloatMin = 0
                        .TypeFloatSeparator = False
                        .TypeFloatDecimalPlaces = 2

                        .Col = RateWiseDtlGrid.Col_Change
                        .Value = Convert.ToString(row("PercentageChange"))
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatMax = 999.99
                        .TypeFloatMin = 0
                        .TypeFloatSeparator = False
                        .TypeFloatDecimalPlaces = 2

                        CalculateRateEffect(.Row)
                        'End If

                        .Col = RateWiseDtlGrid.Col_ReasonChange
                        .Value = Convert.ToString(row("ChangeReason"))
                        .TypeMaxEditLen = 450

                        .Col = RateWiseDtlGrid.Col_CorrectionNature
                        .Value = Convert.ToString(row("CORRECTIONDESC"))
                        .CellTag = Convert.ToString(row("CORRECTIONID"))

                        .Col = RateWiseDtlGrid.Col_ShowInv
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                        .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                        .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                        .TypeButtonText = "Show Invoices"

                        .Col = RateWiseDtlGrid.Col_Select
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                        .Value = row("SELECT")
                        LockUnlockRateWiseGrid(.Row, .Row, True)  'lock grid
                    Next
                    'GetTotalBasicValEffect()
                    'CalculateTaxes()
                End With
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ''' <summary>Column addition in Dispatch Detail DataGridView</summary>
    Private Sub AddColumnDispatchDtlGrid()
        Try
            If dgvDispatchDtl.Columns.Count = 0 Then
                'Dim dgc As New DataGridViewCheckBoxColumn
                'dgc.DataPropertyName = "Select"
                'dgc.Name = "Select"
                'dgc.Width = 25
                'dgc.ReadOnly = False
                'dgc.HeaderText = ""
                'dgc.SortMode = DataGridViewColumnSortMode.Automatic
                'dgvDispatchDtl.Columns.Add(dgc)

                Dim dgvC As New DataGridViewTextBoxColumn
                dgvC.DataPropertyName = "ITEM_CODE"
                dgvC.Name = "Part_Code"
                dgvC.HeaderText = "Part Code"
                dgvC.Width = 85
                dgvC.ReadOnly = True
                dgvDispatchDtl.Columns.Add(dgvC)

                Dim dgvID As New DataGridViewTextBoxColumn
                dgvID.DataPropertyName = "ITEM_DESC"
                dgvID.Name = "ITEM_DESC"
                dgvID.HeaderText = "Part Description"
                dgvID.Width = 200
                dgvID.ReadOnly = True
                dgvDispatchDtl.Columns.Add(dgvID)

                Dim dgvD As New DataGridViewTextBoxColumn
                dgvD.DataPropertyName = "VarModel"
                dgvD.Name = "VarModel"
                dgvD.HeaderText = "Model Code"
                dgvD.Width = 50
                dgvD.ReadOnly = True
                dgvDispatchDtl.Columns.Add(dgvD)

                Dim dgvMD As New DataGridViewTextBoxColumn
                dgvMD.DataPropertyName = "ModelDesc"
                dgvMD.Name = "ModelDesc"
                dgvMD.HeaderText = "Model Description"
                dgvMD.Width = 100
                dgvMD.ReadOnly = True
                dgvDispatchDtl.Columns.Add(dgvMD)

                Dim dgvQ As New DataGridViewTextBoxColumn
                dgvQ.DataPropertyName = "Invoice_Qty"
                dgvQ.Name = "Invoice_Qty"
                dgvQ.HeaderText = "Invoice Quantity"
                dgvQ.Width = 70
                dgvQ.ReadOnly = True
                dgvQ.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgvDispatchDtl.Columns.Add(dgvQ)

                Dim dgvN As New DataGridViewTextBoxColumn
                dgvN.DataPropertyName = "ActualInvoiceQty"
                dgvN.Name = "ActualInvoiceQty"
                dgvN.HeaderText = "Actual Invoice Quantity"
                dgvN.Width = 90
                dgvN.ReadOnly = True
                dgvN.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgvDispatchDtl.Columns.Add(dgvN)

                dgvDispatchDtl.AutoGenerateColumns = False
                dgvDispatchDtl.RowsDefaultCellStyle.BackColor = Color.Lavender
                dgvDispatchDtl.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(190, 200, 255)
                dgvDispatchDtl.ColumnHeadersHeight = 35
                dgvDispatchDtl.AllowUserToResizeRows = False
                dgvDispatchDtl.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Bind Dispatch Detail Grid
    ''' </summary>
    Private Sub bindDispatchGrid()
        Dim strQuery As String = String.Empty
        Try
            ClearTmpTable()
            FillTmpTable()
            strQuery = " SELECT CASE ISNULL(SRQ.ITEM_CODE,'') WHEN '' THEN CAST(0 AS BIT) ELSE CAST(1 AS BIT) END [SELECT],TMP.CUSTOMER_CODE, TMP.ITEM_CODE, TMP.VARMODEL" & _
                      " , cast(SUM(TMP.INVOICE_QTY) as numeric(18,2)) INVOICE_QTY, cast(SUM(TMP.ACTUALINVOICEQTY) as numeric(18,2)) ACTUALINVOICEQTY" & _
                      " , ISNULL(IM.DESCRIPTION,'') [ITEM_DESC], ISNULL(BM.MODEL_DESC,'') MODELDESC" & _
                      " FROM SALES_PROV_TMPPARTDETAIL_AUTH TMP" & _
                      " INNER JOIN(" & _
                      "  SELECT DISTINCT ITEM_CODE, VARMODEL, UNIT_CODE FROM SALES_PROV_RATEWISEDETAIL WHERE CAST(PROV_DOCNO AS VARCHAR(20))='" + txtProvDocNo.Text.Trim() + "' AND UNIT_CODE='" + gstrUNITID + "' " & _
                      " ) SRQ " & _
                      " ON TMP.ITEM_CODE=SRQ.ITEM_CODE AND TMP.VARMODEL=SRQ.VARMODEL AND TMP.UNIT_CODE=SRQ.UNIT_CODE " & _
                      " LEFT JOIN ITEM_MST IM" & _
                      " ON TMP.ITEM_CODE=IM.ITEM_CODE AND TMP.UNIT_CODE=IM.UNIT_CODE" & _
                      " LEFT JOIN BUDGET_MODEL_MST BM" & _
                      " ON TMP.VARMODEL=BM.MODEL_CODE AND TMP.UNIT_CODE=BM.UNIT_CODE AND BM.ACTIVE=1" & _
                      " WHERE TMP.IPADDRESS='" + gstrIpaddressWinSck + "' AND TMP.UNIT_CODE='" + gstrUNITID + "'" & _
                      " GROUP BY TMP.CUSTOMER_CODE, TMP.ITEM_CODE, TMP.VARMODEL,CASE ISNULL(SRQ.ITEM_CODE,'') WHEN '' THEN CAST(0 AS BIT) ELSE CAST(1 AS BIT) END" & _
                      " , ISNULL(IM.DESCRIPTION,''), ISNULL(BM.MODEL_DESC,'')"

            dtSelItems = New System.Data.DataTable
            dtSelItems = SqlConnectionclass.GetDataTable(strQuery)
            If (Not IsNothing(dtSelItems)) Then
                dgvDispatchDtl.DataSource = dtSelItems
                If (dtSelItems.Rows.Count = 0) Then
                    MessageBox.Show("No Record found.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ''' <summary>Fill data in Tmp Table.</summary>
    Private Sub FillTmpTable()
        Dim strQry As String = String.Empty
        Try
            strQry = " DELETE FROM SALES_PROV_TMPPARTDETAIL_AUTH WHERE UNIT_CODE='" + gstrUNITID + "' AND IPADDRESS='" + gstrIpaddressWinSck + "';" & _
                   " INSERT INTO Sales_Prov_tmpPartDetail_Auth( CUSTOMER_CODE, ITEM_CODE, VARMODEL, INVOICE_NO, INVOICEDATE, CUSTREFNO, CUSTREFDATE, AMENDNO, AMENDDATE" & _
                   " , RATE, INVOICE_QTY, ACTUALINVOICEQTY , PRICECHANGETYPE, NEWRATE, CHANGEEFFECT, PERCENTAGECHANGE, CHANGEREASON" & _
                   " , REJECTIONQTY, CDRQTY, NATUREOFCORRECTION, IPADDRESS, UNIT_CODE) " & _
                   " SELECT CUSTOMER_CODE, ITEM_CODE, VARMODEL, INVOICE_NO, INVOICEDATE, CUSTREFNO, CUSTREFDATE, AMENDNO, AMENDDATE" & _
                   " , RATE, INVOICE_QTY, ACTUALINVOICEQTY, PRICECHANGETYPE, NEWRATE, CHANGEEFFECT, PERCENTAGECHANGE, CHANGEREASON" & _
                   " , REJECTIONQTY, CDRQTY, NATUREOFCORRECTION, '" + gstrIpaddressWinSck + "', UNIT_CODE" & _
                   " FROM SALES_PROV_PARTINVOICEDETAIL" & _
                   " WHERE CAST(PROV_DOCNO as varchar(20))='" + txtProvDocNo.Text.Trim() + "' AND UNIT_CODE='" + gstrUNITID + "'"
            SqlConnectionclass.ExecuteNonQuery(strQry)
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ''' <summary>Clear data in Tmp Table.</summary>
    Private Sub ClearTmpTable()
        Dim strQry As String = String.Empty
        Try
            strQry = "DELETE FROM SALES_PROV_TMPPARTDETAIL_AUTH WHERE UNIT_CODE='" + gstrUNITID + "' AND IPADDRESS='" + gstrIpaddressWinSck + "';"
            SqlConnectionclass.ExecuteNonQuery(strQry)
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ''' <summary>Fill data in Tmp Table.</summary>
    Private Function GetRateWisePartDetail() As System.Data.DataTable
        Dim strQry As String = String.Empty
        Dim odt As New System.Data.DataTable
        Try
            For Each odr As DataRow In dtSelItems.Rows
                strQry = strQry + Convert.ToString(odr("ITEM_CODE")).Trim() + "','"
            Next
            If String.IsNullOrEmpty(strQry) Then
                MessageBox.Show("Select atlease one Part Code.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Return odt
            End If
            strQry = strQry.Remove(strQry.LastIndexOf("','"), 3)
            strQry = "SELECT CASE ISNULL(SRQ.ITEM_CODE,'') WHEN '' THEN 0 ELSE 1 END [SELECT],TMP.CUSTOMER_CODE, TMP.ITEM_CODE, TMP.VARMODEL" & _
                    " , TMP.RATE, SUM(TMP.INVOICE_QTY) INVOICE_QTY, SUM(TMP.ACTUALINVOICEQTY) ACTUALINVOICEQTY, ISNULL(SRQ.PRICECHANGETYPE, TMP.PRICECHANGETYPE) PRICECHANGETYPE" & _
                    " , ISNULL(SRQ.NEWRATE, TMP.NEWRATE) NEWRATE, ISNULL(SRQ.CHANGEEFFECT, TMP.CHANGEEFFECT) CHANGEEFFECT" & _
                    " , ISNULL(SRQ.PERCENTAGECHANGE, TMP.PERCENTAGECHANGE) PERCENTAGECHANGE, ISNULL(SRQ.CHANGEREASON, TMP.CHANGEREASON) CHANGEREASON" & _
                    " , ISNULL(SPNC.CORRECTIONDESC,'') CORRECTIONDESC, ISNULL(SPNC.CORRECTIONID,'') CORRECTIONID" & _
                    " FROM Sales_Prov_tmpPartDetail_Auth TMP LEFT JOIN SALES_PROV_RATEWISEDETAIL SRQ" & _
                    " ON TMP.CUSTOMER_CODE=SRQ.CUSTOMER_CODE AND TMP.ITEM_CODE=SRQ.ITEM_CODE AND TMP.VARMODEL=SRQ.VARMODEL AND TMP.UNIT_CODE=SRQ.UNIT_CODE" & _
                    " AND TMP.RATE=SRQ.RATE AND CAST(SRQ.PROV_DOCNO AS VARCHAR(20))='" + txtProvDocNo.Text.Trim() + "'" & _
                    " LEFT JOIN SALES_PROV_NATUREOFCORRECTION SPNC ON SRQ.NATUREOFCORRECTION=SPNC.CORRECTIONID" & _
                    " WHERE TMP.IPADDRESS='" + gstrIpaddressWinSck + "' AND TMP.UNIT_CODE='" + gstrUNITID + "'" & _
                    " and tmp.Item_code in ('" + strQry.Trim() + "')" & _
                    " GROUP BY TMP.CUSTOMER_CODE, TMP.ITEM_CODE, TMP.VARMODEL, CASE ISNULL(SRQ.ITEM_CODE,'') WHEN '' THEN 0 ELSE 1 END, TMP.RATE" & _
                    " , ISNULL(SRQ.PRICECHANGETYPE, TMP.PRICECHANGETYPE), ISNULL(SRQ.NEWRATE, TMP.NEWRATE), ISNULL(SRQ.CHANGEEFFECT, TMP.CHANGEEFFECT)" & _
                    " , ISNULL(SRQ.PERCENTAGECHANGE, TMP.PERCENTAGECHANGE), ISNULL(SRQ.CHANGEREASON, TMP.CHANGEREASON), ISNULL(SPNC.CORRECTIONDESC,'')" & _
                    " , ISNULL(SPNC.CORRECTIONID,'')"
            odt = SqlConnectionclass.GetDataTable(strQry)
        Catch ex As Exception
            RaiseException(ex)
        End Try
        Return odt
    End Function

    ''' <summary>
    ''' Lock/Unlock Rate wise grid columns
    ''' <param name="row1">Starting rowindex</param>
    ''' <paramref name=" row2">Ending rowIndex</paramref>
    ''' <paramref name="status">True for Locking and false for unlocking grid</paramref>
    ''' </summary>
    Private Sub LockUnlockRateWiseGrid(ByVal row1 As Integer, ByVal row2 As Integer, ByVal status As Boolean)
        Dim val As Integer = 0
        Try
            With fpSpreadRateWiseDtl
                .Col = RateWiseDtlGrid.Col_Select
                .Col2 = RateWiseDtlGrid.Col_CorrectionNature
                .Row = row1
                .Row2 = row2
                .BlockMode = True
                .Lock = True
                .BlockMode = False
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Calculate Invoice Rate
    ''' </summary>
    ''' <param name="row"></param>
    ''' <remarks></remarks>
    Private Sub CalculateRateEffect(ByVal row As Integer)
        Dim InvRate As Double = 0.0
        Dim RateChange As Double = 0.0
        Dim NewInvRate As Double = 0.0
        Dim TotEffVal As Double = 0.0
        Dim priceChangeType As Integer = 0
        Dim changeEffect As Integer = 0
        Dim ActInvQty As Double = 0.0
        Try
            With fpSpreadRateWiseDtl
                .Row = row
                .Col = RateWiseDtlGrid.Col_InvRate
                InvRate = Convert.ToDecimal(.Value)
                .Col = RateWiseDtlGrid.Col_PriceChange
                priceChangeType = .Value
                .Col = RateWiseDtlGrid.Col_ChangeEff
                changeEffect = .Value
                .Col = RateWiseDtlGrid.Col_ActInvQty
                ActInvQty = Convert.ToDouble(.Value)
                If priceChangeType = 0 Then
                    .Col = RateWiseDtlGrid.Col_NewRate
                    If String.IsNullOrEmpty(.Value) Or .Value = "0.00" Then
                        RateChange = 0.0
                        .Value = ""
                    Else
                        RateChange = Math.Round(Convert.ToDouble(.Value), 2)
                    End If
                Else
                    .Col = RateWiseDtlGrid.Col_Change
                    If String.IsNullOrEmpty(.Value) Or .Value = "0.00" Then
                        RateChange = 0.0
                        .Value = ""
                    Else
                        RateChange = Math.Round(Convert.ToDouble(.Value), 2)
                    End If
                    .Col = RateWiseDtlGrid.Col_NewInvRate
                    If RateChange = 0.0 Then
                        .Value = ""
                    ElseIf changeEffect = 0 Then
                        RateChange = Math.Round(InvRate * (1 - RateChange / 100), 2)
                        .Value = RateChange
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatMax = 9999999999.99
                        .TypeFloatMin = 0
                        .TypeFloatSeparator = False
                        .TypeFloatDecimalPlaces = 2
                    Else
                        RateChange = Math.Round(InvRate * (1 + RateChange / 100), 2)
                        .Value = RateChange
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatMax = 9999999999.99
                        .TypeFloatMin = 0
                        .TypeFloatSeparator = False
                        .TypeFloatDecimalPlaces = 2
                    End If
                End If
                .Col = RateWiseDtlGrid.Col_TotEffVal
                TotEffVal = ActInvQty * (RateChange - InvRate)
                If RateChange = 0.0 Then
                    .Value = ""
                Else
                    .Value = TotEffVal
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                    .TypeFloatMax = 9999999999.99
                    .TypeFloatMin = 0
                    .TypeFloatSeparator = False
                    .TypeFloatDecimalPlaces = 2
                End If
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Clear all tax Text Boxes.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearAllTaxes()
        Try
            txtExcise.Text = String.Empty
            txtPerExcise.Text = String.Empty
            txtCESS.Text = String.Empty
            txtPerCESS.Text = String.Empty
            txtSalesTaxCode.Text = String.Empty
            txtPerSalesTax.Text = String.Empty
            txtAEDCode.Text = String.Empty
            txtPerAED.Text = String.Empty
            txtSECESS.Text = String.Empty
            txtPerSECESS.Text = String.Empty
            txtSurcharge.Text = String.Empty
            txtPerSurcharge.Text = String.Empty
            txtAddVAT.Text = String.Empty
            txtPerAddVAT.Text = String.Empty
            txtExciseVal.Text = String.Empty
            txtCESSVal.Text = String.Empty
            txtSalesTaxVal.Text = String.Empty
            txtAEDVal.Text = String.Empty
            txtSECESSVal.Text = String.Empty
            txtSurchargeVal.Text = String.Empty
            txtAddVATVal.Text = String.Empty
            txtTotAssVal.Text = String.Empty
            txtNetVal.Text = String.Empty
            txtRoundOffBy.Text = String.Empty
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Generate Credit Note document No.
    ''' </summary>
    ''' <param name="pstrUnitCode"></param>
    ''' <param name="pdtDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GenerateDocNO(ByVal pstrUnitCode As String, ByVal pdtDate As Date) As Boolean
        Dim varRetVal As Object
        Dim varDocNoRow As Object
        Dim varDocNoCol As Object
        Dim strRetStr As String
        Try

            GenerateDocNO = False
            If Trim(pstrUnitCode) <> "" And IsDate(pdtDate) Then

                varRetVal = objCrDr.GetDocumentNumber(prj_DrCrNote.cls_DrCrNote.udtDocumentType.doctAccountsReceivable, getDateForDB(pdtDate), pstrUnitCode, ConnectionString:=gstrCONNECTIONSTRING)
                'Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                strRetStr = CheckString(varRetVal)
                If Trim(strRetStr) <> "Y" Then
                    strFinDocNo = ""
                    MsgBox(strRetStr, MsgBoxStyle.Critical, ResolveResString(100))
                    GenerateDocNO = False
                    Exit Function
                Else
                    varRetVal = Mid(varRetVal, 3)
                    varDocNoRow = SplitIntoRows(varRetVal)
                    varDocNoCol = SplitIntoColumns(varDocNoRow(0))
                    strFinDocNo = Trim(varDocNoCol(0))
                    GenerateDocNO = True
                    Exit Function
                End If
            Else
                strFinDocNo = ""
                GenerateDocNO = True
            End If
            Exit Function

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function

    ''' <summary>
    ''' Get exchange Rate
    ''' </summary>
    ''' <param name="pstrCurCode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ShowExchangeRate(ByVal pstrCurCode As String) As String
        Dim varCurDetails As Object = Nothing
        Dim varCurRow As Object = Nothing
        Dim varCurCol As Object = Nothing
        Dim strRetStr As String = String.Empty
        Try
            pstrCurCode = Trim(QuoteRem(pstrCurCode))
            varCurDetails = mobjExchangeRate.GetExchangeRate(pstrCurCode, getDateForDB(ServerDate), prj_ExchRateGetter.cls_ExchRateGetter.udtExchangeRateType.ertSellingRate, gstrCONNECTIONSTRING)
            strRetStr = CheckString(varCurDetails)
            If Trim(strRetStr) = "Y" Then
                varCurDetails = Mid(varCurDetails, 3)
                varCurRow = SplitIntoRows(varCurDetails) 'Getting Rows from Details
                varCurCol = SplitIntoColumns(varCurRow(0)) 'Getting Col from Details
                Return VB6.Format(Trim(varCurCol(0)), "0.0000")
            Else
                Return "1"
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function

    ''' <summary>
    ''' Generate Credit Note.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function generateCreditNote() As Boolean
        Dim mstrMasterDataString As String = String.Empty
        Dim mstrDetailDataString As String = String.Empty
        Dim result As Object
        Dim varPartyDetails As Object
        Dim varPartyRow As Object
        Dim varPartyCol As Object
        Dim varPartyBal As Object
        Dim strRetBalance As String = String.Empty
        Dim strRetBalStr As String = String.Empty
        Dim strRetStr As String = String.Empty
        Dim dtpDueDate As Date
        Dim dtpPayDueDate As Date
        Dim dtpExpPayDate As Date
        Dim mstrCtrlGLAc As String = String.Empty
        Dim mstrCtrlSLAc As String = String.Empty
        Dim ExchRate As String = "0.00"
        Dim strCurrencyCode As String = String.Empty
        Dim DebitGL As String = String.Empty
        Dim DebitGLBal As String = "0"
        Dim DebitGLName As String = String.Empty
        Dim DebitSL As String = String.Empty
        Dim DebitCCCode As String = String.Empty
        Dim odt As New System.Data.DataTable
        Dim count As Int16
        Try
            GenerateDocNO(gstrUNITID, getDateForDB(ServerDate))
            varPartyDetails = objCrDr.GetCustomerDetails(gstrUNITID.Trim, txtCustCode.Text.Trim(), getDateForDB(ServerDate), prj_FinanceCommon.cls_common.udtExchangeRateType.ertMoneyIn, gstrCONNECTIONSTRING)
            odt = SqlConnectionclass.GetDataTable("SELECT ISNULL(DEBITGL,'') DEBITGL, ISNULL(DEBITSL,'') DEBITSL, ISNULL(DEBITCCCODE,'') DEBITCCCODE" & _
                                                  " FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'")
            If odt.Rows.Count > 0 Then
                DebitGL = Convert.ToString(odt.Rows(0)("DebitGL"))
                If String.IsNullOrEmpty(DebitGL) Then
                    MessageBox.Show("Please define Debit GL for Unit Code " + gstrUNITID + ".", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If
                DebitSL = Convert.ToString(odt.Rows(0)("DebitSL"))
                If String.IsNullOrEmpty(DebitSL) Then
                    MessageBox.Show("Please define Debit SL for Unit Code " + gstrUNITID + ".", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If
                DebitCCCode = Convert.ToString(odt.Rows(0)("DebitCCCode"))
            End If
            strRetStr = CheckString(varPartyDetails)
            If Trim(strRetStr) = "Y" Then
                varPartyDetails = Mid(varPartyDetails, 3)
                varPartyRow = SplitIntoRows(varPartyDetails) 'Getting Rows from Details
                varPartyCol = SplitIntoColumns(varPartyRow(0)) 'Getting Col from Details
                strCurrencyCode = Trim(varPartyCol(5))
                ExchRate = ShowExchangeRate(strCurrencyCode)
                If String.IsNullOrEmpty(ExchRate) Then
                    MessageBox.Show("Please define Exchange Rate for Unit Code " + gstrUNITID + ".", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If
                mstrCtrlGLAc = Trim(varPartyCol(1))
                If String.IsNullOrEmpty(mstrCtrlGLAc) Then
                    MessageBox.Show("Please define Credit GL for Unit Code " + gstrUNITID + ".", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If
                mstrCtrlSLAc = Trim(varPartyCol(3))
                If String.IsNullOrEmpty(mstrCtrlSLAc) Then
                    MessageBox.Show("Please define Credit SL for Unit Code " + gstrUNITID + ".", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If
                dtpDueDate = IIf(IsDate(varPartyCol(11)), varPartyCol(11), GetServerDate)
                dtpPayDueDate = IIf(IsDate(varPartyCol(12)), varPartyCol(12), GetServerDate)
                dtpExpPayDate = IIf(IsDate(varPartyCol(13)), varPartyCol(13), GetServerDate)


            Else
                MessageBox.Show(strRetStr, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Return False
            End If
            mstrMasterDataString = "M»" + strFinDocNo.Trim() + "»" + ServerDate.ToString("dd MMM yyyy") + "»0»»" + gstrUNITID + "»" + txtCustCode.Text.Trim() & _
                                   "»»" + ServerDate.ToString("dd MMM yyyy") + "»" + dtpDueDate.ToString("dd MMM yyyy") + "»" + dtpPayDueDate.ToString("dd MMM yyyy") & _
                                   "»" + dtpExpPayDate.ToString("dd MMM yyyy") + "»INR»1»" + (Convert.ToDouble(txtNetVal_CR.Text.Trim())).ToString("0.00") & _
                                   "»0»»»Sales Provisioning Header»»»CR»" + mstrCtrlGLAc + "»" + mstrCtrlSLAc + "» »INR»" + mP_User + "»getdate()»0»AR»»»»" + ExchRate + "¦"

            'mstrDetailDataString = "M»" + strFinDocNo.Trim() + "»1»»»" + DebitGL + "»" + DebitSL + "»" + DebitCCCode + "»»»DR»" & _
            '                       (Convert.ToDouble(txtNegativeVal.Text.Trim()) * (-1)).ToString("0.00") + "»" + "Sales Provisining" + "»" + DebitGLName + "»" & _
            '                       DebitGLBal + "»»0»»0»»0»»0¦"

            '10816097 — eMPro -- Sales Provisioning 
            count = 1

            mstrDetailDataString = "M»" + strFinDocNo.Trim() + "»" + count.ToString() + "»»»" + DebitGL + "»" + DebitSL + "»" + DebitCCCode + "»»»DR»" & _
                                  (Convert.ToDouble(txtNegativeVal.Text.Trim()) * (-1)).ToString("0.00") + "»»" + "Sales" + "»" + DebitGLBal + "»" & _
                                  DebitGLBal + "»0»»0»»0»»0¦"

            If Val(txtExciseVal_CR.Text.Trim) > 0.0 Then
                odt.Clear()
                odt = SqlConnectionclass.GetDataTable("SELECT ISNULL(tx_taxid,'') TAX_ID, ISNULL(TxRt_Rate_No,'') TAX_RATE, ISNULL(tx_glCode,'') AS DEBITGL, ISNULL(tx_slCode,'') As DEBITSL FROM fin_TaxGlRel A inner join gen_taxrate B on a.UNIT_CODE = b.UNIT_CODE and a.tx_taxId = b.Tx_TaxeID WHERE A.UNIT_CODE='" + gstrUNITID + "'  AND tx_rowType = 'ARTAX' AND TxRt_Rate_No ='" + txtExcise_CR.Text.Trim().ToString() + "'")

                DebitGL = Convert.ToString(odt.Rows(0)("DebitGL"))
                If String.IsNullOrEmpty(DebitGL) Then
                    MessageBox.Show("Please define Credit GL for Tax " + txtExcise_CR.Text.Trim().ToString() + ".", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If

                DebitSL = Convert.ToString(odt.Rows(0)("DebitSL"))
                If String.IsNullOrEmpty(DebitSL) Then
                    MessageBox.Show("Please define Credit SL for Tax " + txtExcise_CR.Text.Trim().ToString() + ".", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If

                count = count + 1
                mstrDetailDataString = mstrDetailDataString + "M»" + strFinDocNo.Trim() + "»" + count.ToString() + "»»»" + DebitGL.Trim() + "»" + DebitSL.Trim() + "»»»»DR»" & _
                     (Convert.ToDouble(txtExciseVal_CR.Text.Trim())).ToString("0.00") + "»»" + "Excise" + "»" + DebitGLBal + "»" & _
                     DebitGLBal + "»0»»0»»0»»0¦"

            End If


            If Val(txtAEDVal_CR.Text.Trim) > 0.0 Then
                odt.Clear()
                odt = SqlConnectionclass.GetDataTable("SELECT ISNULL(tx_taxid,'') TAX_ID, ISNULL(TxRt_Rate_No,'') TAX_RATE, ISNULL(tx_glCode,'') AS DEBITGL, ISNULL(tx_slCode,'') As DEBITSL FROM fin_TaxGlRel A inner join gen_taxrate B on a.UNIT_CODE = b.UNIT_CODE and a.tx_taxId = b.Tx_TaxeID WHERE A.UNIT_CODE='" + gstrUNITID + "'  AND tx_rowType = 'ARTAX' AND TxRt_Rate_No ='" + txtAEDCode_CR.Text.Trim().ToString() + "'")

                DebitGL = Convert.ToString(odt.Rows(0)("DebitGL"))
                If String.IsNullOrEmpty(DebitGL) Then
                    MessageBox.Show("Please define Credit GL for Tax " + txtAEDCode_CR.Text.Trim().ToString() + ".", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If

                DebitSL = Convert.ToString(odt.Rows(0)("DebitSL"))
                If String.IsNullOrEmpty(DebitSL) Then
                    MessageBox.Show("Please define Credit SL for Tax " + txtAEDCode_CR.Text.Trim().ToString() + ".", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If
                count = count + 1
                mstrDetailDataString = mstrDetailDataString + "M»" + strFinDocNo.Trim() + "»" + count.ToString() + "»»»" + DebitGL.Trim() + "»" + DebitSL.Trim() + "»»»»DR»" & _
                     (Convert.ToDouble(txtAEDVal_CR.Text.Trim())).ToString("0.00") + "»»" + "AED" + "»" + DebitGLBal + "»" & _
                     DebitGLBal + "»0»»0»»0»»0¦"

            End If


            If Val(txtCESSVal_CR.Text.Trim) > 0.0 Then
                odt.Clear()
                odt = SqlConnectionclass.GetDataTable("SELECT ISNULL(tx_taxid,'') TAX_ID, ISNULL(TxRt_Rate_No,'') TAX_RATE, ISNULL(tx_glCode,'') AS DEBITGL, ISNULL(tx_slCode,'') As DEBITSL FROM fin_TaxGlRel A inner join gen_taxrate B on a.UNIT_CODE = b.UNIT_CODE and a.tx_taxId = b.Tx_TaxeID WHERE A.UNIT_CODE='" + gstrUNITID + "'  AND tx_rowType = 'ARTAX' AND TxRt_Rate_No ='" + txtCESS_CR.Text.Trim().ToString() + "'")

                DebitGL = Convert.ToString(odt.Rows(0)("DebitGL"))
                If String.IsNullOrEmpty(DebitGL) Then
                    MessageBox.Show("Please define Credit GL for Tax " + txtCESS_CR.Text.Trim().ToString() + ".", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If

                DebitSL = Convert.ToString(odt.Rows(0)("DebitSL"))
                If String.IsNullOrEmpty(DebitSL) Then
                    MessageBox.Show("Please define Credit SL for Tax " + txtCESS_CR.Text.Trim().ToString() + ".", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If
                count = count + 1
                mstrDetailDataString = mstrDetailDataString + "M»" + strFinDocNo.Trim() + "»" + count.ToString() + "»»»" + DebitGL.Trim() + "»" + DebitSL.Trim() + "»»»»DR»" & _
                     (Convert.ToDouble(txtCESSVal_CR.Text.Trim())).ToString("0.00") + "»»" + "CESS" + "»" + DebitGLBal + "»" & _
                     DebitGLBal + "»0»»0»»0»»0¦"

            End If


            If Val(txtSECESSVal_CR.Text.Trim) > 0.0 Then
                odt.Clear()
                odt = SqlConnectionclass.GetDataTable("SELECT ISNULL(tx_taxid,'') TAX_ID, ISNULL(TxRt_Rate_No,'') TAX_RATE, ISNULL(tx_glCode,'') AS DEBITGL, ISNULL(tx_slCode,'') As DEBITSL FROM fin_TaxGlRel A inner join gen_taxrate B on a.UNIT_CODE = b.UNIT_CODE and a.tx_taxId = b.Tx_TaxeID WHERE A.UNIT_CODE='" + gstrUNITID + "'  AND tx_rowType = 'ARTAX' AND TxRt_Rate_No ='" + txtSECESS_CR.Text.Trim().ToString() + "'")

                DebitGL = Convert.ToString(odt.Rows(0)("DebitGL"))
                If String.IsNullOrEmpty(DebitGL) Then
                    MessageBox.Show("Please define Credit GL for Tax " + txtSECESS_CR.Text.Trim().ToString() + ".", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If

                DebitSL = Convert.ToString(odt.Rows(0)("DebitSL"))
                If String.IsNullOrEmpty(DebitSL) Then
                    MessageBox.Show("Please define Credit SL for Tax " + txtSECESS_CR.Text.Trim().ToString() + ".", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If
                count = count + 1
                mstrDetailDataString = mstrDetailDataString + "M»" + strFinDocNo.Trim() + "»" + count.ToString() + "»»»" + DebitGL.Trim() + "»" + DebitSL.Trim() + "»»»»DR»" & _
                     (Convert.ToDouble(txtSECESSVal_CR.Text.Trim())).ToString("0.00") + "»»" + "SECESS" + "»" + DebitGLBal + "»" & _
                     DebitGLBal + "»0»»0»»0»»0¦"

            End If


            If Val(txtSalesTaxVal_CR.Text.Trim) > 0.0 Then
                odt.Clear()
                odt = SqlConnectionclass.GetDataTable("SELECT ISNULL(tx_taxid,'') TAX_ID, ISNULL(TxRt_Rate_No,'') TAX_RATE, ISNULL(tx_glCode,'') AS DEBITGL, ISNULL(tx_slCode,'') As DEBITSL FROM fin_TaxGlRel A inner join gen_taxrate B on a.UNIT_CODE = b.UNIT_CODE and a.tx_taxId = b.Tx_TaxeID WHERE A.UNIT_CODE='" + gstrUNITID + "'  AND tx_rowType = 'ARTAX' AND TxRt_Rate_No ='" + txtSalesTaxCode_CR.Text.Trim().ToString() + "'")

                DebitGL = Convert.ToString(odt.Rows(0)("DebitGL"))
                If String.IsNullOrEmpty(DebitGL) Then
                    MessageBox.Show("Please define Credit GL for Tax " + txtSalesTaxCode_CR.Text.Trim().ToString() + ".", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If

                DebitSL = Convert.ToString(odt.Rows(0)("DebitSL"))
                If String.IsNullOrEmpty(DebitSL) Then
                    MessageBox.Show("Please define Credit SL for Tax " + txtSalesTaxCode_CR.Text.Trim().ToString() + ".", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If
                count = count + 1
                mstrDetailDataString = mstrDetailDataString + "M»" + strFinDocNo.Trim() + "»" + count.ToString() + "»»»" + DebitGL.Trim() + "»" + DebitSL.Trim() + "»»»»DR»" & _
                     (Convert.ToDouble(txtSalesTaxVal_CR.Text.Trim())).ToString("0.00") + "»»" + "Sales Tax" + "»" + DebitGLBal + "»" & _
                     DebitGLBal + "»0»»0»»0»»0¦"

            End If



            If Val(txtAddVATVal_CR.Text.Trim) > 0.0 Then
                odt.Clear()
                odt = SqlConnectionclass.GetDataTable("SELECT ISNULL(tx_taxid,'') TAX_ID, ISNULL(TxRt_Rate_No,'') TAX_RATE, ISNULL(tx_glCode,'') AS DEBITGL, ISNULL(tx_slCode,'') As DEBITSL FROM fin_TaxGlRel A inner join gen_taxrate B on a.UNIT_CODE = b.UNIT_CODE and a.tx_taxId = b.Tx_TaxeID WHERE A.UNIT_CODE='" + gstrUNITID + "'  AND tx_rowType = 'ARTAX' AND TxRt_Rate_No ='" + txtAddVAT_CR.Text.Trim().ToString() + "'")

                DebitGL = Convert.ToString(odt.Rows(0)("DebitGL"))
                If String.IsNullOrEmpty(DebitGL) Then
                    MessageBox.Show("Please define Credit GL for Tax " + txtAddVAT_CR.Text.Trim().ToString() + ".", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If

                DebitSL = Convert.ToString(odt.Rows(0)("DebitSL"))
                If String.IsNullOrEmpty(DebitSL) Then
                    MessageBox.Show("Please define Credit SL for Tax " + txtAddVAT_CR.Text.Trim().ToString() + ".", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If
                count = count + 1
                mstrDetailDataString = mstrDetailDataString + "M»" + strFinDocNo.Trim() + "»" + count.ToString() + "»»»" + DebitGL.Trim() + "»" + DebitSL.Trim() + "»»»»DR»" & _
                     (Convert.ToDouble(txtAddVATVal_CR.Text.Trim())).ToString("0.00") + "»»" + "ADD VAT" + "»" + DebitGLBal + "»" & _
                     DebitGLBal + "»0»»0»»0»»0¦"

            End If


            If Val(txtSurchargeVal_CR.Text.Trim) > 0.0 Then
                odt.Clear()
                odt = SqlConnectionclass.GetDataTable("SELECT ISNULL(tx_taxid,'') TAX_ID, ISNULL(TxRt_Rate_No,'') TAX_RATE, ISNULL(tx_glCode,'') AS DEBITGL, ISNULL(tx_slCode,'') As DEBITSL FROM fin_TaxGlRel A inner join gen_taxrate B on a.UNIT_CODE = b.UNIT_CODE and a.tx_taxId = b.Tx_TaxeID WHERE A.UNIT_CODE='" + gstrUNITID + "'  AND tx_rowType = 'ARTAX' AND TxRt_Rate_No ='" + txtSurcharge_CR.Text.Trim().ToString() + "'")

                DebitGL = Convert.ToString(odt.Rows(0)("DebitGL"))
                If String.IsNullOrEmpty(DebitGL) Then
                    MessageBox.Show("Please define Credit GL for Tax " + txtSurcharge_CR.Text.Trim().ToString() + ".", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If

                DebitSL = Convert.ToString(odt.Rows(0)("DebitSL"))
                If String.IsNullOrEmpty(DebitSL) Then
                    MessageBox.Show("Please define Credit SL for Tax " + txtSurcharge_CR.Text.Trim().ToString() + ".", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return False
                End If
                count = count + 1
                mstrDetailDataString = mstrDetailDataString + "M»" + strFinDocNo.Trim() + "»" + count.ToString() + "»»»" + DebitGL.Trim() + "»" + DebitSL.Trim() + "»»»»DR»" & _
                     (Convert.ToDouble(txtSurchargeVal_CR.Text.Trim())).ToString("0.00") + "»»" + "Surcharge" + "»" + DebitGLBal + "»" & _
                     DebitGLBal + "»0»»0»»0»»0¦"

            End If

            ' 10816097 — eMPro -- Sales Provisioning


            'result = objCrDr.SetAPDocument(gstrUNITID, mstrMasterDataString, mstrDetailDataString, prj_DrCrNote.cls_DrCrNote.udtOperationType.optInsert, gstrCONNECTIONSTRING)
            result = objCrDr.SetARDocument(gstrUNITID, mstrMasterDataString, mstrDetailDataString, prj_DrCrNote.cls_DrCrNote.udtOperationType.optInsert, gstrCONNECTIONSTRING)
            strRetStr = CheckString(result)
            If strRetStr = "Y" Then

                Return True
            Else
                MessageBox.Show(strRetStr, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Return False
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
        Return True
    End Function

    ''' <summary>
    ''' Create New Supplementary Invoice No.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SelectChallanNoFromSupplementatryInvHdr()
        Dim strQry As String = String.Empty
        Try
            strQry = "Select isnull(max(Doc_No),0) as Doc_No from SupplementaryInv_hdr where Unit_code='" & gstrUNITID & "' and Doc_No>" & 99000000
            strSupplInvNo = SqlConnectionclass.ExecuteScalar(strQry)
            If Val(strSupplInvNo) = 0 Then
                strSupplInvNo = "99000001"
            Else
                strSupplInvNo = CStr(Val(strSupplInvNo) + 1)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Generate Annexure in Excel Form
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub GenerateAnexure()
        Dim strQry As String = String.Empty
        Dim odt As New System.Data.DataTable
        Try
           
            strQry = "SELECT tbl.SNo, tbl.ITEM_CODE, DESCRIPTION, isnull(cust .Cust_Drgno,'') CustomeritemCode ,isnull(cust.Drg_Desc,'') CustomerItemDesc,INVOICE_NO, ' '+convert(varchar(20), INVOICEDATE,105) INVOICEDATE, cast(ROUND(RATE,2,1) as varchar(50)) RATE, " & _
                      "INVOICE_QTY, ACTUALINVOICEQTY, BASICAMOUNT, NEWRATE, NEWBASICAMOUNT , BASICAMOUNTDIFF  , " & _
                      "ISNULL([BOP Rate Revised],0.0) [BOP Rate Revised],ISNULL([Paint Rate Revised],0.0) [Paint Rate Revised], " & _
                      "ISNULL([Other],0.0) [Other],ISNULL([Claim],0.0) [Claim],ISNULL([RM Rate Revised],0.0) [RM Rate Revised] " & _
                      "FROM  ( " & _
                      "Select SNO,ITEM_CODE,DESCRIPTION,INVOICE_NO,INVOICEDATE,table1.Rate,INVOICE_QTY,ACTUALINVOICEQTY,BASICAMOUNT " & _
                      ",NEWRATE,NEWBASICAMOUNT,BASICAMOUNTDIFF,ISNULL([BOP Rate Revised],0.0) [BOP Rate Revised],ISNULL([Paint Rate Revised],0.0) [Paint Rate Revised], " & _
                      "ISNULL([Other],0.0) [Other],ISNULL([Claim],0.0) [Claim],ISNULL([RM Rate Revised],0.0) [RM Rate Revised] " & _
                      "from ( " & _
                      "SELECT row_number() over (partition by TMP.item_code, " & _
                      "TMP.RATE order by TMP.Item_COde)+1 SNO, TMP.ITEM_CODE, '' DESCRIPTION, TMP.INVOICE_NO, TMP.INVOICEDATE, TMP.RATE, SUM(ISNULL(TMP.INVOICE_QTY,0.0)) INVOICE_QTY, " & _
                      "SUM(ISNULL(TMP.ACTUALINVOICEQTY,0.0)) ACTUALINVOICEQTY, TMP.RATE*SUM(ISNULL(TMP.INVOICE_QTY,0.0)) BASICAMOUNT , ROUND(RWD.NEWRATE,2,1) NEWRATE  ,  " & _
                      "RWD.NEWRATE*SUM(ISNULL(TMP.INVOICE_QTY,0.0)) NEWBASICAMOUNT, RWD.NEWRATE*SUM(ISNULL(TMP.INVOICE_QTY,0.0))-TMP.RATE*SUM(ISNULL(TMP.INVOICE_QTY,0.0)) BASICAMOUNTDIFF " & _
                      "FROM Sales_Prov_tmpPartDetail_Auth TMP   INNER JOIN SALES_PROV_RATEWISEDETAIL RWD " & _
                      "ON TMP.CUSTOMER_CODE=RWD.CUSTOMER_CODE AND TMP.UNIT_CODE=RWD.UNIT_CODE  " & _
                      "AND TMP.ITEM_CODE=RWD.ITEM_CODE  and TMP.Rate=RWD.RATE " & _
                      "WHERE TMP.UNIT_CODE='" + gstrUNITID + "' AND TMP.IPADDRESS='" + gstrIpaddressWinSck + "' AND CONVERT(VARCHAR(20),  " & _
                      "RWD.PROV_DOCNO)='" + txtProvDocNo.Text.Trim() + "' " & _
                      "GROUP BY TMP.ITEM_CODE, TMP.INVOICE_NO, TMP.INVOICEDATE, TMP.RATE, RWD.NEWRATE)table1 " & _
                      "inner join " & _
                      "( " & _
                      "SELECT * FROM " & _
                      "(select Prov_DocNo,itemcode, reason, (effect+convert(varchar(10),value)) value , convert(numeric(9,2),new_rate) as new_rate, convert(numeric(9,2),rate) as rate from PRICECHANGE_DTL where PROV_DOCNO='" + txtProvDocNo.Text.Trim() + "' AND UNITCODE='" + gstrUNITID + "') tbl " & _
                      "pivot " & _
                      "(max(value) for Reason in ([BOP Rate Revised],[Paint Rate Revised],[Other], [Claim], [RM Rate Revised] )) pvttbl) " & _
                      "table2 " & _
                      "on table1 .Item_code =table2.itemcode and table1 .Rate =table2 .rate and table1 .NEWRATE =table2 .new_rate  " & _
                      "UNION ALL  " & _
                      "SELECT ROW_NUMBER() OVER (PARTITION BY RWD.ITEM_CODE,RWD.RATE ORDER BY RWD.ITEM_CODE) SNO, RWD.ITEM_CODE, IM.DESCRIPTION, " & _
                      "'' INVOICE_NO, '' INVOICEDATE, RWD.RATE RATE, 0.0 INVOICE_QTY, 0.0 ACTUALINVOICEQTY, 0.0 BASICAMOUNT  , 0.0 NEWRATE, 0.0 NEWBASICAMOUNT, " & _
                      "'0.0' BASICAMOUNTDIFF ,'0.0' [BOP Rate Revised],'0.0' [Paint Rate Revised], " & _
                      "'0.0' [Other],'0.0' [Claim],'0.0' [RM Rate Revised] " & _
                      "FROM SALES_PROV_RATEWISEDETAIL RWD   LEFT JOIN ITEM_MST IM   ON RWD.UNIT_CODE=IM.UNIT_CODE AND RWD.ITEM_CODE=IM.ITEM_CODE " & _
                      "WHERE RWD.UNIT_CODE='" + gstrUNITID + "' AND CONVERT(VARCHAR(20), RWD.PROV_DOCNO)='" + txtProvDocNo.Text.Trim() + "' " & _
                      "UNION ALL " & _
                      "SELECT row_number() over (partition by RWD.item_code,RWD.RATE order by RWD.Item_COde)+1+count(RWD.Item_Code), " & _
                      "RWD.ITEM_CODE,'PART TOTAL' DESCRIPTION, '' INVOICE_NO, '' INVOICEDATE, RWD.RATE RATE, 0.0 INVOICE_QTY, 0.0 ACTUALINVOICEQTY, " & _
                      "0.0 BASICAMOUNT  , 0.0 NEWRATE, 0.0 NEWBASICAMOUNT, SUM(RWD.NEWRATE*ISNULL(TMP.INVOICE_QTY,0.0))-SUM(TMP.RATE*ISNULL(TMP.INVOICE_QTY,0.0)) BASICAMOUNTDIFF, " & _
                      "'0.0' [BOP Rate Revised],'0.0' [Paint Rate Revised],'0.0' [Other],'0.0' [Claim],'0.0' [RM Rate Revised] " & _
                      "FROM Sales_Prov_tmpPartDetail_Auth TMP  INNER JOIN SALES_PROV_RATEWISEDETAIL RWD    ON TMP.CUSTOMER_CODE=RWD.CUSTOMER_CODE AND " & _
                      "TMP.UNIT_CODE=RWD.UNIT_CODE AND TMP.ITEM_CODE=RWD.ITEM_CODE  and TMP.Rate=RWD.Rate  WHERE TMP.UNIT_CODE='" + gstrUNITID + "' AND " & _
                      "TMP.IPADDRESS='" + gstrIpaddressWinSck + "' AND CONVERT(VARCHAR(20), RWD.PROV_DOCNO)='" + txtProvDocNo.Text.Trim() + "'  " & _
                      "GROUP BY RWD.item_code, RWD.RATE " & _
                      "UNION ALL   " & _
                      "SELECT 0 SNo,'' ITEM_CODE, 'PROVISION TOTAL' DESCRIPTION, " & _
                      "'' INVOICE_NO, '' INVOICEDATE, 0.0 RATE, 0.0 INVOICE_QTY, 0.0 ACTUALINVOICEQTY, 0.0 BASICAMOUNT  , 0.0 NEWRATE  , " & _
                      "0.0 NEWBASICAMOUNT, SUM(RWD.NEWRATE*ISNULL(TMP.INVOICE_QTY,0.0))-SUM(TMP.RATE*ISNULL(TMP.INVOICE_QTY,0.0)) BASICAMOUNTDIFF, " & _
                      "'0.0' [BOP Rate Revised],'0.0' [Paint Rate Revised],'0.0' [Other],'0.0' [Claim],'0.0' [RM Rate Revised]  " & _
                      "FROM Sales_Prov_tmpPartDetail_Auth TMP   INNER JOIN SALES_PROV_RATEWISEDETAIL RWD    ON TMP.CUSTOMER_CODE=RWD.CUSTOMER_CODE " & _
                      "AND TMP.UNIT_CODE=RWD.UNIT_CODE AND TMP.ITEM_CODE=RWD.ITEM_CODE  and TMP.Rate=RWD.RATE  WHERE TMP.UNIT_CODE='" + gstrUNITID + "' AND " & _
                      "TMP.IPADDRESS='" + gstrIpaddressWinSck + "' AND CONVERT(VARCHAR(20), RWD.PROV_DOCNO)='" + txtProvDocNo.Text.Trim() + "' " & _
                      ")TBL  LEFT JOIN CUSTITEM_MST CUST ON CUST .ITEM_CODE =TBL .ITEM_CODE  AND CUST .UNIT_CODE ='MTM' AND CUST.ACTIVE =1 ORDER BY TBL.ITEM_CODE desc, TBL.Rate, TBL.SNo "
           
            odt = SqlConnectionclass.GetDataTable(strQry)
            xApp = New Application()
            xbook = xApp.Workbooks.Add
            xSheet = xbook.Worksheets(1)
            If odt.Rows.Count > 0 Then
                If Not IsNothing(xSheet) Then
                    Dim ROW As Integer = 1
                    Dim HeaderRepeat As Boolean = False
                    xSheet.Cells(ROW, 1) = "Provision Doc. No."
                    xSheet.Cells(ROW, 1).BorderAround(, Excel.XlBorderWeight.xlMedium)
                    xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 1)).Font.Bold = True
                    xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 1)).Interior.Color = RGB(115, 151, 253)
                    xSheet.Cells(ROW, 2) = txtProvDocNo.Text.Trim()
                    xSheet.Cells(ROW, 2).BorderAround(, Excel.XlBorderWeight.xlMedium)
                    ROW = ROW + 1
                    For Each odr As DataRow In odt.Rows
                        If Convert.ToString(odr("DESCRIPTION")).Equals("PART TOTAL") Then
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 7)).Merge()
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 7)).Value = "Part Total"
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 7)).Font.Bold = True
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 7)).BorderAround(, XlBorderWeight.xlMedium)
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 7)).HorizontalAlignment = Excel.Constants.xlRight
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 8)).Interior.Color = RGB(115, 151, 253)
                            xSheet.Cells(ROW, 8) = Convert.ToString(odr("BasicAmountDiff"))
                            xSheet.Cells(ROW, 8).BorderAround(, Excel.XlBorderWeight.xlMedium)
                        ElseIf Convert.ToString(odr("DESCRIPTION")).Equals("PROVISION TOTAL") Then
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 7)).Merge()
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 7)).Value = "PROVISION TOTAL"
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 7)).Font.Bold = True
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 7)).BorderAround(, XlBorderWeight.xlMedium)
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 7)).HorizontalAlignment = Excel.Constants.xlRight
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 8)).Interior.Color = RGB(115, 151, 253)
                            xSheet.Cells(ROW, 8) = Convert.ToString(odr("BasicAmountDiff"))
                            xSheet.Cells(ROW, 8).BorderAround(, Excel.XlBorderWeight.xlMedium)
                        ElseIf String.IsNullOrEmpty(Convert.ToString(odr("DESCRIPTION")).Trim()) Then
                            If HeaderRepeat Then
                                xSheet.Cells(ROW, 1) = "Invoice No."
                                xSheet.Cells(ROW, 1).BorderAround(, Excel.XlBorderWeight.xlMedium)
                                xSheet.Cells(ROW, 2) = "Invoice Date"
                                xSheet.Cells(ROW, 2).BorderAround(, Excel.XlBorderWeight.xlMedium)
                                xSheet.Range(xSheet.Cells(ROW, 2), xSheet.Cells(ROW, 2)).HorizontalAlignment = Excel.Constants.xlCenter
                                xSheet.Cells(ROW, 3) = "Item Qty."
                                xSheet.Cells(ROW, 3).BorderAround(, Excel.XlBorderWeight.xlMedium)
                                xSheet.Cells(ROW, 4) = "Invoice Rate"
                                xSheet.Cells(ROW, 4).BorderAround(, Excel.XlBorderWeight.xlMedium)
                                xSheet.Cells(ROW, 5) = "Basic Amt."
                                xSheet.Cells(ROW, 5).BorderAround(, Excel.XlBorderWeight.xlMedium)
                                xSheet.Cells(ROW, 6) = "New Rate"
                                xSheet.Cells(ROW, 6).BorderAround(, Excel.XlBorderWeight.xlMedium)
                                xSheet.Cells(ROW, 7) = "New Basic Amt."
                                xSheet.Cells(ROW, 7).BorderAround(, Excel.XlBorderWeight.xlMedium)
                                xSheet.Cells(ROW, 8) = "Basic Value Diff."
                                xSheet.Cells(ROW, 8).BorderAround(, Excel.XlBorderWeight.xlMedium)
                                xSheet.Cells(ROW, 9) = "BOP Rate Revised."
                                xSheet.Cells(ROW, 9).BorderAround(, Excel.XlBorderWeight.xlMedium)
                                xSheet.Cells(ROW, 10) = "Paint Rate Revised."
                                xSheet.Cells(ROW, 10).BorderAround(, Excel.XlBorderWeight.xlMedium)
                                xSheet.Cells(ROW, 11) = "RM Rate Revised."
                                xSheet.Cells(ROW, 11).BorderAround(, Excel.XlBorderWeight.xlMedium)
                                xSheet.Cells(ROW, 12) = "Claim."
                                xSheet.Cells(ROW, 12).BorderAround(, Excel.XlBorderWeight.xlMedium)
                                xSheet.Cells(ROW, 13) = "Other."
                                xSheet.Cells(ROW, 13).BorderAround(, Excel.XlBorderWeight.xlMedium)
                                xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 13)).Interior.Color = RGB(170, 160, 199)
                                xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 13)).Font.Bold = True
                                HeaderRepeat = False
                                ROW = ROW + 1
                            End If
                            xSheet.Cells(ROW, 1) = Convert.ToString(odr("INVOICE_NO"))
                            xSheet.Cells(ROW, 1).BorderAround(, Excel.XlBorderWeight.xlThin)
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 1)).HorizontalAlignment = Excel.Constants.xlLeft
                            xSheet.Cells(ROW, 2) = Convert.ToString(odr("InvoiceDate"))
                            xSheet.Cells(ROW, 2).BorderAround(, Excel.XlBorderWeight.xlThin)
                            xSheet.Range(xSheet.Cells(ROW, 2), xSheet.Cells(ROW, 2)).HorizontalAlignment = Excel.Constants.xlRight
                            xSheet.Cells(ROW, 3) = Convert.ToString(odr("Invoice_Qty"))
                            xSheet.Cells(ROW, 3).BorderAround(, Excel.XlBorderWeight.xlThin)
                            xSheet.Range(xSheet.Cells(ROW, 3), xSheet.Cells(ROW, 3)).HorizontalAlignment = Excel.Constants.xlRight
                            xSheet.Cells(ROW, 4) = Convert.ToString(odr("Rate"))
                            xSheet.Cells(ROW, 4).BorderAround(, Excel.XlBorderWeight.xlThin)
                            xSheet.Range(xSheet.Cells(ROW, 4), xSheet.Cells(ROW, 4)).HorizontalAlignment = Excel.Constants.xlRight
                            xSheet.Cells(ROW, 5) = Convert.ToString(odr("BasicAmount"))
                            xSheet.Cells(ROW, 5).BorderAround(, Excel.XlBorderWeight.xlThin)
                            xSheet.Range(xSheet.Cells(ROW, 5), xSheet.Cells(ROW, 5)).HorizontalAlignment = Excel.Constants.xlRight
                            xSheet.Cells(ROW, 6) = Convert.ToString(odr("NewRate"))
                            xSheet.Cells(ROW, 6).BorderAround(, Excel.XlBorderWeight.xlThin)
                            xSheet.Range(xSheet.Cells(ROW, 6), xSheet.Cells(ROW, 6)).HorizontalAlignment = Excel.Constants.xlRight
                            xSheet.Cells(ROW, 7) = Convert.ToString(odr("NewBasicAmount"))
                            xSheet.Cells(ROW, 7).BorderAround(, Excel.XlBorderWeight.xlThin)
                            xSheet.Range(xSheet.Cells(ROW, 7), xSheet.Cells(ROW, 7)).HorizontalAlignment = Excel.Constants.xlRight
                            xSheet.Cells(ROW, 8) = Convert.ToString(odr("BasicAmountDiff"))
                            xSheet.Cells(ROW, 8).BorderAround(, Excel.XlBorderWeight.xlThin)
                            xSheet.Range(xSheet.Cells(ROW, 8), xSheet.Cells(ROW, 8)).HorizontalAlignment = Excel.Constants.xlRight
                            ' Mayur
                            xSheet.Cells(ROW, 9) = Convert.ToString(odr("BOP Rate Revised"))
                            xSheet.Cells(ROW, 9).BorderAround(, Excel.XlBorderWeight.xlThin)
                            xSheet.Range(xSheet.Cells(ROW, 9), xSheet.Cells(ROW, 9)).HorizontalAlignment = Excel.Constants.xlRight

                            xSheet.Cells(ROW, 10) = Convert.ToString(odr("Paint Rate Revised"))
                            xSheet.Cells(ROW, 10).BorderAround(, Excel.XlBorderWeight.xlThin)
                            xSheet.Range(xSheet.Cells(ROW, 10), xSheet.Cells(ROW, 10)).HorizontalAlignment = Excel.Constants.xlRight

                            xSheet.Cells(ROW, 11) = Convert.ToString(odr("RM Rate Revised"))
                            xSheet.Cells(ROW, 11).BorderAround(, Excel.XlBorderWeight.xlThin)
                            xSheet.Range(xSheet.Cells(ROW, 11), xSheet.Cells(ROW, 11)).HorizontalAlignment = Excel.Constants.xlRight

                            xSheet.Cells(ROW, 12) = Convert.ToString(odr("Claim"))
                            xSheet.Cells(ROW, 12).BorderAround(, Excel.XlBorderWeight.xlThin)
                            xSheet.Range(xSheet.Cells(ROW, 12), xSheet.Cells(ROW, 12)).HorizontalAlignment = Excel.Constants.xlRight

                            xSheet.Cells(ROW, 13) = Convert.ToString(odr("Other"))
                            xSheet.Cells(ROW, 13).BorderAround(, Excel.XlBorderWeight.xlThin)
                            xSheet.Range(xSheet.Cells(ROW, 13), xSheet.Cells(ROW, 13)).HorizontalAlignment = Excel.Constants.xlRight
                            ' Mayur
                        Else
                            xSheet.Cells(ROW, 1) = "Part Code"
                            xSheet.Cells(ROW, 1).BorderAround(, Excel.XlBorderWeight.xlMedium)
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 1)).Interior.Color = RGB(230, 184, 183)
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 1)).Font.Bold = True
                            xSheet.Cells(ROW, 2) = odr("ITEM_CODE")
                            xSheet.Cells(ROW, 2).BorderAround(, Excel.XlBorderWeight.xlMedium)
                            xSheet.Cells(ROW, 3) = "Part Desc."
                            xSheet.Cells(ROW, 3).BorderAround(, Excel.XlBorderWeight.xlMedium)
                            xSheet.Range(xSheet.Cells(ROW, 3), xSheet.Cells(ROW, 3)).Interior.Color = RGB(230, 184, 183)
                            xSheet.Range(xSheet.Cells(ROW, 3), xSheet.Cells(ROW, 3)).Font.Bold = True
                            xSheet.Range(xSheet.Cells(ROW, 4), xSheet.Cells(ROW, 6)).Merge()
                            xSheet.Range(xSheet.Cells(ROW, 4), xSheet.Cells(ROW, 6)).Value = Convert.ToString(odr("DESCRIPTION"))
                            xSheet.Range(xSheet.Cells(ROW, 4), xSheet.Cells(ROW, 6)).BorderAround(, Excel.XlBorderWeight.xlMedium)

                            xSheet.Cells(ROW, 7) = "Customer Part Code"
                            xSheet.Cells(ROW, 7).BorderAround(, Excel.XlBorderWeight.xlMedium)
                            xSheet.Range(xSheet.Cells(ROW, 7), xSheet.Cells(ROW, 7)).Interior.Color = RGB(230, 184, 183)
                            xSheet.Range(xSheet.Cells(ROW, 7), xSheet.Cells(ROW, 7)).Font.Bold = True
                            xSheet.Cells(ROW, 8) = odr("CustomeritemCode")
                            xSheet.Cells(ROW, 8).BorderAround(, Excel.XlBorderWeight.xlMedium)
                            xSheet.Cells(ROW, 9) = "Customer Part Desc."
                            xSheet.Cells(ROW, 9).BorderAround(, Excel.XlBorderWeight.xlMedium)
                            xSheet.Range(xSheet.Cells(ROW, 9), xSheet.Cells(ROW, 9)).Interior.Color = RGB(230, 184, 183)
                            xSheet.Range(xSheet.Cells(ROW, 9), xSheet.Cells(ROW, 9)).Font.Bold = True
                            xSheet.Range(xSheet.Cells(ROW, 10), xSheet.Cells(ROW, 13)).Merge()
                            xSheet.Range(xSheet.Cells(ROW, 10), xSheet.Cells(ROW, 13)).Value = Convert.ToString(odr("CustomerItemDesc"))
                            xSheet.Range(xSheet.Cells(ROW, 10), xSheet.Cells(ROW, 13)).BorderAround(, Excel.XlBorderWeight.xlMedium)

                            HeaderRepeat = True
                        End If
                        ROW = ROW + 1
                    Next
                    xSheet.Cells(ROW, 1).EntireColumn.ColumnWidth = 18
                    xSheet.Cells(ROW, 2).EntireColumn.ColumnWidth = 20
                    xSheet.Cells(ROW, 3).EntireColumn.ColumnWidth = 10
                    xSheet.Cells(ROW, 4).EntireColumn.ColumnWidth = 10
                    xSheet.Cells(ROW, 5).EntireColumn.ColumnWidth = 10
                    xSheet.Cells(ROW, 6).EntireColumn.ColumnWidth = 10
                    xSheet.Cells(ROW, 7).EntireColumn.ColumnWidth = 10
                    xSheet.Cells(ROW, 8).EntireColumn.ColumnWidth = 10
                    xSheet.Cells(ROW, 9).EntireColumn.ColumnWidth = 10
                    xSheet.Cells(ROW, 10).EntireColumn.ColumnWidth = 10
                    xSheet.Cells(ROW, 11).EntireColumn.ColumnWidth = 10
                    xSheet.Cells(ROW, 12).EntireColumn.ColumnWidth = 10
                    xSheet.Cells(ROW, 13).EntireColumn.ColumnWidth = 10
                    xSheet.Cells.WrapText = True
                    xSheet.Cells.VerticalAlignment = Excel.Constants.xlCenter
                    xSheet.Name = "Annexure"
                End If
            End If
            xApp.Workbooks(1).Activate()
            xApp.Visible = True
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

#End Region

#Region "Control Events"
    Private Sub BtnHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnHelpProvDocNo.Click, BtnHelpPartCode.Click
        Dim strSQL As String = String.Empty
        Dim strHelp As String()
        Dim odt As New System.Data.DataTable
        Try

            With ctlHelp
                .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                .ConnectAsUser = gstrCONNECTIONUSER
                .ConnectThroughDSN = gstrCONNECTIONDSN
                .ConnectWithPWD = gstrCONNECTIONPASSWORD
            End With
            If sender Is BtnHelpProvDocNo Then
                strSQL = " select PROV_DOCNO, fromDate, TODATE, Customercode, IsAuthorized from Sales_Prov_Hdr where IsSubmitforAuthorized=1 AND Unit_Code='" + gstrUNITID + "' order by PROV_DOCNO desc"
                strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select Sales Provision Document No", 1, 0, txtProvDocNo.Text.Trim)
                If Not IsNothing(strHelp) Then
                    If strHelp.Length > 0 Then
                        txtProvDocNo.Text = strHelp(0).Trim
                        'strSQL = " SELECT SPH.PROV_DOCNO, SPH.FROMDATE, SPH.TODATE, SPH.CUSTOMERCODE, CM.CUST_NAME, SPH.KAMCODE, EMKAM.Name KAMName, SPH.EXCISECODE" & _
                        '         " , SPH.EXCISEPER, SPH.CESSCODE, SPH.CESSPER, SPH.SALESTAXCODE, SPH.SALESTAXPER, SPH.AEDCODE, SPH.AEDPER, SPH.SECESSCODE, SPH.SECESSPER" & _
                        '         " , SPH.SURCHARGECODE, SPH.SURCHARGEPER, SPH.ADDVATCODE, SPH.ADDVATPER" & _
                        '         " , SPH.EXCISECODE_CR,SPH.EXCISEPER_CR,SPH.CESSCODE_CR,SPH.CESSPER_CR,SPH.SALESTAXCODE_CR,SPH.SALESTAXPER_CR" & _
                        '         " , SPH.AEDCODE_CR,SPH.AEDPER_CR,SPH.SECESSCODE_CR,SPH.SECESSPER_CR,SPH.SURCHARGECODE_CR, SPH.SURCHARGEPER_CR,SPH.ADDVATCODE_CR,SPH.ADDVATPER_CR" & _
                        '         " , SPH.ISAUTHORIZED, SPH.IsText, SPH.TextPartCode, IM.Description" & _
                        '         " , SPH.ExciseVal, SPH.CessVal, SPH.AEDVal, SPH.SEcessVal, SPH.SurchargeVal, SPH.AddVATVal, SPH.TotalTaxableVal, SPH.TotalNetVal, SPH.SalesTaxVal" & _
                        '         " , SPH.ExciseVal_CR,SPH.CessVal_CR,SPH.AEDVal_CR,SPH.SEcessVal_CR,SPH.SurchargeVal_CR,SPH.AddVATVal_CR,SPH.TotalTaxableVal_CR,SPH.TotalNetVal_CR,SPH.SalesTaxVal_CR" & _
                        '         " , cast(SPH.CreditNoteValue as numeric(18,2)) CreditNoteValue, cast(SPH.SupplInvoiceValue as numeric(18,2)) SupplInvoiceValue, cast(SPH.RoundOff_Diff as numeric(18,2)) RoundOff_Diff" & _
                        '         " , cast(SPH.RoundOff_Diff_CR as numeric(18,2)) RoundOff_Diff_CR" & _
                        '         " , ISNULL(SPH.CRDOCNO,'') CRDOCNO, ISNULL(SPH.SUPPLINVDOCNO,'') SUPPLINVDOCNO" & _
                        strSQL = "SELECT ISNULL(SPH.PROV_DOCNO,0) PROV_DOCNO, SPH.FROMDATE, SPH.TODATE, SPH.CUSTOMERCODE, CM.CUST_NAME, SPH.KAMCODE,EMKAM.Name KAMName, ISNULL(SPH.EXCISECODE,'') EXCISECODE,ISNULL(SPH.EXCISEPER,0.0) EXCISEPER,ISNULL(SPH.CESSCODE,'') CESSCODE,ISNULL(SPH.CESSPER,0.0) CESSPER, ISNULL(SPH.SALESTAXCODE,'') SALESTAXCODE,  ISNULL(SPH.SALESTAXPER,0.0) SALESTAXPER,ISNULL(SPH.AEDCODE,'') AEDCODE,ISNULL(SPH.AEDPER,0.0) AEDPER,ISNULL(SPH.SECESSCODE,'') SECESSCODE, ISNULL(SPH.SECESSPER,0.0) SECESSPER  , ISNULL(SPH.SURCHARGECODE,'') SURCHARGECODE,  ISNULL(SPH.SURCHARGEPER,0.0) SURCHARGEPER, ISNULL(SPH.ADDVATCODE,'') ADDVATCODE, ISNULL(SPH.ADDVATPER,0.0) ADDVATPER , ISNULL(SPH.EXCISECODE_CR,'') EXCISECODE_CR,ISNULL(SPH.EXCISEPER_CR,0.0) EXCISEPER_CR, ISNULL(SPH.CESSCODE_CR,'') CESSCODE_CR, ISNULL(SPH.CESSPER_CR,0.0) CESSPER_CR,ISNULL(SPH.SALESTAXCODE_CR,'') SALESTAXCODE_CR,ISNULL(SPH.SALESTAXPER_CR,0.0) SALESTAXPER_CR, ISNULL(SPH.AEDCODE_CR,'') AEDCODE_CR,ISNULL(SPH.AEDPER_CR,0.0) AEDPER_CR,ISNULL(SPH.SECESSCODE_CR,'') SECESSCODE_CR, ISNULL(SPH.SECESSPER_CR,0.0) SECESSPER_CR,ISNULL(SPH.SURCHARGECODE_CR,'') SURCHARGECODE_CR,ISNULL(SPH.SURCHARGEPER_CR,0.0) SURCHARGEPER_CR ,ISNULL(SPH.ADDVATCODE_CR,'') ADDVATCODE_CR,ISNULL(SPH.ADDVATPER_CR,0.0) ADDVATPER_CR ,  SPH.ISAUTHORIZED, SPH.IsText, SPH.TextPartCode, IM.Description ,  ISNULL(SPH.ExciseVal,0.0) ExciseVal, ISNULL(SPH.CessVal,0.0) CessVal,ISNULL(SPH.AEDVal,0.0) AEDVal, ISNULL(SPH.SEcessVal,0.0) SEcessVal,  ISNULL(SPH.SurchargeVal,0.0) SurchargeVal, ISNULL(SPH.AddVATVal,0.0) AddVATVal, ISNULL(SPH.TotalTaxableVal,0.0) TotalTaxableVal,  ISNULL(SPH.TotalNetVal,0.0) TotalNetVal,ISNULL(SPH.SalesTaxVal,0.0) SalesTaxVal, ISNULL(SPH.ExciseVal_CR,0.0) ExciseVal_CR,ISNULL(SPH.CessVal_CR,0.0) CessVal_CR,  ISNULL(SPH.AEDVal_CR,0.0) AEDVal_CR ,ISNULL(SPH.SEcessVal_CR,0.0) SEcessVal_CR, ISNULL(SPH.SurchargeVal_CR,0.0) SurchargeVal_CR,ISNULL(SPH.AddVATVal_CR,0.0) AddVATVal_CR, ISNULL(SPH.TotalTaxableVal_CR,0.0) TotalTaxableVal_CR,ISNULL(SPH.TotalNetVal_CR,0.0) TotalNetVal_CR,ISNULL(SPH.SalesTaxVal_CR,0.0) SalesTaxVal_CR,   ISNULL(cast(SPH.CreditNoteValue as numeric(18,2)),0.0) CreditNoteValue, ISNULL(cast(SPH.SupplInvoiceValue as numeric(18,2)),0.0) SupplInvoiceValue,   isnull(cast(SPH.RoundOff_Diff as numeric(18,2)),0.0) RoundOff_Diff , ISNULL(cast(SPH.RoundOff_Diff_CR as numeric(18,2)),0.0) RoundOff_Diff_CR ,   ISNULL(SPH.CRDOCNO,'') CRDOCNO, ISNULL(SPH.SUPPLINVDOCNO,'') SUPPLINVDOCNO " & _
                                 " FROM SALES_PROV_HDR SPH left join Employee_Mst EMKAM" & _
                                 " ON SPH.KAMCODE=EMKAM.EMPLOYEE_CODE AND SPH.UNIT_CODE=EMKAM.UNIT_CODE" & _
                                 " LEFT JOIN ITEM_MST IM" & _
                                 " ON SPH.UNIT_CODE=IM.UNIT_CODE AND SPH.TEXTPARTCODE=IM.ITEM_CODE " & _
                                 " LEFT JOIN CUSTOMER_MST CM ON SPH.CUSTOMERCODE=CM.CUSTOMER_CODE AND SPH.UNIT_CODE=CM.UNIT_CODE" & _
                                 " WHERE convert(varchar(20),SPH.PROV_DOCNO)='" + txtProvDocNo.Text.Trim() + "' AND SPH.UNIT_CODE='" + gstrUNITID + "'"
                        odt = SqlConnectionclass.GetDataTable(strSQL)
                        If odt.Rows.Count > 0 Then
                            dtpFrm.Value = Convert.ToDateTime(odt.Rows(0)("FROMDATE"))
                            dtpToDt.Value = Convert.ToDateTime(odt.Rows(0)("TODATE"))
                            txtCustCode.Text = Convert.ToString(odt.Rows(0)("CUSTOMERCODE"))
                            lblCustDesc.Text = Convert.ToString(odt.Rows(0)("CUST_NAME"))
                            lblKAMName.Text = Convert.ToString(odt.Rows(0)("KAMName"))
                            txtExcise.Text = Convert.ToString(odt.Rows(0)("EXCISECODE"))
                            txtExcise_CR.Text = Convert.ToString(odt.Rows(0)("EXCISECODE_CR"))
                            txtPerExcise.Text = Convert.ToDecimal(odt.Rows(0)("EXCISEPER")).ToString("0.00")
                            txtPerExcise_CR.Text = Convert.ToDecimal(odt.Rows(0)("EXCISEPER_CR")).ToString("0.00")
                            txtCESS.Text = Convert.ToString(odt.Rows(0)("CESSCODE"))
                            txtCESS_CR.Text = Convert.ToString(odt.Rows(0)("CESSCODE_CR"))
                            txtPerCESS.Text = Convert.ToDecimal(odt.Rows(0)("CESSPER")).ToString("0.00")
                            txtPerCESS_CR.Text = Convert.ToDecimal(odt.Rows(0)("CESSPER_CR")).ToString("0.00")
                            txtSalesTaxCode.Text = Convert.ToString(odt.Rows(0)("SALESTAXCODE"))
                            txtSalesTaxCode_CR.Text = Convert.ToString(odt.Rows(0)("SALESTAXCODE_CR"))
                            txtPerSalesTax.Text = Convert.ToDecimal(odt.Rows(0)("SALESTAXPER")).ToString("0.00")
                            txtPerSalesTax_CR.Text = Convert.ToDecimal(odt.Rows(0)("SALESTAXPER_CR")).ToString("0.00")
                            txtAEDCode.Text = Convert.ToString(odt.Rows(0)("AEDCODE"))
                            txtAEDCode_CR.Text = Convert.ToString(odt.Rows(0)("AEDCODE_CR"))
                            txtPerAED.Text = Convert.ToDecimal(odt.Rows(0)("AEDPER")).ToString("0.00")
                            txtPerAED_CR.Text = Convert.ToDecimal(odt.Rows(0)("AEDPER_CR")).ToString("0.00")
                            txtSECESS.Text = Convert.ToString(odt.Rows(0)("SECESSCODE"))
                            txtSECESS_CR.Text = Convert.ToString(odt.Rows(0)("SECESSCODE_CR"))
                            txtPerSECESS.Text = Convert.ToDecimal(odt.Rows(0)("SECESSPER")).ToString("0.00")
                            txtPerSECESS_CR.Text = Convert.ToDecimal(odt.Rows(0)("SECESSPER_CR")).ToString("0.00")
                            txtSurcharge.Text = Convert.ToString(odt.Rows(0)("SURCHARGECODE"))
                            txtSurcharge_CR.Text = Convert.ToString(odt.Rows(0)("SURCHARGECODE_CR"))
                            txtPerSurcharge.Text = Convert.ToDecimal(odt.Rows(0)("SURCHARGEPER")).ToString("0.00")
                            txtPerSurcharge_CR.Text = Convert.ToDecimal(odt.Rows(0)("SURCHARGEPER_CR")).ToString("0.00")
                            txtAddVAT.Text = Convert.ToString(odt.Rows(0)("ADDVATCODE"))
                            txtAddVAT_CR.Text = Convert.ToString(odt.Rows(0)("ADDVATCODE_CR"))
                            txtPerAddVAT.Text = Convert.ToDecimal(odt.Rows(0)("ADDVATPER")).ToString("0.00")
                            txtPerAddVAT_CR.Text = Convert.ToDecimal(odt.Rows(0)("ADDVATPER_CR")).ToString("0.00")
                            gPDocNo_Authorized = Convert.ToBoolean(odt.Rows(0)("ISAUTHORIZED"))
                            txtExciseVal.Text = Convert.ToDecimal(odt.Rows(0)("ExciseVal")).ToString("0.00")
                            txtExciseVal_CR.Text = Convert.ToDecimal(odt.Rows(0)("ExciseVal_CR")).ToString("0.00")
                            txtCESSVal.Text = Convert.ToDecimal(odt.Rows(0)("CessVal")).ToString("0.00")
                            txtCESSVal_CR.Text = Convert.ToDecimal(odt.Rows(0)("CessVal_CR")).ToString("0.00")
                            txtSECESSVal.Text = Convert.ToDecimal(odt.Rows(0)("SEcessVal")).ToString("0.00")
                            txtSECESSVal_CR.Text = Convert.ToDecimal(odt.Rows(0)("SEcessVal_CR")).ToString("0.00")
                            txtAEDVal.Text = Convert.ToDecimal(odt.Rows(0)("AEDVal")).ToString("0.00")
                            txtAEDVal_CR.Text = Convert.ToDecimal(odt.Rows(0)("AEDVal_CR")).ToString("0.00")
                            txtSalesTaxVal.Text = Convert.ToDecimal(odt.Rows(0)("SalesTaxVal")).ToString("0.00")
                            txtSalesTaxVal_CR.Text = Convert.ToDecimal(odt.Rows(0)("SalesTaxVal_CR")).ToString("0.00")
                            txtSurchargeVal.Text = Convert.ToDecimal(odt.Rows(0)("SurchargeVal")).ToString("0.00")
                            txtSurchargeVal_CR.Text = Convert.ToDecimal(odt.Rows(0)("SurchargeVal_CR")).ToString("0.00")
                            txtAddVATVal.Text = Convert.ToDecimal(odt.Rows(0)("AddVATVal")).ToString("0.00")
                            txtAddVATVal_CR.Text = Convert.ToDecimal(odt.Rows(0)("AddVATVal_CR")).ToString("0.00")
                            txtTotAssVal.Text = Convert.ToString(odt.Rows(0)("TotalTaxableVal"))
                            txtTotAssVal_CR.Text = Convert.ToString(odt.Rows(0)("TotalTaxableVal_CR"))
                            txtNetVal.Text = Convert.ToDecimal(odt.Rows(0)("TotalNetVal")).ToString("0.00")
                            txtNetVal_CR.Text = Convert.ToDecimal(odt.Rows(0)("TotalNetVal_CR")).ToString("0.00")
                            txtPositiveVal.Text = Convert.ToString(odt.Rows(0)("SupplInvoiceValue"))
                            txtNegativeVal.Text = Convert.ToString(odt.Rows(0)("CreditNoteValue"))
                            txtRoundOffBy.Text = Convert.ToString(odt.Rows(0)("RoundOff_Diff"))
                            txtRoundOffBy_CR.Text = Convert.ToString(odt.Rows(0)("RoundOff_Diff_CR"))
                            txtCRDocNo_CR.Text = Convert.ToString(odt.Rows(0)("CRDOCNO"))
                            txtSuppInvNo.Text = Convert.ToString(SqlConnectionclass.ExecuteScalar("SELECT ISNULL(DOC_NO,'') FROM SUPPLEMENTARYINV_HDR(NOLOCK) WHERE UNIT_CODE='" + gstrUNITID + "' AND SALESPROV_DOCNO='" + txtProvDocNo.Text.Trim() + "'"))
                            If Convert.ToBoolean(odt.Rows(0)("IsText")) Then
                                rbText.Checked = True
                                rbPartCode.Checked = False
                                txtMultilinetxt.Text = Convert.ToString(odt.Rows(0)("TextPartCode"))
                                txtPartCode.Text = String.Empty
                                txtPartDesc.Text = String.Empty
                            Else
                                rbPartCode.Checked = True
                                rbText.Checked = False
                                txtPartCode.Text = Convert.ToString(odt.Rows(0)("TextPartCode"))
                                txtPartDesc.Text = Convert.ToString(odt.Rows(0)("Description"))
                                txtMultilinetxt.Text = String.Empty
                            End If
                            bindDispatchGrid()
                            bindRateItemfpGrid()
                            InitializeForm(2)
                        End If
                    End If
                End If
            ElseIf sender Is BtnHelpPartCode Then
                strSQL = " SELECT distinct SPR.ITEM_CODE, IM.DESCRIPTION FROM SALES_PROV_RATEWISEDETAIL SPR LEFT JOIN ITEM_MST IM" & _
                         " ON SPR.UNIT_CODE=IM.UNIT_CODE AND SPR.ITEM_CODE=IM.ITEM_CODE" & _
                         " WHERE CAST(SPR.PROV_DOCNO AS VARCHAR(20))='" + txtProvDocNo.Text.Trim() + "' AND SPR.UNIT_CODE='" + gstrUNITID + "'"
                strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select Sales Provision Document No", 1, 0, txtProvDocNo.Text.Trim)
                If Not IsNothing(strHelp) Then
                    If strHelp.Length > 0 Then
                        txtPartCode.Text = strHelp(0).Trim
                        txtPartDesc.Text = strHelp(1).Trim
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub fpSpreadRateWiseDtl_ButtonClicked(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles fpSpreadRateWiseDtl.ButtonClicked
        Try
            If e.row > 0 And e.col > 0 Then
                With fpSpreadRateWiseDtl
                    If e.col = RateWiseDtlGrid.Col_Select Then
                        .Col = e.col
                        .Row = e.row
                        If .Value = 1 Then
                            LockUnlockRateWiseGrid(e.row, e.row, False)
                        Else
                            .Col = RateWiseDtlGrid.Col_NewRate
                            .Value = ""
                            .Col = RateWiseDtlGrid.Col_Change
                            .Value = ""
                            .Col = RateWiseDtlGrid.Col_ReasonChange
                            .Value = ""
                            .Col = RateWiseDtlGrid.Col_NewInvRate
                            .Value = ""
                            .Col = RateWiseDtlGrid.Col_TotEffVal
                            .Value = ""
                            .Col = RateWiseDtlGrid.Col_CorrectionNature
                            .Value = ""
                            LockUnlockRateWiseGrid(e.row, e.row, True)
                        End If
                    ElseIf e.col = RateWiseDtlGrid.Col_ShowInv Then
                        .Col = RateWiseDtlGrid.Col_Part_Code
                        .Row = e.row
                        Dim frmObj As New FRMMKTTRN0084A
                        frmObj.gProvDocNo = txtProvDocNo.Text.Trim()
                        frmObj.gItem_Code = Convert.ToString(.Value).Trim()
                        .Col = RateWiseDtlGrid.Col_InvRate
                        frmObj.gInvoiceRate = Convert.ToString(.Value).Trim()
                        frmObj.gProvDocNo_Type = "AUTH"
                        frmObj.ShowDialog()
                    End If
                End With
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub fpSpreadRateWiseDtl_ClickEvent(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles fpSpreadRateWiseDtl.ClickEvent
        Try
            With fpSpreadRateWiseDtl
                If e.row > 0 Then
                    .Col = RateWiseDtlGrid.Col_Model_Code
                    .Row = e.row
                    txtRModelDesc.Text = Convert.ToString(SqlConnectionclass.ExecuteScalar("SELECT MODEL_DESC FROM BUDGET_MODEL_MST(NOLOCK) WHERE UNIT_CODE='" + gstrUNITID + "' AND ACTIVE=1 AND MODEL_CODE='" + Convert.ToString(.Value).Trim() + "'"))
                    .Col = RateWiseDtlGrid.Col_Part_Code
                    .Row = e.row
                    txtRPartDesc.Text = Convert.ToString(SqlConnectionclass.ExecuteScalar("SELECT [DESCRIPTION] FROM ITEM_MST(NOLOCK) WHERE UNIT_CODE='" + gstrUNITID + "' AND ITEM_CODE='" + Convert.ToString(.Value).Trim() + "'"))
                End If
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub BtnCommand_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click, BtnViewAnnex.Click, BtnAuthorize.Click, BtnReject.Click
        Dim sqlCmd As New SqlCommand
        Try
            If sender Is BtnExit Then
                Me.Close()
            ElseIf sender Is BtnAuthorize Then
                If SaveData("A") Then
                    InitializeForm(1)
                End If
            ElseIf sender Is BtnReject Then
                If SaveData("R") Then
                    InitializeForm(1)
                End If
            ElseIf sender Is BtnViewAnnex Then
                If String.IsNullOrEmpty(txtProvDocNo.Text.Trim()) Then
                    txtProvDocNo.Focus()
                    MessageBox.Show("First Select Provision Doc No.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return
                End If
                GenerateAnexure()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub rb_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbText.CheckedChanged, rbPartCode.CheckedChanged
        Try
            txtPartCode.Text = String.Empty
            txtPartDesc.Text = String.Empty
            txtMultilinetxt.Text = String.Empty
            If rbText.Checked Then
                BtnHelpPartCode.Enabled = False
                txtMultilinetxt.Enabled = True
            Else
                BtnHelpPartCode.Enabled = True
                txtMultilinetxt.Enabled = False
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txt_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPartCode.KeyDown, txtProvDocNo.KeyDown
        Try

            If e.KeyCode = Keys.F1 Then
                If BtnHelpPartCode.Enabled And sender Is txtPartCode Then
                    BtnHelp_Click(BtnHelpPartCode, New EventArgs())
                ElseIf sender Is txtProvDocNo Then
                    BtnHelp_Click(BtnHelpProvDocNo, New EventArgs())
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txt_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMultilinetxt.KeyPress, txtPartCode.KeyPress, txtProvDocNo.KeyPress
        Try
            If e.KeyChar = "'" And sender Is txtMultilinetxt Then
                e.Handled = True
            ElseIf sender Is txtPartCode Or sender Is txtProvDocNo Then
                e.Handled = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
#End Region

#Region "Upload Document"

    Private Sub AddColumnDocListGrid()
        Try
            If dgvDocList.Columns.Count = 0 Then

                Dim dgvC As New DataGridViewTextBoxColumn
                dgvC.DataPropertyName = "DocName"
                dgvC.Name = "FileName"
                dgvC.HeaderText = "File Name"
                dgvC.Width = 120
                dgvC.ReadOnly = True
                dgvDocList.Columns.Add(dgvC)

                Dim dgvID As New DataGridViewTextBoxColumn
                dgvID.DataPropertyName = "DocPath"
                dgvID.Name = "FilePath"
                dgvID.HeaderText = "FilePath"
                dgvID.Width = 250
                dgvID.ReadOnly = True
                dgvDocList.Columns.Add(dgvID)

                Dim dgvQ As New DataGridViewTextBoxColumn
                dgvQ.DataPropertyName = "DocExt"
                dgvQ.Name = "FileExt"
                dgvQ.HeaderText = "File Extension"
                dgvQ.Width = 80
                dgvQ.ReadOnly = True
                dgvQ.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                dgvDocList.Columns.Add(dgvQ)

                Dim dgvD As New DataGridViewButtonColumn
                dgvD.DataPropertyName = "Show"
                dgvD.Name = "Show"
                dgvD.HeaderText = "Show"
                dgvD.Width = 50
                dgvD.Text = "Show"
                dgvDocList.Columns.Add(dgvD)

                dgvDocList.AutoGenerateColumns = False
                dgvDocList.ColumnHeadersHeight = 35
                dgvDocList.AllowUserToResizeRows = False
                dgvDocList.AllowUserToAddRows = False
                dgvDocList.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Sub RetrieveFile(ByVal imavar1 As String, ByVal FLCODE As String, ByVal OpenFile As Boolean)
        Dim fs As FileStream ' Writes the BLOB to a file (*.bmp). 
        Dim bw As BinaryWriter ' Streams the binary data to the FileStream object. 
        Dim retval As Long ' The bytes returned from GetBytes. 
        Dim bufferSize As Integer = 100 ' The size of the BLOB buffer. 
        Dim outbyte(bufferSize - 1) As Byte ' The BLOB byte() buffer to be filled by 
        Dim startIndex As Long = 0 ' The starting position in the BLOB output. 
        Dim strTmp As Object
        Dim bWavFile() As Byte
        Dim fdt As SqlDataReader
        Try
            fdt = SqlConnectionclass.ExecuteReader("select DocData from Sales_Prov_DocList(NOLOCK) where cast(Prov_DocNo as varchar(20))='" + txtProvDocNo.Text.Trim() + "' and UNIT_CODE='" + gstrUNITID + "' and DocName='" + imavar1 + "'")
            fdt.Read()
            'strTmp = fdt.GetSqlBytes(0) ' Create a file to hold the output. 
            imavar1 = gstrUserMyDocPath + DateTime.Now.ToString("ddMMyyyyHHmmss") + Path.GetExtension(imavar1)
            fs = New FileStream(imavar1, FileMode.OpenOrCreate, FileAccess.Write)
            bw = New BinaryWriter(fs)
            startIndex = 0 ' Reset the starting byte for a new BLOB. 
            retval = fdt.GetBytes(0, startIndex, outbyte, 0, bufferSize) ' Read bytes into outbyte() and retain the number of bytes returned. 
            ' Continue reading and writing while there are bytes beyond the size of the 
            Do While retval = bufferSize
                bw.Write(outbyte)
                bw.Flush() ' Reposition the start index to the end of the last buffer and fill the buffer. 
                startIndex += bufferSize
                retval = fdt.GetBytes(0, startIndex, outbyte, 0, bufferSize)
            Loop
            Try
                bw.Write(outbyte, 0, retval - 1) ' Write the remaining buffer. 
            Catch Ex As Exception
                RaiseException(Ex)
            End Try
            bw.Flush()
            bw.Close() ' Close the output file. 
            fs.Close()

            If OpenFile = True And File.Exists(imavar1) Then
                OpenDBFile(imavar1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub OpenDBFile(ByVal ProcessPath As String)
        Dim objProcess As System.Diagnostics.Process
        Try
            objProcess = New System.Diagnostics.Process()
            objProcess.StartInfo.FileName = ProcessPath
            objProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal
            objProcess.Start()
        Catch
            MessageBox.Show("Could not start process " & ProcessPath, "Error")
        End Try
    End Sub

    Private Sub bindDocList()
        Dim strQry As String = String.Empty
        Dim dtDocTable As System.Data.DataTable
        Try
            If dgvDocList.ColumnCount <= 0 Then
                AddColumnDocListGrid()
            End If
            If dgvDocList.Rows.Count > 0 Then
                dgvDocList.DataSource = Nothing
            End If
            strQry = "select Prov_DocNo, Unit_Code, DocName, '' DocPath, DocExt, 'Show' Show from Sales_Prov_DocList where convert(varchar(20), Prov_DocNo)='" + txtProvDocNo.Text.Trim() + "' and Unit_Code='" + gstrUNITID + "'"

            dtDocTable = SqlConnectionclass.GetDataTable(strQry)
            dgvDocList.DataSource = dtDocTable
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub BtnUploadDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnUploadDoc.Click, btnDocFormExit.Click
        Try
            If sender Is BtnUploadDoc Then
                DocFrm = New Form()
                DocFrm.StartPosition = FormStartPosition.CenterParent
                DocFrm.Width = 665
                DocFrm.Height = 260
                DocFrm.Text = "Retrieve Document"
                DocFrm.Controls.Add(DocUploadPanel)
                DocUploadPanel.Visible = True
                DocUploadPanel.Dock = DockStyle.Fill
                DocFrm.MaximizeBox = False
                DocFrm.MinimizeBox = False
                DocFrm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
                DocFrm.ShowDialog()
                DocUploadPanel.Visible = False
            ElseIf sender Is btnDocFormExit Then
                If Not IsNothing(DocFrm) Then
                    DocFrm.Close()
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub dgvDocList_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDocList.CellContentClick
        Try
            If e.ColumnIndex = dgvDocList.Columns("Show").Index And e.RowIndex > -1 Then
                RetrieveFile(dgvDocList.Rows(e.RowIndex).Cells("FileName").Value, "", True)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

#End Region

    Private Sub fpSpreadRateWiseDtl_KeyDownEvent(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles fpSpreadRateWiseDtl.KeyDownEvent
        Try
            If e.keyCode = Keys.F1 Then

                If txtProvDocNo.Text <> "" Then
                    If fpSpreadRateWiseDtl.ActiveCol = RateWiseDtlGrid.Col_NewRate And fpSpreadRateWiseDtl.ActiveRow > 0 Then

                        fpSpreadRateWiseDtl.Col = RateWiseDtlGrid.Col_PriceChange
                        fpSpreadRateWiseDtl.Row = fpSpreadRateWiseDtl.ActiveRow

                        If (fpSpreadRateWiseDtl.Text = "Value") Then

                            fpSpreadRateWiseDtl.Col = RateWiseDtlGrid.Col_InvRate
                            fpSpreadRateWiseDtl.Row = fpSpreadRateWiseDtl.ActiveRow
                            rate = fpSpreadRateWiseDtl.Value
                            fpSpreadRateWiseDtl.Col = RateWiseDtlGrid.Col_Part_Code
                            fpSpreadRateWiseDtl.Row = fpSpreadRateWiseDtl.ActiveRow
                            part_code = fpSpreadRateWiseDtl.Value.Trim.ToString()

                            If (part_code.Trim().ToString() <> "") Then
                                GetPriceChangeDetails(part_code.Trim().ToString(), rate, "VIEW")
                            End If

                        End If

                    End If

                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    ' Code Added By Mayur Against issue ID 10816097 
    Private Sub GetPriceChangeDetails(ByRef partcode As String, ByRef rate As Double, ByRef mode As String)
        Try
            Dim frmObj_PCD As New FRMMKTTRN0084B
            frmObj_PCD.g_mode = mode
            frmObj_PCD.gItem_Code = partcode
            frmObj_PCD.gRate = rate
            frmObj_PCD.gcust_code = txtCustCode.Text.Trim.ToString()
            If mode = "VIEW" Then
                frmObj_PCD.g_ProvDoc_No = txtProvDocNo.Text.Trim.ToString
            End If
            frmObj_PCD.ShowDialog()
            frmObj_PCD.Dispose()
        Catch Ex As Exception
            RaiseException(Ex)
        End Try
    End Sub
    ' Code Added By Mayur Against issue ID 10816097 
End Class