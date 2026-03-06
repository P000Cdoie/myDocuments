Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports VB = Microsoft.VisualBasic
Imports Excel = Microsoft.Office.Interop.Excel


Public Class FRMMKTTRN0084
    '========================================================================================
    'COPYRIGHT          :   MOTHERSONSUMI INFOTECH & DESIGN LTD.
    'AUTHOR             :   GEETANJALI AGGARWAL
    'CREATION DATE      :   10 Nov 2014
    'DESCRIPTION        :   10688280-Sales Provision Generation
    '========================================================================================

#Region "Global variables"
    Dim dtSelItems As DataTable
    Dim dtDocTable As DataTable
    Dim dblBasicValue As Decimal
    Dim DocFrm As Form
    Dim part_code As String = String.Empty
    Dim rate As Double = 0.0
    Dim strPartNo As String = String.Empty
    Dim strBillNo As Integer = 0
    Dim strPricDt As String = String.Empty
    Dim strBilldate As String = String.Empty
    Dim strshp As Int16 = 0
    Dim stracp As Int16 = 0
    Dim strOldRate As Decimal = 0.0
    Dim strNewRate As Decimal = 0.0
    Dim i As Integer = 0
    'Dim frmPriceChangeDetails As New FRMMKTTRN0084B
    '======================RoundoffSetting=====================
    Dim blnISInsExcisable As Boolean
    'Dim blnEOUFlag As Boolean
    Dim blnISBasicRoundOff As Boolean
    Dim blnISExciseRoundOff As Boolean
    Dim blnISSalesTaxRoundOff As Boolean
    Dim blnISSurChargeTaxRoundOff As Boolean
    Dim blnAddCustMatrl As Boolean
    Dim blnInsIncSTax As Boolean
    Dim blnTotalToolCostRoundOff As Boolean
    Dim blnTCSTax As Boolean
    Dim intBasicRoundOffDecimal As Integer
    Dim intSaleTaxRoundOffDecimal As Integer
    Dim intExciseRoundOffDecimal As Integer
    Dim intSSTRoundOffDecimal As Integer
    Dim intTCSRoundOffDecimal As Integer
    Dim intToolCostRoundOffDecimal As Integer
    Dim blnECSSTax As Boolean
    Dim intECSRoundOffDecimal As Integer
    Dim blnECSSOnSaleTax As Boolean
    Dim intECSSOnSaleRoundOffDecimal As Integer
    Dim blnTurnOverTax As Boolean
    Dim intTurnOverTaxRoundOffDecimal As Integer
    Dim blnTotalInvoiceAmount As Boolean
    Dim intTotalInvoiceAmountRoundOffDecimal As Integer
    Dim blnIsSDTRoundoff As Boolean
    Dim intSDTNoofDecimal As Integer
    Dim blnSameUnitLoading As Boolean
    Dim blnServiceTax_Roundoff As Boolean
    Dim intServiceTaxRoundoff_Decimal As Integer
    Dim blnPackingRoundoff As Boolean
    Dim intPackingRoundoff_Decimal As Integer
    Dim intNOOFTRADINGINVOICEWITHOUTLOCKING As Integer
    Dim xbook As Microsoft.Office.Interop.Excel.Workbook
    Dim xSheet As Microsoft.Office.Interop.Excel.Worksheet
    Dim xApp As Microsoft.Office.Interop.Excel.Application
    'Dim intUNLOCKED_INVOICES As Integer
    '======================RoundoffSetting=====================

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
    Private Sub FRMMKTTRN0084_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Call FitToClient(Me, GrpMain, ctlFormHeader, GrpCmdBtn, 600)
            Me.MdiParent = mdifrmMain
            GrpCmdBtn.Left = GrpCmdBtn.Left + 50
            dtpToDt.MaxDate = GetServerDate()
            txtFilePath.Visible = False
            InitializeForm(1)
            Me.BringToFront()
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
                dtpFrm.Value = GetServerDate().AddMonths(-1)
                dtpToDt.Value = GetServerDate()
                lblKAMName.Text = ""
                txtEmpCode.Text = ""
                txtEmpName.Text = ""
                txtCustCode.Text = String.Empty
                lblCustDesc.Text = String.Empty
                txtCustCode.Enabled = False
                txtEmpCode.Enabled = False
                txtEmpName.Enabled = False
                txtRModelDesc.Text = String.Empty
                txtRPartDesc.Text = String.Empty
                txtPositiveVal.Text = String.Empty
                txt_marutiFile.Text = String.Empty
                BtnHelpCustCode.Enabled = False
                BtnHelpEmp.Enabled = False
                BtnHelpProvDocNo.Enabled = True
                BtnFetch.Enabled = False
                rbtn_MarutiFile.Enabled = False
                rbtn_NormalInvoice.Enabled = False
                btn_fileName.Enabled = False
                BtnSubmitforAuth.Enabled = False
                BtnUploadDoc.Enabled = False
                AddColumnDispatchDtlGrid()
                AddRateWiseGridColumn()
                AddColumnDocListGrid()
                ClearAllTaxes()
                ClearDataGridView(dgvDispatchDtl)
                ClearDataGridView(dgvDocList)
                dtDocTable = Nothing
                bindDocList()
                txtProvDocNo.Text = ""
                txtProvDocNo.Enabled = True
                dtpFrm.Enabled = False
                dtpToDt.Enabled = False
                DocUploadPanel.Visible = False
                bindDocList()
                cmdGrpSalesProv.Revert()
                cmdGrpSalesProv.Top = 10
                cmdGrpSalesProv.Left = 10
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
            ElseIf Form_Status_flag = 2 Then    'Add Mode
                txtProvDocNo.Text = ""
                txtProvDocNo.Enabled = False
                'dtpFrm.Enabled = False
                'dtpToDt.Enabled = False
                rbtn_MarutiFile.Enabled = True
                rbtn_NormalInvoice.Enabled = True
                lblKAMName.Text = ""
                txtEmpCode.Text = ""
                txtEmpName.Text = ""
                txtCustCode.Text = ""
                lblCustDesc.Text = String.Empty
                txtProvDocNo.Text = ""
                txtRModelDesc.Text = String.Empty
                txtRPartDesc.Text = String.Empty
                txtPositiveVal.Text = String.Empty
                txtEmpCode.Enabled = False
                txtEmpName.Enabled = False
                txtCustCode.Enabled = False
                BtnHelpCustCode.Enabled = False
                BtnHelpEmp.Enabled = False
                BtnHelpProvDocNo.Enabled = False
                BtnFetch.Enabled = False
                BtnSubmitforAuth.Enabled = False
                txtProvDocNo.Enabled = False
                dgvDispatchDtl.ReadOnly = False
                BtnUploadDoc.Enabled = True
                dtDocTable = Nothing
                bindDocList()
                ClearDataGridView(dgvDispatchDtl)
                clearFarGrid()
                ClearAllTaxes()
                dtpFrm.Focus()
                ''InitializeForm(2)
            ElseIf Form_Status_flag = 3 Then    'Sales Prov View Mode for not submiited Doc No
                dgvDispatchDtl.ReadOnly = True
                BtnSubmitforAuth.Enabled = True
                txtProvDocNo.Enabled = True
                dtpFrm.Enabled = False
                dtpToDt.Enabled = False
                txtCustCode.Enabled = False
                txtEmpCode.Enabled = False
                txtEmpName.Enabled = False
                txtRModelDesc.Text = String.Empty
                txtRPartDesc.Text = String.Empty
                'txtTotBasValEff.Text = String.Empty
                BtnHelpCustCode.Enabled = False
                BtnHelpEmp.Enabled = False
                BtnHelpProvDocNo.Enabled = True
                BtnFetch.Enabled = False
                rbtn_MarutiFile.Enabled = False
                rbtn_NormalInvoice.Enabled = False
                BtnUploadDoc.Enabled = True
                dtDocTable = Nothing
                bindDocList()
                cmdGrpSalesProv.Revert()
                cmdGrpSalesProv.Top = 10
                cmdGrpSalesProv.Left = 10
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True  ''Used for Delete button enable
                LockUnlockRateWiseGrid(1, fpSpreadRateWiseDtl.MaxRows, True)
            ElseIf Form_Status_flag = 4 Then    'Edit Mode
                BtnHelpCustCode.Enabled = False
                BtnHelpEmp.Enabled = False
                BtnHelpProvDocNo.Enabled = False
                BtnFetch.Enabled = True
                BtnSubmitforAuth.Enabled = False
                BtnUploadDoc.Enabled = True
                txtProvDocNo.Enabled = False
                dgvDispatchDtl.ReadOnly = False
                rbtn_MarutiFile.Enabled = False
                rbtn_NormalInvoice.Enabled = False
                LockUnlockRateWiseGrid(1, fpSpreadRateWiseDtl.MaxRows, False)
            ElseIf Form_Status_flag = 5 Then    'Sales Prov View Mode for submiited Doc No
                dgvDispatchDtl.ReadOnly = True
                BtnSubmitforAuth.Enabled = False
                txtProvDocNo.Enabled = True
                dtpFrm.Enabled = False
                dtpToDt.Enabled = False
                txtCustCode.Enabled = False
                txtEmpCode.Enabled = False
                txtEmpName.Enabled = False
                txtRModelDesc.Text = String.Empty
                txtRPartDesc.Text = String.Empty
                BtnHelpCustCode.Enabled = False
                BtnHelpEmp.Enabled = False
                BtnHelpProvDocNo.Enabled = True
                BtnFetch.Enabled = False
                rbtn_MarutiFile.Enabled = False
                rbtn_NormalInvoice.Enabled = False
                BtnUploadDoc.Enabled = True
                dtDocTable = Nothing
                bindDocList()
                LockUnlockRateWiseGrid(1, fpSpreadRateWiseDtl.MaxRows, True)
                cmdGrpSalesProv.Revert()
                cmdGrpSalesProv.Top = 10
                cmdGrpSalesProv.Left = 10
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False  ''Used for Delete button enable
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub GenerateAnexure()
        Dim strQry As String = String.Empty
        Dim odt As New System.Data.DataTable
        Try
            strQry = "  SELECT tbl.SNo, ITEM_CODE, DESCRIPTION, INVOICE_NO, ' '+convert(varchar(20), INVOICEDATE,105) INVOICEDATE, RATE, INVOICE_QTY, ACTUALINVOICEQTY, BASICAMOUNT , NEWRATE, NEWBASICAMOUNT" & _
                      " , BASICAMOUNTDIFF " & _
                      " FROM " & _
                      " ( " & _
                      "  SELECT ROW_NUMBER() OVER (PARTITION BY TMP.ITEM_CODE, RWD.RATE ORDER BY TMP.ITEM_CODE)+1 SNO, TMP.ITEM_CODE, '' DESCRIPTION, TMP.INVOICE_NO, TMP.INVOICEDATE, RWD.RATE RATE, SUM(ISNULL(TMP.INVOICE_QTY,0.0)) INVOICE_QTY" & _
                      "  , SUM(ISNULL(TMP.ACTUALINVOICEQTY,0.0)) ACTUALINVOICEQTY,  RWD.RATE*SUM(ISNULL(TMP.INVOICE_QTY,0.0)) BASICAMOUNT  , RWD.NEWRATE" & _
                      "  , RWD.NEWRATE*SUM(ISNULL(TMP.INVOICE_QTY,0.0)) NEWBASICAMOUNT, RWD.NEWRATE*SUM(ISNULL(TMP.INVOICE_QTY,0.0))- RWD.RATE*SUM(ISNULL(TMP.INVOICE_QTY,0.0)) BASICAMOUNTDIFF  " & _
                      "  FROM Sales_Prov_tmpPartDetail TMP " & _
                      "  INNER JOIN SALES_PROV_RATEWISEDETAIL(NOLOCK) RWD  " & _
                      "  ON TMP.CUSTOMER_CODE=RWD.CUSTOMER_CODE AND TMP.UNIT_CODE=RWD.UNIT_CODE AND TMP.ITEM_CODE=RWD.ITEM_CODE  AND ROUND(TMP.RATE,2,1)=RWD.RATE" & _
                      "  WHERE TMP.UNIT_CODE='" + gstrUNITID + "' AND TMP.IPADDRESS='" + gstrIpaddressWinSck + "' AND CONVERT(VARCHAR(20), RWD.PROV_DOCNO)='" + txtProvDocNo.Text.Trim() + "'  " & _
                      "  GROUP BY TMP.ITEM_CODE, TMP.INVOICE_NO, TMP.INVOICEDATE, RWD.RATE, RWD.NEWRATE  " & _
                      "  UNION ALL" & _
                      "  SELECT ROW_NUMBER() OVER (PARTITION BY RWD.ITEM_CODE,RWD.RATE ORDER BY RWD.ITEM_CODE) SNO, RWD.ITEM_CODE, IM.DESCRIPTION, '' INVOICE_NO, '' INVOICEDATE, RWD.RATE RATE, 0.0 INVOICE_QTY, 0.0 ACTUALINVOICEQTY, 0.0 BASICAMOUNT  , 0.0 NEWRATE, 0.0 NEWBASICAMOUNT" & _
                      "  , 0.0 BASICAMOUNTDIFF  " & _
                      "  FROM SALES_PROV_RATEWISEDETAIL(NOLOCK) RWD " & _
                      "  LEFT JOIN ITEM_MST IM " & _
                      "  ON RWD.UNIT_CODE=IM.UNIT_CODE AND RWD.ITEM_CODE=IM.ITEM_CODE  " & _
                      "  WHERE RWD.UNIT_CODE='" + gstrUNITID + "' AND CONVERT(VARCHAR(20), RWD.PROV_DOCNO)='" + txtProvDocNo.Text.Trim() + "'  " & _
                      "  UNION ALL" & _
                      "  SELECT row_number() over (partition by RWD.item_code,RWD.RATE order by RWD.Item_COde)+1+count(RWD.Item_Code), RWD.ITEM_CODE,'PART TOTAL' DESCRIPTION" & _
                      ", '' INVOICE_NO, '' INVOICEDATE, RWD.RATE RATE, 0.0 INVOICE_QTY, 0.0 ACTUALINVOICEQTY, 0.0 BASICAMOUNT  , 0.0 NEWRATE, 0.0 NEWBASICAMOUNT" & _
                      ", SUM(RWD.NEWRATE*ISNULL(TMP.INVOICE_QTY,0.0))-SUM(RWD.RATE*ISNULL(TMP.INVOICE_QTY,0.0)) BASICAMOUNTDIFF  " & _
                      "  FROM Sales_Prov_tmpPartDetail TMP" & _
                      "  INNER JOIN SALES_PROV_RATEWISEDETAIL(NOLOCK) RWD  " & _
                      "  ON TMP.CUSTOMER_CODE=RWD.CUSTOMER_CODE AND TMP.UNIT_CODE=RWD.UNIT_CODE AND TMP.ITEM_CODE=RWD.ITEM_CODE  and ROUND(TMP.Rate,2,1)=RWD.Rate" & _
                      "  WHERE TMP.UNIT_CODE='" + gstrUNITID + "' AND TMP.IPADDRESS='" + gstrIpaddressWinSck + "' AND CONVERT(VARCHAR(20), RWD.PROV_DOCNO)='" + txtProvDocNo.Text.Trim() + "' " & _
                      "  GROUP BY RWD.item_code, RWD.RATE " & _
                      "  UNION ALL " & _
                      "  SELECT 0 SNo,'' ITEM_CODE, 'PROVISION TOTAL' DESCRIPTION, '' INVOICE_NO, '' INVOICEDATE, 0.0 RATE, 0.0 INVOICE_QTY, 0.0 ACTUALINVOICEQTY, 0.0 BASICAMOUNT  , 0.0 NEWRATE" & _
                      "  , 0.0 NEWBASICAMOUNT, SUM(RWD.NEWRATE*ISNULL(TMP.INVOICE_QTY,0.0))-SUM(RWD.RATE*ISNULL(TMP.INVOICE_QTY,0.0)) BASICAMOUNTDIFF " & _
                      "  FROM Sales_Prov_tmpPartDetail TMP " & _
                      "  INNER JOIN SALES_PROV_RATEWISEDETAIL(NOLOCK) RWD  " & _
                      "  ON TMP.CUSTOMER_CODE=RWD.CUSTOMER_CODE AND TMP.UNIT_CODE=RWD.UNIT_CODE AND TMP.ITEM_CODE=RWD.ITEM_CODE  and ROUND(TMP.RATE,2,1)=RWD.RATE" & _
                      "  WHERE TMP.UNIT_CODE='" + gstrUNITID + "' AND TMP.IPADDRESS='" + gstrIpaddressWinSck + "' AND CONVERT(VARCHAR(20), RWD.PROV_DOCNO)='" + txtProvDocNo.Text.Trim() + "' " & _
                      ") TBL " & _
                      " ORDER BY TBL.ITEM_CODE desc, TBL.Rate, TBL.SNo "
            odt = SqlConnectionclass.GetDataTable(strQry)
            xApp = New Microsoft.Office.Interop.Excel.Application()
            xbook = xApp.Workbooks.Add
            xSheet = xbook.Worksheets(1)
            xSheet.Name = "Annexure"
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
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 7)).BorderAround(, Excel.XlBorderWeight.xlMedium)
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 7)).HorizontalAlignment = Excel.Constants.xlRight
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 8)).Interior.Color = RGB(115, 151, 253)
                            xSheet.Cells(ROW, 8) = Convert.ToString(odr("BasicAmountDiff"))
                            xSheet.Cells(ROW, 8).BorderAround(, Excel.XlBorderWeight.xlMedium)
                        ElseIf Convert.ToString(odr("DESCRIPTION")).Equals("PROVISION TOTAL") Then
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 7)).Merge()
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 7)).Value = "PROVISION TOTAL"
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 7)).Font.Bold = True
                            xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 7)).BorderAround(, Excel.XlBorderWeight.xlMedium)
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
                                xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 8)).Interior.Color = RGB(170, 160, 199)
                                xSheet.Range(xSheet.Cells(ROW, 1), xSheet.Cells(ROW, 8)).Font.Bold = True
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
                            xSheet.Range(xSheet.Cells(ROW, 4), xSheet.Cells(ROW, 8)).Merge()
                            xSheet.Range(xSheet.Cells(ROW, 4), xSheet.Cells(ROW, 8)).Value = Convert.ToString(odr("DESCRIPTION"))
                            xSheet.Range(xSheet.Cells(ROW, 4), xSheet.Cells(ROW, 8)).BorderAround(, Excel.XlBorderWeight.xlMedium)
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
                    xSheet.Cells.WrapText = True
                    xSheet.Cells.VerticalAlignment = Excel.Constants.xlCenter
                End If
            End If
            xApp.Workbooks(1).Activate()
            xApp.Visible = True
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
    Private Function SaveData() As Boolean
        Dim SqlCmd As New SqlCommand
        Dim strQuery As String = String.Empty
        Dim fs As FileStream
        Dim rawData As Object
        Dim odt As DataTable
        Dim strProvDocNo As String = String.Empty
        Try
            If ValidateSave() Then
                strQuery = " UPDATE SALES_PROV_TMPPARTDETAIL SET PRICECHANGETYPE=@PRICECHANGETYPE, NEWRATE=@NEWRATE, CHANGEEFFECT=@CHANGEEFFECT" & _
                           " , PERCENTAGECHANGE=@PERCENTAGECHANGE, CHANGEREASON=@CHANGEREASON, NATUREOFCORRECTION=@NATUREOFCORRECTION, IsSelected=1" & _
                           " WHERE CUSTOMER_CODE=@CUSTOMER_CODE AND ITEM_CODE=@ITEM_CODE AND VARMODEL=@VARMODEL AND RATE=@RATE" & _
                           " AND IPADDRESS=@IP_ADDRESS AND UNIT_CODE=@UNIT_CODE"
                SqlCmd = New SqlCommand
                With SqlCmd
                    .Connection = SqlConnectionclass.GetConnection()
                    .CommandText = strQuery
                    .CommandType = CommandType.Text
                    .CommandTimeout = 0
                End With
                With fpSpreadRateWiseDtl
                    For Count As Integer = 1 To .MaxRows
                        SqlCmd.Parameters.Clear()
                        .Row = Count
                        .Col = RateWiseDtlGrid.Col_Select
                        If .Value = True Then
                            .Col = RateWiseDtlGrid.Col_PriceChange
                            SqlCmd.Parameters.AddWithValue("@PRICECHANGETYPE", Convert.ToString(.Value))
                            .Col = RateWiseDtlGrid.Col_NewRate
                            SqlCmd.Parameters.AddWithValue("@NEWRATE", Convert.ToString(.Value))
                            .Col = RateWiseDtlGrid.Col_ChangeEff
                            SqlCmd.Parameters.AddWithValue("@CHANGEEFFECT", Convert.ToString(.Value))
                            .Col = RateWiseDtlGrid.Col_Change
                            SqlCmd.Parameters.AddWithValue("@PERCENTAGECHANGE", Convert.ToString(.Value))
                            .Col = RateWiseDtlGrid.Col_ReasonChange
                            SqlCmd.Parameters.AddWithValue("@CHANGEREASON", Convert.ToString(.Value))
                            .Col = RateWiseDtlGrid.Col_CorrectionNature
                            SqlCmd.Parameters.AddWithValue("@NATUREOFCORRECTION", .CellTag)
                            .Col = RateWiseDtlGrid.Col_Part_Code
                            SqlCmd.Parameters.AddWithValue("@ITEM_CODE", Convert.ToString(.Value))
                            .Col = RateWiseDtlGrid.Col_Model_Code
                            SqlCmd.Parameters.AddWithValue("@VARMODEL", Convert.ToString(.Value))
                            .Col = RateWiseDtlGrid.Col_InvRate
                            SqlCmd.Parameters.AddWithValue("@RATE", Convert.ToDecimal(.Value).ToString("0.00"))
                            SqlCmd.Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                            SqlCmd.Parameters.AddWithValue("@CUSTOMER_CODE", txtCustCode.Text.Trim())
                            SqlCmd.Parameters.AddWithValue("@IP_Address", gstrIpaddressWinSck)
                            SqlCmd.ExecuteNonQuery()
                        End If
                    Next
                End With
                With SqlCmd
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "USP_SALES_PROV_Generation"
                    .Transaction = .Connection.BeginTransaction
                    .Parameters.Clear()
                    If String.IsNullOrEmpty(txtProvDocNo.Text.Trim) Then
                        .Parameters.Add(New SqlParameter("@p_ProvDocNo", SqlDbType.VarChar, 20, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                        .Parameters("@p_ProvDocNo").Value = 0
                        .Parameters.AddWithValue("@p_TRANTYPE", "A")
                    Else
                        .Parameters.Add(New SqlParameter("@p_ProvDocNo", SqlDbType.VarChar, 20, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                        .Parameters("@p_ProvDocNo").Value = txtProvDocNo.Text.Trim
                        .Parameters.AddWithValue("@p_TRANTYPE", "U")
                    End If
                    .Parameters.AddWithValue("@p_FROMDATE", dtpFrm.Value.ToString("dd MMM yyyy"))
                    .Parameters.AddWithValue("@p_TODATE", dtpToDt.Value.ToString("dd MMM yyyy"))
                    .Parameters.AddWithValue("@p_PERSONTO_AUTH", txtEmpCode.Text.Trim())
                    .Parameters.AddWithValue("@p_KAMCODE", Convert.ToString(lblKAMName.Tag))

                    .Parameters.AddWithValue("@p_ExciseCode", txtExcise.Text.Trim())
                    If String.IsNullOrEmpty(txtPerExcise.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_Excise", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_Excise", txtPerExcise.Text.Trim())
                    End If
                    .Parameters.AddWithValue("@p_CessCode", txtCESS.Text.Trim())
                    If String.IsNullOrEmpty(txtPerCESS.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_Cess", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_Cess", txtPerCESS.Text.Trim())
                    End If
                    .Parameters.AddWithValue("@p_SalesTaxCode", txtSalesTaxCode.Text.Trim())
                    If String.IsNullOrEmpty(txtPerSalesTax.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_SalesTax", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_SalesTax", txtPerSalesTax.Text.Trim())
                    End If

                    .Parameters.AddWithValue("@p_AEDCode", txtAEDCode.Text.Trim())
                    If String.IsNullOrEmpty(txtPerAED.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_AED", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_AED", txtPerAED.Text.Trim())
                    End If

                    .Parameters.AddWithValue("@p_SEcessCode", txtSECESS.Text.Trim())
                    If String.IsNullOrEmpty(txtPerSECESS.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_SEcess", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_SEcess", txtPerSECESS.Text.Trim())
                    End If

                    .Parameters.AddWithValue("@p_SurchargeCode", txtSurcharge.Text.Trim())
                    If String.IsNullOrEmpty(txtPerSurcharge.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_Surcharge", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_Surcharge", txtPerSurcharge.Text.Trim())
                    End If

                    .Parameters.AddWithValue("@p_AddVATCode", txtAddVAT.Text.Trim())
                    If String.IsNullOrEmpty(txtPerAddVAT.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_AddVAT", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_AddVAT", txtPerAddVAT.Text.Trim())
                    End If

                    ' credit taxes
                    .Parameters.AddWithValue("@p_ExciseCode_CR", txtExcise_CR.Text.Trim())
                    If String.IsNullOrEmpty(txtExcisePER_CR.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_Excise_CR", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_Excise_CR", txtExcisePER_CR.Text.Trim())
                    End If

                    .Parameters.AddWithValue("@p_CessCode_CR", txtCESS_CR.Text.Trim())
                    If String.IsNullOrEmpty(txtPerCESS_CR.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_Cess_CR", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_Cess_CR", txtPerCESS_CR.Text.Trim())
                    End If

                    .Parameters.AddWithValue("@p_SalesTaxCode_CR", txtSalesTaxCode_CR.Text.Trim())
                    If String.IsNullOrEmpty(txtPerSalesTax_CR.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_SalesTax_CR", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_SalesTax_CR", txtPerSalesTax_CR.Text.Trim())
                    End If

                    .Parameters.AddWithValue("@p_AEDCode_CR", txtAEDCode_CR.Text.Trim())
                    If String.IsNullOrEmpty(txtPerAED_CR.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_AED_CR", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_AED_CR", txtPerAED_CR.Text.Trim())
                    End If

                    .Parameters.AddWithValue("@p_SEcessCode_CR", txtSECESS_CR.Text.Trim())
                    If String.IsNullOrEmpty(txtPerSECESS_CR.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_SEcess_CR", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_SEcess_CR", txtPerSECESS_CR.Text.Trim())
                    End If

                    .Parameters.AddWithValue("@p_SurchargeCode_CR", txtSurcharge_Neg.Text.Trim())
                    If String.IsNullOrEmpty(txtPerSurcharge_neg.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_Surcharge_CR", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_Surcharge_CR", txtPerSurcharge_neg.Text.Trim())
                    End If

                    .Parameters.AddWithValue("@p_AddVATCode_CR", txtAddVAT_CR.Text.Trim())
                    If String.IsNullOrEmpty(txtAddPerVAT_CR.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_AddVAT_CR", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_AddVAT_CR", txtAddPerVAT_CR.Text.Trim())
                    End If
                    ' credit taxes

                    .Parameters.AddWithValue("@p_UserId", mP_User)
                    .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                    .Parameters.AddWithValue("@p_CUSTOMERCODE", txtCustCode.Text.Trim())
                    .Parameters.AddWithValue("@p_IPAddress", gstrIpaddressWinSck)
                    .Parameters.Add(New SqlParameter("@p_ERROR", SqlDbType.VarChar, 200, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                    .ExecuteNonQuery()

                    If String.IsNullOrEmpty(Convert.ToString(SqlCmd.Parameters("@p_ERROR").Value)) Then
                        If String.IsNullOrEmpty(txtProvDocNo.Text.Trim()) Then
                            strProvDocNo = Convert.ToString(SqlCmd.Parameters("@p_ProvDocNo").Value)
                        Else
                            strProvDocNo = txtProvDocNo.Text.Trim()
                        End If
                        strQuery = "SELECT DOCNAME, CAST(0 AS BIT) ISSAVED FROM SALES_PROV_DOCLIST WHERE CAST(PROV_DOCNO AS VARCHAR(20))='" + txtProvDocNo.Text.Trim() + "' AND UNIT_CODE='" + gstrUNITID + "'"
                        odt = SqlConnectionclass.GetDataTable(strQuery)
                        If Not IsNothing(dtDocTable) Then
                            If dtDocTable.Rows.Count > 0 Then
                                strQuery = " DELETE FROM SALES_PROV_DOCLIST WHERE CAST(PROV_DOCNO AS VARCHAR(20))=@PROV_DOCNO AND UNIT_CODE=@UNIT_CODE AND DOCNAME=@DOCNAME" & _
                                           " INSERT INTO SALES_PROV_DOCLIST(PROV_DOCNO, UNIT_CODE, DOCNAME, DOCDATA, DOCEXT)" & _
                                           " SELECT @PROV_DOCNO, @UNIT_CODE, @DOCNAME, @DOCDATA, @DOCEXT "
                                .CommandText = strQuery
                                .CommandType = CommandType.Text
                                For Each odr As DataRow In dtDocTable.Rows
                                    If Not String.IsNullOrEmpty(Convert.ToString(odr("DocPath"))) Then
                                        fs = New FileStream(Convert.ToString(odr("DocPath")), FileMode.Open, FileAccess.Read)
                                        rawData = New Byte(fs.Length) {}
                                        fs.Read(rawData, 0, fs.Length)
                                        fs.Close()
                                        .Parameters.Clear()
                                        .Parameters.AddWithValue("@PROV_DOCNO", strProvDocNo)
                                        .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                                        .Parameters.AddWithValue("@DOCNAME", Convert.ToString(odr("DocName")))
                                        .Parameters.AddWithValue("@DOCDATA", rawData)
                                        .Parameters.AddWithValue("@DOCEXT", Convert.ToString(odr("DocExt")))
                                        .ExecuteNonQuery()
                                    End If
                                    If odt.Select("DocName='" + Convert.ToString(odr("DOCName")) + "'").Length > 0 Then
                                        odt.Select("DocName='" + Convert.ToString(odr("DOCName")) + "'")(0)("ISSAVED") = True
                                    End If
                                Next
                                .CommandText = " DELETE FROM SALES_PROV_DOCLIST WHERE CAST(PROV_DOCNO AS VARCHAR(20))=@PROV_DOCNO AND UNIT_CODE=@UNIT_CODE AND DOCNAME=@DOCNAME"
                                .CommandType = CommandType.Text
                                For Each odr As DataRow In odt.Select("ISSAVED=0")
                                    .Parameters.Clear()
                                    .Parameters.AddWithValue("@PROV_DOCNO", strProvDocNo)
                                    .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                                    .Parameters.AddWithValue("@DOCNAME", Convert.ToString(odr("DocName")))
                                    .ExecuteNonQuery()
                                Next
                            End If
                        ElseIf odt.Rows.Count > 0 Then
                            SqlConnectionclass.ExecuteNonQuery("DELETE FROM SALES_PROV_DOCLIST WHERE CAST(PROV_DOCNO AS VARCHAR(20))='" + strProvDocNo + "' AND UNIT_CODE='" + gstrUNITID + "'")
                        End If
                        .Transaction.Commit()
                        If String.IsNullOrEmpty(txtProvDocNo.Text.Trim()) Then
                            txtProvDocNo.Text = strProvDocNo
                            MessageBox.Show("Sale Provision Doc No " + txtProvDocNo.Text.Trim() + " created successfully.")
                        Else
                            MessageBox.Show("Record updated successfully.")
                        End If
                        Return True
                    Else
                        .Transaction.Rollback()
                        MessageBox.Show(Convert.ToString(SqlCmd.Parameters("@p_ERROR").Value), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Return False
                    End If
                End With
            Else
                Return False
            End If
        Catch ex As Exception
            If (Not IsNothing(SqlCmd.Transaction)) Then
                SqlCmd.Transaction.Rollback()
            End If
            RaiseException(ex)
        Finally
            SqlCmd.Dispose()
        End Try
        Return True
    End Function

    ''' <summary>Validate manadatory fields</summary>
    Private Function ValidateSave() As Boolean
        Dim strQry As String = String.Empty
        Dim isValid As Boolean = False
        Dim newrate As Double = 0.0
        Dim oldrate As Double = 0.0
        Try
            If String.IsNullOrEmpty(txtEmpCode.Text.Trim) Then
                MessageBox.Show("Authorization Person cannot be blank.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                BtnHelpEmp.Focus()
                Return False
            End If
            If String.IsNullOrEmpty(txtCustCode.Text.Trim) Then
                MessageBox.Show("Customer cannot be blank.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                BtnHelpCustCode.Focus()
                Return False
            End If
            With fpSpreadRateWiseDtl
                If .MaxRows = 0 Then
                    MessageBox.Show("Select atleast one Part Code.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                    Return False
                End If
                For row As Integer = 1 To .MaxRows
                    .Row = row
                    .Col = RateWiseDtlGrid.Col_Select
                    If .Value = 1 Then                      'Is row selected
                        isValid = True
                        .Col = RateWiseDtlGrid.Col_PriceChange
                        If .Value = 0 Then                      ' check for New Rate in case of Value
                            .Col = RateWiseDtlGrid.Col_NewRate
                            If .Value = "" Then
                                MessageBox.Show("Enter New Rate.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                                Return False
                            End If
                        Else                                    ' check for Change in Rate in case of Percentage change
                            .Col = RateWiseDtlGrid.Col_Change
                            If .Value = "" Then
                                MessageBox.Show("Enter Change in Rate.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                                Return False
                            End If
                        End If
                        .Col = RateWiseDtlGrid.Col_TotEffVal
                        If .Value = 0.0 Then
                            MessageBox.Show("New Rate must be different from Old rate.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            Return False
                        End If

                        .Col = RateWiseDtlGrid.Col_ReasonChange    'Check for Reason Change
                        If .Value = "" Then
                            MessageBox.Show("Enter Reason for Change.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            Return False
                        End If
                        .Col = RateWiseDtlGrid.Col_CorrectionNature     'Validate Nature of Correction
                        If .Value = "" Then
                            MessageBox.Show("Choose Nature of Correction.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            Return False
                        End If
                    End If
                Next
                .Col = RateWiseDtlGrid.Col_Select
                .Row = 1
                If Not isValid Then
                    MessageBox.Show("Select atlease one Part Code.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    .Focus()
                    Return False
                End If
            End With


            With fpSpreadRateWiseDtl
                For row As Integer = 1 To .MaxRows
                    .Row = row
                    .Col = RateWiseDtlGrid.Col_Select
                    If .Value = 1 Then
                        .Col = RateWiseDtlGrid.Col_NewRate
                        newrate = .Value

                        .Col = RateWiseDtlGrid.Col_InvRate
                        oldrate = .Value

                        If (newrate = oldrate) Then
                            MessageBox.Show("New Rate and Old rate should not be Same.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                            Return False
                        End If
                    End If
                Next
            End With

            If (Math.Abs(Convert.ToDouble(Val(txtPositiveVal.Text.Trim())))) > 0.0 Then

                If String.IsNullOrEmpty(txtExcise.Text.Trim()) Then
                    MessageBox.Show("Select Excise Code.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                    txtExcise.Focus()
                    Return False
                End If

                If String.IsNullOrEmpty(txtCESS.Text.Trim()) Then
                    MessageBox.Show("Select CESS on ED.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                    txtCESS.Focus()
                    Return False
                End If

                If String.IsNullOrEmpty(txtSalesTaxCode.Text.Trim()) Then
                    MessageBox.Show("Select Sales Tax Code.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                    txtSalesTaxCode.Focus()
                    Return False
                End If

                If String.IsNullOrEmpty(txtAEDCode.Text.Trim()) Then
                    MessageBox.Show("Select AED Code.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                    txtAEDCode.Focus()
                    Return False
                End If

                If String.IsNullOrEmpty(txtSECESS.Text.Trim()) Then
                    MessageBox.Show("Select SECESS Code.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                    txtSECESS.Focus()
                    Return False
                End If

                If String.IsNullOrEmpty(txtSurcharge.Text.Trim()) Then
                    MessageBox.Show("Select Surcharge Code.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                    txtSurcharge.Focus()
                    Return False
                End If

                If String.IsNullOrEmpty(txtAddVAT.Text.Trim()) Then
                    MessageBox.Show("Select Add VAT Code.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                    txtSurcharge.Focus()
                    Return False
                End If
            End If

            If (Math.Abs(Convert.ToDouble(Val(txtNegativeVal.Text.Trim())))) > 0.0 Then

                If String.IsNullOrEmpty(txtExcise_CR.Text.Trim()) Then
                    MessageBox.Show("Select Excise Code.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                    txtExcise_CR.Focus()
                    Return False
                End If

                'If String.IsNullOrEmpty(txtCESS_CR.Text.Trim()) Then
                '    MessageBox.Show("Select CESS on ED.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                '    txtCESS_CR.Focus()
                '    Return False
                'End If

                'If String.IsNullOrEmpty(txtSalesTaxCode_CR.Text.Trim()) Then
                '    MessageBox.Show("Select Sales Tax Code.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                '    txtSalesTaxCode_CR.Focus()
                '    Return False
                'End If

                'If String.IsNullOrEmpty(txtAEDCode_CR.Text.Trim()) Then
                '    MessageBox.Show("Select AED Code.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                '    txtAEDCode_CR.Focus()
                '    Return False
                'End If

                'If String.IsNullOrEmpty(txtSECESS_CR.Text.Trim()) Then
                '    MessageBox.Show("Select SECESS Code.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                '    txtSECESS_CR.Focus()
                '    Return False
                'End If

                'If String.IsNullOrEmpty(txtSurcharge_Neg.Text.Trim()) Then
                '    MessageBox.Show("Select Surcharge Code.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                '    txtSurcharge_Neg.Focus()
                '    Return False
                'End If

                'If String.IsNullOrEmpty(txtAddVAT_CR.Text.Trim()) Then
                '    MessageBox.Show("Select Add VAT Code.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                '    txtAddVAT_CR.Focus()
                '    Return False
                'End If

            End If

            Return True
        Catch ex As Exception
            RaiseException(ex)
        End Try
        Return True
    End Function

    ''' <summary>To Delete Sales Provision on Delete press.
    ''' </summary>
    ''' <returns>TRUE IF RECORD DELETED SUCCESSFULLY.</returns>
    Private Function DeleteData() As Boolean
        Try
            Dim sqlCmd As New SqlCommand
            With sqlCmd
                .Connection = SqlConnectionclass.GetConnection()
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .CommandText = "USP_SALES_PROV_Generation"
                .Parameters.Add(New SqlParameter("@p_ProvDocNo", SqlDbType.VarChar, 20, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                .Parameters("@p_ProvDocNo").Value = txtProvDocNo.Text.Trim
                .Parameters.Add(New SqlParameter("@p_ERROR", SqlDbType.NChar, 200, ParameterDirection.Output, True, 0, 0, "", DataRowVersion.Default, ""))
                .Parameters.AddWithValue("@p_UserId", mP_User)
                .Parameters.AddWithValue("@p_TRANTYPE", "D")
                .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                .Parameters.AddWithValue("@p_CUSTOMERCODE", txtCustCode.Text.Trim())
                .Parameters.AddWithValue("@p_IPAddress", gstrIpaddressWinSck)
                .ExecuteNonQuery()
                If Not String.IsNullOrEmpty(Convert.ToString(.Parameters("@p_ERROR").Value).Trim) Then
                    MessageBox.Show(Convert.ToString(sqlCmd.Parameters("@p_ERROR").Value).Trim, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                    Return False
                Else
                    'MessageBox.Show("Sales Provision Document No " + Convert.ToString(.Parameters("@p_ProvDocNo").Value).Trim + " deleted Successfully.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    MessageBox.Show("Provisioning deleted Successfully.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    Return True
                End If
            End With
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
        Dim odt As New DataTable
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

                        If Not cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW And Convert.ToBoolean(row("SELECT")) Then
                            LockUnlockRateWiseGrid(.Row, .Row, False)   'Unlock grid
                        Else
                            LockUnlockRateWiseGrid(.Row, .Row, True)  'lock grid
                        End If
                    Next
                    GetTotalBasicValEffect()
                    CalculateTaxes()
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

    ''' <summary>Fill data in Tmp Table.</summary>
    Private Sub FillTmpTable()
        Dim strQry As String = String.Empty
        Dim sqlCmd As New SqlCommand()
        Try
            'strQry = "DELETE FROM SALES_PROV_TMPPARTDETAIL WHERE UNIT_CODE='" + gstrUNITID + "' AND IPADDRESS='" + gstrIpaddressWinSck + "';" & _
            '       " INSERT INTO SALES_PROV_TMPPARTDETAIL( CUSTOMER_CODE, ITEM_CODE, VARMODEL, INVOICE_NO, INVOICEDATE, CUSTREFNO, CUSTREFDATE, AMENDNO, AMENDDATE" & _
            '       " , RATE, INVOICE_QTY, ACTUALINVOICEQTY , PRICECHANGETYPE, NEWRATE, CHANGEEFFECT, PERCENTAGECHANGE, CHANGEREASON" & _
            '       " , REJECTIONQTY, CDRQTY, IPADDRESS, UNIT_CODE) " & _
            '       " SELECT SCD.ACCOUNT_CODE, SD.ITEM_CODE, ISNULL(CM.VARMODEL,'') VARMODEL, SCD.DOC_NO, SCD.INVOICE_DATE, SCD.CUst_ref, COH.ORDER_DATE, SCD.Amendment_No" & _
            '       " , COH.AMENDMENT_DATE, ROUND(SD.RATE,2,1) RATE, ISNULL(SD.SALES_QUANTITY,0) INVOICE_QTY, ISNULL(SD.SALES_QUANTITY,0)-(ISNULL(GID.QUANTITY_REC,0) + ISNULL(CDR_TBL.CDR_QTY,0)) ACTUALINVOICEQTY " & _
            '       " , CAST('0' AS BIT) PRICECHANGETYPE, '' NEWRATE, CAST('0' AS BIT) CHANGEEFFECT, '' PERCENTAGECHANGE, '' CHANGEREASON" & _
            '       " , ISNULL(GID.QUANTITY_REC,0) RejectionQty, ISNULL(CDR_TBL.CDR_QTY,0) CDRQty,'" + gstrIpaddressWinSck + "' IPADDRESS, '" + gstrUNITID + "' UNITCODE " & _
            '       " FROM SALESCHALLAN_DTL(NOLOCK) SCD" & _
            '       " INNER JOIN SALES_DTL(NOLOCK) SD " & _
            '       " ON SD.DOC_NO=SCD.DOC_NO AND SD.UNIT_CODE=SCD.UNIT_CODE" & _
            '       " LEFT JOIN CUSTITEM_MST(NOLOCK) CM" & _
            '       " ON SCD.UNIT_CODE=CM.UNIT_CODE AND SD.ITEM_CODE=CM.ITEM_CODE AND SCD.ACCOUNT_CODE=CM.ACCOUNT_CODE AND SD.CUST_ITEM_CODE=CM.CUST_DRGNO" & _
            '       " LEFT JOIN GRN_INVOICE_DTL(NOLOCK) GID" & _
            '       " ON SCD.DOC_NO=GID.INVOICE_NO AND SCD.UNIT_CODE=GID.UNIT_CODE AND SD.ITEM_CODE=GID.ITEM_CODE " & _
            '       " AND GRIN_NO IN (SELECT DOC_NO FROM GRN_HDR WHERE UNIT_CODE='M11' AND DOC_CATEGORY ='E')" & _
            '       " LEFT JOIN (" & _
            '       "   SELECT ACH.DOC_NO, ACD.INVOICE, ACH.CUSTOMER_CODE, ACD.ITEM_CODE, ACD.CUST_ITEM_CODE, ACD.UNIT_CODE, SUM(ISNULL(ACD.ADJUSTED_CUMMS,0)) CDR_QTY" & _
            '       "   FROM ASN_CUMMSADJST_HDR ACH(NOLOCK) INNER JOIN ASN_CUMMSADJST_DTL(NOLOCK) ACD" & _
            '       "   ON ACH.DOC_NO=ACD.DOC_NO" & _
            '       "   AND ACH.DOC_TYPE=ACD.DOC_TYPE" & _
            '       "   AND ACH.UNIT_CODE=ACD.UNIT_CODE" & _
            '       "   WHERE(ACH.DOC_TYPE = 9997 AND LEN(AUTHORIZED_CODE) > 0)" & _
            '       "   AND INVOICE_TYPE='INV' AND ACH.UNIT_CODE='" + gstrUNITID + "' AND ACD.NATURE='SUBTRACTION'" & _
            '       "   GROUP BY ACH.DOC_NO,INVOICE, ACH.CUSTOMER_CODE, ACD.ITEM_CODE, ACD.CUST_ITEM_CODE, ACD.UNIT_CODE" & _
            '       " ) CDR_TBL" & _
            '       " ON SCD.ACCOUNT_CODE=CDR_TBL.CUSTOMER_CODE AND SCD.UNIT_CODE=CDR_TBL.UNIT_CODE AND SD.ITEM_CODE=CDR_TBL.ITEM_CODE AND SD.CUST_ITEM_CODE=CDR_TBL.CUST_ITEM_CODE" & _
            '       " AND SD.DOC_NO=CDR_TBL.INVOICE" & _
            '       " LEFT JOIN CUST_ORD_HDR(NOLOCK) COH ON SCD.ACCOUNT_CODE=COH.ACCOUNT_CODE AND SCD.UNIT_CODE=COH.UNIT_CODE AND SCD.CUST_REF=ISNULL(COH.CUST_REF,'') AND SCD.AMENDMENT_NO=ISNULL(COH.AMENDMENT_NO,'')" & _
            '       " WHERE SCD.UNIT_CODE='" + gstrUNITID + "' AND SCD.INVOICE_TYPE='INV' AND SCD.BILL_FLAG=1 AND SCD.CANCEL_FLAG=0" & _
            '       " AND SCD.ACCOUNT_CODE='" + txtCustCode.Text.Trim + "' AND SCD.INVOICE_DATE BETWEEN '" + dtpFrm.Value.ToString("dd MMM yyyy") + "' AND '" + dtpToDt.Value.ToString("dd MMM yyyy") + "'"
            With sqlCmd
                .CommandText = "USP_SALES_PROV_Generation"
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .Connection = SqlConnectionclass.GetConnection()
                .Parameters.Clear()
                .Parameters.AddWithValue("@p_FROMDATE", dtpFrm.Value.ToString("dd MMM yyyy"))
                .Parameters.AddWithValue("@p_TODATE", dtpToDt.Value.ToString("dd MMM yyyy"))
                If String.IsNullOrEmpty(txtProvDocNo.Text.Trim()) Then
                    .Parameters.AddWithValue("@p_ProvDocNo", 0.0)
                Else
                    .Parameters.AddWithValue("@p_ProvDocNo", txtProvDocNo.Text.Trim())
                End If
                .Parameters.AddWithValue("@p_CUSTOMERCODE", txtCustCode.Text.Trim())
                .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                .Parameters.AddWithValue("@p_IPAddress", gstrIpaddressWinSck)
                .Parameters.AddWithValue("@p_UserId", mP_User)
                .Parameters.AddWithValue("@p_TRANTYPE", "F")  'fill Temporary Table
                .Parameters.Add(New SqlParameter("@p_ERROR", SqlDbType.VarChar, 200, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                .ExecuteNonQuery()
                If Not String.IsNullOrEmpty(.Parameters("@p_ERROR").Value) Then
                    MessageBox.Show(Convert.ToString(sqlCmd.Parameters("@p_ERROR").Value), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                End If
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ''' <summary>Clear data in Tmp Table.</summary>
    Private Sub ClearTmpTable()
        Dim strQry As String = String.Empty
        Try
            strQry = "DELETE FROM SALES_PROV_TMPPARTDETAIL WHERE UNIT_CODE='" + gstrUNITID + "' AND IPADDRESS='" + gstrIpaddressWinSck + "';"
            SqlConnectionclass.ExecuteNonQuery(strQry)
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ''' <summary>Fill data in Tmp Table.</summary>
    Private Function GetRateWisePartDetail() As DataTable
        Dim strQry As String = String.Empty
        Dim odt As New DataTable
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
                    " FROM SALES_PROV_TMPPARTDETAIL TMP LEFT JOIN SALES_PROV_RATEWISEDETAIL SRQ" & _
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
                If cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    .Col = RateWiseDtlGrid.Col_Select
                    .Col2 = RateWiseDtlGrid.Col_CorrectionNature
                    .Row = row1
                    .Row2 = row2
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                Else
                    .Col = RateWiseDtlGrid.Col_Part_Code
                    .Col2 = RateWiseDtlGrid.Col_CorrectionNature
                    .Row = row1
                    .Row2 = row2
                    .BlockMode = True
                    .Lock = True
                    .Col = RateWiseDtlGrid.Col_Select
                    .Col2 = RateWiseDtlGrid.Col_Select
                    .Lock = False
                    .BlockMode = False

                End If

                If Not status And Not cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then  'Unlock Grid
                    For i As Integer = row1 To row2
                        .Col = RateWiseDtlGrid.Col_PriceChange
                        .Row = i
                        val = .Value
                        .Row2 = i
                        If val = 0 Then
                            .Col = RateWiseDtlGrid.Col_Change
                            .Value = ""
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                            .Col = RateWiseDtlGrid.Col_NewInvRate
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                            .Value = ""
                        Else
                            .Col = RateWiseDtlGrid.Col_NewRate
                            .Value = ""
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        End If
                        .BlockMode = True
                        If rbtn_NormalInvoice.Checked = True Then
                            .Col = RateWiseDtlGrid.Col_PriceChange
                            .Col2 = RateWiseDtlGrid.Col_PriceChange
                            .Lock = False
                        End If
                        If rbtn_MarutiFile.Checked = True Then
                            .Col = RateWiseDtlGrid.Col_PriceChange
                            .Col2 = RateWiseDtlGrid.Col_PriceChange
                            .Lock = True
                        End If
                        .Col = RateWiseDtlGrid.Col_NewRate
                        .Col2 = RateWiseDtlGrid.Col_NewRate
                        '.Lock = Not (val = 0)  ' Mayur
                        .Lock = True  ' Mayur
                        .Col = RateWiseDtlGrid.Col_ChangeEff
                        .Col2 = RateWiseDtlGrid.Col_ChangeEff
                        .Lock = (val = 0)
                        .Col = RateWiseDtlGrid.Col_Change
                        .Col2 = RateWiseDtlGrid.Col_Change
                        .Lock = (val = 0)
                        .Col = RateWiseDtlGrid.Col_ReasonChange
                        .Col2 = RateWiseDtlGrid.Col_ReasonChange
                        .Lock = False
                        .BlockMode = False
                    Next
                End If
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

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

    Private Sub GetTotalBasicValEffect()
        Dim TotBaiscValPos As Double = 0.0
        Dim TotBaiscValNeg As Double = 0.0
        Try
            With fpSpreadRateWiseDtl
                For i As Integer = 1 To .MaxRows
                    .Row = i
                    .Col = RateWiseDtlGrid.Col_Select
                    If .Value = 1 Then
                        .Col = RateWiseDtlGrid.Col_TotEffVal
                        If .Value = "" Then
                            Continue For
                        ElseIf Convert.ToDouble(.Value) < 0 Then
                            TotBaiscValNeg = TotBaiscValNeg + Convert.ToDouble(.Value)
                        Else
                            TotBaiscValPos = TotBaiscValPos + Convert.ToDouble(.Value)
                        End If
                    End If
                Next
                txtPositiveVal.Text = TotBaiscValPos.ToString("0.00")
                txtNegativeVal.Text = TotBaiscValNeg.ToString("0.00")
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub ClearAllTaxes()
        Try
            txtExcise.Text = ""
            txtPerExcise.Text = ""
            txtCESS.Text = ""
            txtPerCESS.Text = ""
            txtSalesTaxCode.Text = ""
            txtPerSalesTax.Text = ""
            txtAEDCode.Text = ""
            txtPerAED.Text = ""
            txtSECESS.Text = ""
            txtPerSECESS.Text = ""
            txtSurcharge.Text = ""
            txtPerSurcharge.Text = ""
            txtAddVAT.Text = ""
            txtPerAddVAT.Text = ""
            txtExciseVal.Text = ""
            txtCESSVal.Text = ""
            txtSalesTaxVal.Text = ""
            txtAEDVal.Text = ""
            txtSECESSVal.Text = ""
            txtSurchargeVal.Text = ""
            txtAddVATVal.Text = ""
            txtTotAssVal.Text = ""
            txtNetVal.Text = ""
            txtRoundOffBy.Text = ""

            txtExcise_CR.Text = ""
            txtExcisePER_CR.Text = ""
            txtExciseVal_CR.Text = ""
            txtAEDCode_CR.Text = ""
            txtPerAED_CR.Text = ""
            txtAEDVal_CR.Text = ""
            txtCESS_CR.Text = ""
            txtPerCESS_CR.Text = ""
            txtCESSVal_CR.Text = ""
            txtSECESS_CR.Text = ""
            txtPerSECESS_CR.Text = ""
            txtSECESSVal_CR.Text = ""
            txtSalesTaxCode_CR.Text = ""
            txtPerSalesTax_CR.Text = ""
            txtSalesTaxVal_CR.Text = ""
            txtAddVAT_CR.Text = ""
            txtAddPerVAT_CR.Text = ""
            txtAddVATVal_CR.Text = ""
            txtSurcharge_Neg.Text = ""
            txtPerSurcharge_neg.Text = ""
            txtSurchargeVal_Neg.Text = ""
            txtTotAssVal_Neg.Text = ""
            txtNetVal_neg.Text = ""
            txtRoundOffBy_Neg.Text = ""

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub SubGetRoundoffConfig()
        Dim Sqlcmd As New SqlCommand
        Dim ObjVal As Object = Nothing
        Dim DataRd As SqlDataReader
        Try

            Sqlcmd.CommandTimeout = 0
            Sqlcmd.Connection = SqlConnectionclass.GetConnection
            Sqlcmd.CommandType = CommandType.Text

            Sqlcmd.CommandText = "SELECT InsExc_Excise,CustSupp_Inc,EOU_Flag, Basic_Roundoff, Basic_Roundoff_decimal, SalesTax_Roundoff, SalesTax_Roundoff_decimal, Excise_Roundoff, Excise_Roundoff_decimal, "
            Sqlcmd.CommandText = Sqlcmd.CommandText + " SST_Roundoff, SST_Roundoff_decimal, InsInc_SalesTax, TCSTax_Roundoff, TCSTax_Roundoff_decimal, TotalToolCostRoundoff, TotalToolCostRoundoff_Decimal, ECESS_Roundoff, ECESSRoundoff_Decimal, ECESSOnSaleTax_Roundoff, ECESSOnSaleTaxRoundOff_Decimal, "
            Sqlcmd.CommandText = Sqlcmd.CommandText + " TurnOverTax_RoundOff, TurnOverTaxRoundOff_Decimal, TotalInvoiceAmount_RoundOff,TotalInvoiceAmountRoundOff_Decimal, SDTRoundOff, SDTRoundOff_Decimal,SameUnitLoading,ServiceTax_Roundoff,ServiceTaxRoundoff_Decimal=isnull(ServiceTaxRoundoff_Decimal,0),Packing_Roundoff,PackingRoundoff_Decimal=isnull(PackingRoundoff_Decimal,0) FROM Sales_Parameter WHERE UNIT_CODE='" + gstrUNITID + "'"
            DataRd = Sqlcmd.ExecuteReader()
            If DataRd.HasRows Then
                DataRd.Read()
                blnISInsExcisable = DataRd("InsExc_Excise")
                blnISBasicRoundOff = DataRd("Basic_Roundoff")
                blnISExciseRoundOff = DataRd("Excise_Roundoff")
                blnISSalesTaxRoundOff = DataRd("SalesTax_Roundoff")
                blnISSurChargeTaxRoundOff = DataRd("SST_Roundoff")
                blnAddCustMatrl = DataRd("CustSupp_Inc")
                blnInsIncSTax = DataRd("InsInc_SalesTax")
                blnTotalToolCostRoundOff = DataRd("TotalToolCostRoundoff")
                blnTCSTax = DataRd("TCSTax_Roundoff")
                intBasicRoundOffDecimal = DataRd("Basic_Roundoff_decimal")
                intSaleTaxRoundOffDecimal = DataRd("SalesTax_Roundoff_decimal")
                intExciseRoundOffDecimal = DataRd("Excise_Roundoff_decimal")
                intSSTRoundOffDecimal = DataRd("SST_Roundoff_decimal")
                intTCSRoundOffDecimal = DataRd("TCSTax_Roundoff_decimal")
                intToolCostRoundOffDecimal = DataRd("TotalToolCostRoundoff_decimal")
                blnECSSTax = DataRd("ECESS_Roundoff")
                intECSRoundOffDecimal = DataRd("ECESSRoundoff_Decimal")
                blnECSSOnSaleTax = DataRd("ECESSOnSaleTax_Roundoff")
                intECSSOnSaleRoundOffDecimal = DataRd("ECESSOnSaleTaxRoundOff_Decimal")
                blnTurnOverTax = DataRd("TurnOverTax_RoundOff")
                intTurnOverTaxRoundOffDecimal = DataRd("TurnOverTaxRoundOff_Decimal")
                blnTotalInvoiceAmount = DataRd("TotalInvoiceAmount_RoundOff")
                intTotalInvoiceAmountRoundOffDecimal = DataRd("TotalInvoiceAmountRoundOff_Decimal")
                blnIsSDTRoundoff = DataRd("SDTRoundOff")
                intSDTNoofDecimal = DataRd("SDTRoundOff_Decimal")
                blnSameUnitLoading = DataRd("SameUnitLoading")
                blnServiceTax_Roundoff = DataRd("ServiceTax_Roundoff")
                intServiceTaxRoundoff_Decimal = DataRd("ServiceTaxRoundoff_Decimal")
                blnPackingRoundoff = DataRd("Packing_Roundoff")
                intPackingRoundoff_Decimal = DataRd("PackingRoundoff_Decimal")

            Else
                MessageBox.Show("No Data Define In Sales_Parameter Table", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Sub
            End If
            If DataRd.IsClosed = False Then DataRd.Close()
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            If IsNothing(Sqlcmd.Connection) = False Then
                If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
                Sqlcmd.Connection.Dispose()
            End If
            If IsNothing(Sqlcmd) = False Then
                Sqlcmd.Dispose()
            End If
        End Try

    End Sub

    Private Sub CalculateTaxes()
        Try
            SubGetRoundoffConfig()

            Dim positive_value As Double = 0.0
            Dim negative_value As Double = 0.0

            positive_value = Math.Abs(Convert.ToDouble(Val(txtPositiveVal.Text.Trim())))
            negative_value = Math.Abs(Convert.ToDouble(Val(txtNegativeVal.Text.Trim())))

            If positive_value >= 0.0 Then
                'for Basic Value
                If blnISBasicRoundOff Then
                    dblBasicValue = Math.Round(Val(txtPositiveVal.Text.Trim()), 0).ToString("0.00")
                Else
                    dblBasicValue = Math.Round(Val(txtPositiveVal.Text.Trim()), intBasicRoundOffDecimal).ToString("0.00")
                End If
                'for Excise Value
                txtExciseVal.Text = (Val(txtPositiveVal.Text.Trim()) * (Val(txtPerExcise.Text.Trim()) / 100.0)).ToString("0.00")
                If blnISExciseRoundOff Then
                    txtExciseVal.Text = Math.Round(Val(txtExciseVal.Text.Trim()), 0).ToString("0.00")
                Else
                    txtExciseVal.Text = Math.Round(Val(txtExciseVal.Text.Trim()), intExciseRoundOffDecimal)
                End If
                'for AED, ECESS and SECESS Value
                txtAEDVal.Text = (Val(txtPositiveVal.Text.Trim()) * (Val(txtPerAED.Text.Trim()) / 100.0))
                txtCESSVal.Text = (Val(txtExciseVal.Text.Trim()) * (Val(txtPerCESS.Text.Trim()) / 100.0))
                txtSECESSVal.Text = (Val(txtExciseVal.Text.Trim()) * (Val(txtPerSECESS.Text) / 100.0))
                If blnECSSTax Then
                    txtCESSVal.Text = Math.Round(Val(txtCESSVal.Text.Trim()), 0)
                    txtSECESSVal.Text = Math.Round(Val(txtSECESSVal.Text.Trim()), 0)
                Else
                    txtCESSVal.Text = Math.Round(Val(txtCESSVal.Text.Trim()), intECSRoundOffDecimal)
                    txtSECESSVal.Text = Math.Round(Val(txtSECESSVal.Text.Trim()), intECSRoundOffDecimal)
                End If

                'for Total Assesible Value
                txtTotAssVal.Text = Val(txtPositiveVal.Text.Trim()) + Val(txtExciseVal.Text.Trim()) + Val(txtCESSVal.Text.Trim()) + Val(txtSECESSVal.Text.Trim()) + Val(txtAEDVal.Text.Trim())
                txtTotAssVal.Text = Math.Round(Val(txtTotAssVal.Text), 4)
                'for Sales Tax, Add VAT 
                txtSalesTaxVal.Text = (Val(txtTotAssVal.Text.Trim()) * (Val(txtPerSalesTax.Text.Trim()) / 100.0))
                txtAddVATVal.Text = (Val(txtTotAssVal.Text.Trim()) * (Val(txtPerAddVAT.Text.Trim()) / 100.0))
                If blnISSalesTaxRoundOff Then
                    txtSalesTaxVal.Text = Math.Round(Val(txtSalesTaxVal.Text.Trim()), 0)
                    txtAddVATVal.Text = Math.Round(Val(txtAddVATVal.Text.Trim()), 0)
                Else
                    txtSalesTaxVal.Text = Math.Round(Val(txtSalesTaxVal.Text.Trim()), intSaleTaxRoundOffDecimal)
                    txtAddVATVal.Text = Math.Round(Val(txtAddVATVal.Text.Trim()), intSaleTaxRoundOffDecimal)
                End If
                'for Surcharge Value
                txtSurchargeVal.Text = (Val(txtSalesTaxVal.Text.Trim()) * (Val(txtPerSurcharge.Text.Trim()) / 100.0))
                If blnISSurChargeTaxRoundOff Then
                    txtSurchargeVal.Text = Math.Round(Val(txtSurchargeVal.Text.Trim()), 0)
                Else
                    txtSurchargeVal.Text = Math.Round(Val(txtSurchargeVal.Text.Trim()), intSSTRoundOffDecimal)
                End If
                'for Net Value
                txtNetVal.Text = Val(txtTotAssVal.Text.Trim()) + Val(txtSalesTaxVal.Text.Trim()) + Val(txtAddVATVal.Text.Trim()) + Val(txtSurchargeVal.Text.Trim())
                txtNetVal.Tag = Val(txtTotAssVal.Text.Trim()) + Val(txtSalesTaxVal.Text.Trim()) + Val(txtAddVATVal.Text.Trim()) + Val(txtSurchargeVal.Text.Trim())
                If blnTotalInvoiceAmount Then
                    txtNetVal.Text = Math.Round(Val(txtNetVal.Text.Trim()), 0)
                Else
                    txtNetVal.Text = Math.Round(Val(txtNetVal.Text.Trim()), intTotalInvoiceAmountRoundOffDecimal)
                End If
                'for Round figure difference
                txtRoundOffBy.Text = (Val(Convert.ToString(txtNetVal.Tag)) - Val(txtNetVal.Text)).ToString("0.00")
            End If

            If negative_value >= 0.0 Then
                'for Basic Value
                If blnISBasicRoundOff Then
                    dblBasicValue = Math.Round(negative_value, 0).ToString("0.00")
                Else
                    dblBasicValue = Math.Round(negative_value, intBasicRoundOffDecimal).ToString("0.00")
                End If
                'for Excise Value
                txtExciseVal_CR.Text = (negative_value * (Val(txtExcisePER_CR.Text.Trim()) / 100.0)).ToString("0.00")
                If blnISExciseRoundOff Then
                    txtExciseVal_CR.Text = Math.Round(Val(txtExciseVal_CR.Text.Trim()), 0).ToString("0.00")
                Else
                    txtExciseVal_CR.Text = Math.Round(Val(txtExciseVal_CR.Text.Trim()), intExciseRoundOffDecimal)
                End If
                'for AED, ECESS and SECESS Value    
                txtAEDVal_CR.Text = (negative_value * (Val(txtPerAED_CR.Text.Trim()) / 100.0))
                txtCESSVal_CR.Text = (Val(txtExciseVal_CR.Text.Trim()) * (Val(txtPerCESS_CR.Text.Trim()) / 100.0))
                txtSECESSVal_CR.Text = (Val(txtExciseVal_CR.Text.Trim()) * (Val(txtPerSECESS_CR.Text) / 100.0))

                If blnECSSTax Then
                    txtCESSVal_CR.Text = Math.Round(Val(txtCESSVal_CR.Text.Trim()), 0)
                    txtSECESSVal_CR.Text = Math.Round(Val(txtSECESSVal_CR.Text.Trim()), 0)
                Else
                    txtCESSVal_CR.Text = Math.Round(Val(txtCESSVal_CR.Text.Trim()), intECSRoundOffDecimal)
                    txtSECESSVal_CR.Text = Math.Round(Val(txtSECESSVal_CR.Text.Trim()), intECSRoundOffDecimal)
                End If

                'for Total Assesible Value
                txtTotAssVal_Neg.Text = negative_value + Val(txtExciseVal_CR.Text.Trim()) + Val(txtCESSVal_CR.Text.Trim()) + Val(txtSECESSVal_CR.Text.Trim()) + Val(txtAEDVal_CR.Text.Trim())
                txtTotAssVal_Neg.Text = Math.Round(Val(txtTotAssVal_Neg.Text), 4)
                'for Sales Tax, Add VAT 
                txtSalesTaxVal_CR.Text = (Val(txtTotAssVal_Neg.Text.Trim()) * (Val(txtPerSalesTax_CR.Text.Trim()) / 100.0))
                txtAddVATVal_CR.Text = (Val(txtTotAssVal_Neg.Text.Trim()) * (Val(txtAddPerVAT_CR.Text.Trim()) / 100.0))
                If blnISSalesTaxRoundOff Then
                    txtSalesTaxVal_CR.Text = Math.Round(Val(txtSalesTaxVal_CR.Text.Trim()), 0)
                    txtAddVATVal_CR.Text = Math.Round(Val(txtAddVATVal_CR.Text.Trim()), 0)
                Else
                    txtSalesTaxVal_CR.Text = Math.Round(Val(txtSalesTaxVal_CR.Text.Trim()), intSaleTaxRoundOffDecimal)
                    txtAddVATVal_CR.Text = Math.Round(Val(txtAddVATVal_CR.Text.Trim()), intSaleTaxRoundOffDecimal)
                End If
                'for Surcharge Value
                txtSurchargeVal_Neg.Text = (Val(txtSalesTaxVal_CR.Text.Trim()) * (Val(txtPerSurcharge_neg.Text.Trim()) / 100.0))
                If blnISSurChargeTaxRoundOff Then
                    txtSurchargeVal_Neg.Text = Math.Round(Val(txtSurchargeVal_Neg.Text.Trim()), 0)
                Else
                    txtSurchargeVal_Neg.Text = Math.Round(Val(txtSurchargeVal_Neg.Text.Trim()), intSSTRoundOffDecimal)
                End If
                'for Net Value
                txtNetVal_neg.Text = Val(txtTotAssVal_Neg.Text.Trim()) + Val(txtSalesTaxVal_CR.Text.Trim()) + Val(txtAddVATVal_CR.Text.Trim()) + Val(txtSurchargeVal_Neg.Text.Trim())
                txtNetVal_neg.Tag = Val(txtTotAssVal_Neg.Text.Trim()) + Val(txtSalesTaxVal_CR.Text.Trim()) + Val(txtAddVATVal_CR.Text.Trim()) + Val(txtSurchargeVal_Neg.Text.Trim())
                If blnTotalInvoiceAmount Then
                    txtNetVal_neg.Text = Math.Round(Val(txtNetVal_neg.Text.Trim()), 0)
                Else
                    txtNetVal_neg.Text = Math.Round(Val(txtNetVal_neg.Text.Trim()), intTotalInvoiceAmountRoundOffDecimal)
                End If
                'for Round figure difference
                txtRoundOffBy_Neg.Text = (Val(Convert.ToString(txtNetVal_neg.Tag)) - Val(txtNetVal_neg.Text)).ToString("0.00")
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub AddColumnDocListGrid()
        Try
            If dgvDocList.Columns.Count = 0 Then

                Dim dgvC As New DataGridViewTextBoxColumn
                dgvC.DataPropertyName = "DocName"
                dgvC.Name = "FileName"
                dgvC.HeaderText = "File Name"
                dgvC.Width = 100
                dgvC.ReadOnly = True
                dgvDocList.Columns.Add(dgvC)

                Dim dgvID As New DataGridViewTextBoxColumn
                dgvID.DataPropertyName = "DocPath"
                dgvID.Name = "FilePath"
                dgvID.HeaderText = "FilePath"
                dgvID.Width = 200
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

                Dim dgvMD As New DataGridViewButtonColumn
                dgvMD.DataPropertyName = "Remove"
                dgvMD.Name = "Remove"
                dgvMD.HeaderText = "Remove"
                dgvMD.Width = 50
                dgvMD.Text = "Remove"
                'dgvMD.UseColumnTextForButtonValue = True
                dgvDocList.Columns.Add(dgvMD)

                dgvDocList.AutoGenerateColumns = False
                'dgvDocList.RowsDefaultCellStyle.BackColor = Color.FromArgb(255, 211, 168)
                'dgvDocList.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(174, 87, 0)
                dgvDocList.ColumnHeadersHeight = 35
                dgvDocList.AllowUserToResizeRows = False
                dgvDocList.AllowUserToAddRows = False
                dgvDocList.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

#End Region

#Region "Control Events"
    Private Sub cmdGrpSalesProv_ButtonClick(ByVal Sender As System.Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles cmdGrpSalesProv.ButtonClick
        Try
            Select Case e.Button
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                    InitializeForm(2)
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = True
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = False
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
                    btn_MarutiUpload.Enabled = False
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                    If rbtn_NormalInvoice.Checked = True Then
                        If SaveData() Then
                            cmdGrpSalesProv.Revert()
                            cmdGrpSalesProv.Top = 10
                            cmdGrpSalesProv.Left = 10
                            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True '' At save time for Delete enable
                            InitializeForm(3)
                        End If
                    End If
                    If rbtn_MarutiFile.Checked = True Then
                        If SaveFileData() Then
                            cmdGrpSalesProv.Revert()
                            cmdGrpSalesProv.Top = 10
                            cmdGrpSalesProv.Left = 10
                            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True '' At save time for Delete enable
                            InitializeForm(3)
                            getdetaildata(txtProvDocNo.Text.Trim())
                        End If
                    End If
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                    InitializeForm(1)
                    cmdGrpSalesProv.Revert()
                    cmdGrpSalesProv.Top = 10
                    cmdGrpSalesProv.Left = 10
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = False
                    rbtn_MarutiFile.Checked = False
                    rbtn_NormalInvoice.Checked = False
                    dtpFrm.Enabled = True
                    dtpToDt.Enabled = True
                    btn_MarutiUpload.Enabled = True
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                    InitializeForm(4)
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = True
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = False
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
                    btn_MarutiUpload.Enabled = False
                    If rbtn_MarutiFile.Checked = True Then
                        btn_fileName.Enabled = False
                        BtnFetch.Enabled = False
                    End If
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
                    If MsgBox("Are You Sure To Delete selected entry?", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                        If DeleteData() = True Then
                            InitializeForm(1)
                            cmdGrpSalesProv.Revert()
                            cmdGrpSalesProv.Top = 10
                            cmdGrpSalesProv.Left = 10
                            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = False
                            rbtn_MarutiFile.Checked = False
                            rbtn_NormalInvoice.Checked = False
                        End If
                    End If
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
                    GenerateAnexure()
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                    Me.Close()
            End Select
        Catch ex As Exception
            RaiseException(ex)
        Finally

        End Try
    End Sub
    Private Sub BtnHelpProvDocNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnHelpProvDocNo.Click
        Dim strSQL As String = String.Empty
        Dim strHelp As String()
        Dim odt As DataTable = New DataTable
        Try
            strSQL = " select PROV_DOCNO, fromDate, TODATE, Customercode, IsSubmitforAuthorized,INV_TYPE from Sales_Prov_Hdr where Unit_Code='" + gstrUNITID + "' order by PROV_DOCNO desc"
            With ctlHelp
                .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                .ConnectAsUser = gstrCONNECTIONUSER
                .ConnectThroughDSN = gstrCONNECTIONDSN
                .ConnectWithPWD = gstrCONNECTIONPASSWORD
            End With
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select RFQ", 1, 0, txtProvDocNo.Text.Trim)
            If Not IsNothing(strHelp) Then
                If strHelp.Length > 0 Then
                    txtProvDocNo.Text = strHelp(0).Trim
                    'strSQL = " SELECT SPH.PROV_DOCNO, SPH.FROMDATE, SPH.TODATE, SPH.CUSTOMERCODE, CM.CUST_NAME, SPH.KAMCODE, EMKAM.Name KAMName, SPH.PERSONTO_AUTH, EM.Name PERSONTO_AUTHName" & _
                    '         " , SPH.EXCISECODE, SPH.EXCISEPER, SPH.CESSCODE, SPH.CESSPER, SPH.SALESTAXCODE, SPH.SALESTAXPER, SPH.AEDCODE, SPH.AEDPER, SPH.SECESSCODE, SPH.SECESSPER" & _
                    '         " , SPH.SURCHARGECODE, SPH.SURCHARGEPER, SPH.ADDVATCODE, SPH.ADDVATPER,SPH.EXCISECODE_CR,SPH.EXCISEPER_CR,SPH.CESSCODE_CR,SPH.CESSPER_CR " & _
                    '         " , SPH.SALESTAXCODE_CR,SPH.SALESTAXPER_CR,SPH.AEDCODE_CR,SPH.AEDPER_CR,SPH.SECESSCODE_CR,SPH.SECESSPER_CR,SPH.SURCHARGECODE_CR,SPH.SURCHARGEPER_CR,SPH.ADDVATCODE_CR,SPH.ADDVATPER_CR,SPH.IsSubmitforAuthorized " & _
                    '         " , SPH.ExciseVal, SPH.CessVal, SPH.AEDVal, SPH.SEcessVal, SPH.SurchargeVal, SPH.AddVATVal, SPH.TotalTaxableVal, SPH.TotalNetVal, SPH.SalesTaxVal" & _
                    '         " , SPH.ExciseVal_CR,SPH.CessVal_CR,SPH.AEDVal_CR,SPH.SEcessVal_CR,SPH.SurchargeVal_CR,SPH.AddVATVal_CR,SPH.TotalTaxableVal_CR,SPH.TotalNetVal_CR,SPH.SalesTaxVal_CR" & _
                    '         " , cast(SPH.CreditNoteValue as numeric(18,2)) CreditNoteValue, cast(SPH.SupplInvoiceValue as numeric(18,2)) SupplInvoiceValue, cast(SPH.RoundOff_Diff as numeric(18,2)) RoundOff_Diff , cast(SPH.RoundOff_Diff_CR as numeric(18,2)) RoundOff_Diff_CR" & _
                    '         " FROM SALES_PROV_HDR SPH left join Employee_Mst EMKAM" & _
                    '         " on SPH.KAMCode=EMKAM.employee_code and SPH.Unit_code=EMKAM.Unit_Code" & _
                    '         " left join Employee_Mst EM" & _
                    '         " on SPH.PERSONTO_AUTH=EM.employee_code and SPH.Unit_code=EM.Unit_Code" & _
                    '         " LEFT JOIN CUSTOMER_MST CM ON SPH.CUSTOMERCODE=CM.CUSTOMER_CODE AND SPH.UNIT_CODE=CM.UNIT_CODE" & _
                    '         " WHERE convert(varchar(20),SPH.PROV_DOCNO)='" + txtProvDocNo.Text.Trim() + "' AND SPH.UNIT_CODE='" + gstrUNITID + "'"
                    strSQL = " SELECT SPH.PROV_DOCNO, SPH.FROMDATE, SPH.TODATE, SPH.CUSTOMERCODE, CM.CUST_NAME, SPH.KAMCODE, EMKAM.Name KAMName, SPH.PERSONTO_AUTH, EM.Name PERSONTO_AUTHName" & _
                             " , ISNULL(SPH.EXCISECODE,'') EXCISECODE, ISNULL(SPH.EXCISEPER,0.0) EXCISEPER,ISNULL(SPH.CESSCODE,'') CESSCODE, ISNULL(SPH.CESSPER,0.0) CESSPER, ISNULL(SPH.SALESTAXCODE,'') SALESTAXCODE, ISNULL(SPH.SALESTAXPER,0.0) SALESTAXPER, ISNULL(SPH.AEDCODE,'') AEDCODE , ISNULL(SPH.AEDPER,0.0) AEDPER, ISNULL(SPH.SECESSCODE,'') SECESSCODE, ISNULL(SPH.SECESSPER,0.0) SECESSPER " & _
                             " , ISNULL(SPH.SURCHARGECODE,'') SURCHARGECODE,  ISNULL(SPH.SURCHARGEPER,0.0) SURCHARGEPER,ISNULL(SPH.ADDVATCODE,'') ADDVATCODE, ISNULL(SPH.ADDVATPER,0.0) ADDVATPER , ISNULL(SPH.EXCISECODE_CR,'') EXCISECODE_CR,ISNULL(SPH.EXCISEPER_CR,0.0) EXCISEPER_CR,ISNULL(SPH.CESSCODE_CR,'') CESSCODE_CR, ISNULL(SPH.CESSPER_CR,0.0) CESSPER_CR" & _
                             " , ISNULL(SPH.SALESTAXCODE_CR,'') SALESTAXCODE_CR,ISNULL(SPH.SALESTAXPER_CR,0.0) SALESTAXPER_CR,ISNULL(SPH.AEDCODE_CR,'') AEDCODE_CR,ISNULL(SPH.AEDPER_CR,0.0) AEDPER_CR,ISNULL(SPH.SECESSCODE_CR,'') SECESSCODE_CR, ISNULL(SPH.SECESSPER_CR,0.0) SECESSPER_CR,ISNULL(SPH.SURCHARGECODE_CR,'') SURCHARGECODE_CR,ISNULL(SPH.SURCHARGEPER_CR,0.0) SURCHARGEPER_CR ,ISNULL(SPH.ADDVATCODE_CR,'') ADDVATCODE_CR,ISNULL(SPH.ADDVATPER_CR,0.0) ADDVATPER_CR,SPH.IsSubmitforAuthorized " & _
                             " , ISNULL(SPH.ExciseVal,0.0) ExciseVal,ISNULL(SPH.CessVal,0.0) CessVal,ISNULL(SPH.AEDVal,0.0) AEDVal,ISNULL(SPH.SEcessVal,0.0) SEcessVal,ISNULL(SPH.SurchargeVal,0.0) SurchargeVal, ISNULL(SPH.AddVATVal,0.0) AddVATVal, ISNULL(SPH.TotalTaxableVal,0.0) TotalTaxableVal, ISNULL(SPH.TotalNetVal,0.0) TotalNetVal,ISNULL(SPH.SalesTaxVal,0.0) SalesTaxVal" & _
                             " , ISNULL(SPH.ExciseVal_CR,0.0) ExciseVal_CR, ISNULL(SPH.CessVal_CR,0.0) CessVal_CR,ISNULL(SPH.AEDVal_CR,0.0) AEDVal_CR ,ISNULL(SPH.SEcessVal_CR,0.0) SEcessVal_CR,ISNULL(SPH.SurchargeVal_CR,0.0) SurchargeVal_CR,ISNULL(SPH.AddVATVal_CR,0.0) AddVATVal_CR, ISNULL(SPH.TotalTaxableVal_CR,0.0) TotalTaxableVal_CR,ISNULL(SPH.TotalNetVal_CR,0.0) TotalNetVal_CR,ISNULL(SPH.SalesTaxVal_CR,0.0) SalesTaxVal_CR" & _
                             " , ISNULL(cast(SPH.CreditNoteValue as numeric(18,2)),0.0) CreditNoteValue,ISNULL(cast(SPH.SupplInvoiceValue as numeric(18,2)),0.0) SupplInvoiceValue,isnull(cast(SPH.RoundOff_Diff as numeric(18,2)),0.0) RoundOff_Diff , ISNULL(cast(SPH.RoundOff_Diff_CR as numeric(18,2)),0.0) RoundOff_Diff_CR" & _
                             " FROM SALES_PROV_HDR SPH left join Employee_Mst EMKAM" & _
                             " on SPH.KAMCode=EMKAM.employee_code and SPH.Unit_code=EMKAM.Unit_Code" & _
                             " left join Employee_Mst EM" & _
                             " on SPH.PERSONTO_AUTH=EM.employee_code and SPH.Unit_code=EM.Unit_Code" & _
                             " LEFT JOIN CUSTOMER_MST CM ON SPH.CUSTOMERCODE=CM.CUSTOMER_CODE AND SPH.UNIT_CODE=CM.UNIT_CODE" & _
                             " WHERE convert(varchar(20),SPH.PROV_DOCNO)='" + txtProvDocNo.Text.Trim() + "' AND SPH.UNIT_CODE='" + gstrUNITID + "'"
                    odt = SqlConnectionclass.GetDataTable(strSQL)
                    If odt.Rows.Count > 0 Then
                        dtpFrm.Value = Convert.ToDateTime(odt.Rows(0)("FROMDATE"))
                        dtpToDt.Value = Convert.ToDateTime(odt.Rows(0)("TODATE"))
                        txtEmpCode.Text = Convert.ToString(odt.Rows(0)("PERSONTO_AUTH"))
                        txtEmpName.Text = Convert.ToString(odt.Rows(0)("PERSONTO_AUTHName"))
                        txtCustCode.Text = Convert.ToString(odt.Rows(0)("CUSTOMERCODE"))
                        lblCustDesc.Text = Convert.ToString(odt.Rows(0)("CUST_NAME"))
                        lblKAMName.Text = Convert.ToString(odt.Rows(0)("KAMName"))
                        txtExcise.Text = Convert.ToString(odt.Rows(0)("EXCISECODE"))
                        txtPerExcise.Text = Convert.ToString(odt.Rows(0)("EXCISEPER"))
                        txtCESS.Text = Convert.ToString(odt.Rows(0)("CESSCODE"))
                        txtPerCESS.Text = Convert.ToString(odt.Rows(0)("CESSPER"))
                        txtSalesTaxCode.Text = Convert.ToString(odt.Rows(0)("SALESTAXCODE"))
                        txtPerSalesTax.Text = Convert.ToString(odt.Rows(0)("SALESTAXPER"))
                        txtAEDCode.Text = Convert.ToString(odt.Rows(0)("AEDCODE"))
                        txtPerAED.Text = Convert.ToString(odt.Rows(0)("AEDPER"))
                        txtSECESS.Text = Convert.ToString(odt.Rows(0)("SECESSCODE"))
                        txtPerSECESS.Text = Convert.ToString(odt.Rows(0)("SECESSPER"))
                        txtSurcharge.Text = Convert.ToString(odt.Rows(0)("SURCHARGECODE"))
                        txtPerSurcharge.Text = Convert.ToString(odt.Rows(0)("SURCHARGEPER"))
                        txtAddVAT.Text = Convert.ToString(odt.Rows(0)("ADDVATCODE"))
                        txtPerAddVAT.Text = Convert.ToString(odt.Rows(0)("ADDVATPER"))
                        txtExcise_CR.Text = Convert.ToString(odt.Rows(0)("EXCISECODE_CR"))
                        txtExcisePER_CR.Text = Convert.ToString(odt.Rows(0)("EXCISEPER_CR"))
                        txtCESS_CR.Text = Convert.ToString(odt.Rows(0)("CESSCODE_CR"))
                        txtPerCESS_CR.Text = Convert.ToString(odt.Rows(0)("CESSPER_CR"))
                        txtSalesTaxCode_CR.Text = Convert.ToString(odt.Rows(0)("SALESTAXCODE_CR"))
                        txtPerSalesTax_CR.Text = Convert.ToString(odt.Rows(0)("SALESTAXPER_CR"))
                        txtAEDCode_CR.Text = Convert.ToString(odt.Rows(0)("AEDCODE_CR"))
                        txtPerAED_CR.Text = Convert.ToString(odt.Rows(0)("AEDPER_CR"))
                        txtSECESS_CR.Text = Convert.ToString(odt.Rows(0)("SECESSCODE_CR"))
                        txtPerSECESS_CR.Text = Convert.ToString(odt.Rows(0)("SECESSPER_CR"))
                        txtSurcharge_Neg.Text = Convert.ToString(odt.Rows(0)("SURCHARGECODE_CR"))
                        txtPerSurcharge_neg.Text = Convert.ToString(odt.Rows(0)("SURCHARGEPER_CR"))
                        txtAddVAT_CR.Text = Convert.ToString(odt.Rows(0)("ADDVATCODE_CR"))
                        txtAddPerVAT_CR.Text = Convert.ToString(odt.Rows(0)("ADDVATPER_CR"))

                        If strHelp(5).Trim.ToString = "NORMAL" Then
                            rbtn_NormalInvoice.Checked = True
                            BtnFetch_Click(BtnFetch, New EventArgs())
                            bindRateItemfpGrid()
                        End If
                        If strHelp(5).Trim.ToString = "MARUTI" Then
                            'rbtn_MarutiFile.Checked = True
                            getdetaildata(txtProvDocNo.Text.Trim())
                        End If
                        If Convert.ToBoolean(odt.Rows(0)("IsSubmitforAuthorized")) Then
                            InitializeForm(5)
                            txtExciseVal.Text = Convert.ToDecimal(odt.Rows(0)("ExciseVal")).ToString("0.00")
                            txtCESSVal.Text = Convert.ToDecimal(odt.Rows(0)("CessVal")).ToString("0.00")
                            txtSECESSVal.Text = Convert.ToDecimal(odt.Rows(0)("SEcessVal")).ToString("0.00")
                            txtAEDVal.Text = Convert.ToDecimal(odt.Rows(0)("AEDVal")).ToString("0.00")
                            txtSalesTaxVal.Text = Convert.ToDecimal(odt.Rows(0)("SalesTaxVal")).ToString("0.00")
                            txtSurchargeVal.Text = Convert.ToDecimal(odt.Rows(0)("SurchargeVal")).ToString("0.00")
                            txtAddVATVal.Text = Convert.ToDecimal(odt.Rows(0)("AddVATVal")).ToString("0.00")
                            txtTotAssVal.Text = Convert.ToString(odt.Rows(0)("TotalTaxableVal"))
                            txtNetVal.Text = Convert.ToDecimal(odt.Rows(0)("TotalNetVal")).ToString("0.00")
                            txtRoundOffBy.Text = Convert.ToString(odt.Rows(0)("RoundOff_Diff"))
                            txtExciseVal_CR.Text = Convert.ToDecimal(odt.Rows(0)("ExciseVal_CR")).ToString("0.00")
                            txtCESSVal_CR.Text = Convert.ToDecimal(odt.Rows(0)("CessVal_CR")).ToString("0.00")
                            txtSECESSVal_CR.Text = Convert.ToDecimal(odt.Rows(0)("SEcessVal_CR")).ToString("0.00")
                            txtAEDVal_CR.Text = Convert.ToDecimal(odt.Rows(0)("AEDVal_CR")).ToString("0.00")
                            txtSalesTaxVal_CR.Text = Convert.ToDecimal(odt.Rows(0)("SalesTaxVal_CR")).ToString("0.00")
                            txtSurchargeVal_Neg.Text = Convert.ToDecimal(odt.Rows(0)("SurchargeVal_CR")).ToString("0.00")
                            txtAddVATVal_CR.Text = Convert.ToDecimal(odt.Rows(0)("AddVATVal_CR")).ToString("0.00")
                            txtTotAssVal_Neg.Text = Convert.ToString(odt.Rows(0)("TotalTaxableVal_CR"))
                            txtNetVal_neg.Text = Convert.ToDecimal(odt.Rows(0)("TotalNetVal_CR")).ToString("0.00")
                            txtRoundOffBy_Neg.Text = Convert.ToString(odt.Rows(0)("RoundOff_Diff_CR"))
                            txtPositiveVal.Text = Convert.ToString(odt.Rows(0)("SupplInvoiceValue"))
                            txtNegativeVal.Text = Convert.ToString(odt.Rows(0)("CreditNoteValue"))
                        Else
                            InitializeForm(3)
                            CalculateTaxes()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub dtpEffFrm_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpFrm.Validated
        Try
            If sender Is dtpFrm Then
                If dtpFrm.Value.CompareTo(dtpToDt.Value) > 0 Then
                    MessageBox.Show("From date cannot be greater then To date.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    dtpFrm.Value = GetServerDate().AddMonths(-1)
                    dtpToDt.Value = GetServerDate()
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub dtpEffTo_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpToDt.Validated
        Try
            If dtpToDt.Value.CompareTo(dtpFrm.Value) < 0 Then
                MessageBox.Show("Effective From date cannot be greater then Effective To date.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                dtpToDt.Value = dtpFrm.Value
                dtpFrm.Value = dtpToDt.Value.AddMonths(-1)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub BtnHelpHdr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnHelpEmp.Click, BtnHelpCustCode.Click
        Try
            Dim strSQL As String = String.Empty
            Dim strHelp As String()
            With ctlHelp
                .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                .ConnectAsUser = gstrCONNECTIONUSER
                .ConnectThroughDSN = gstrCONNECTIONDSN
                .ConnectWithPWD = gstrCONNECTIONPASSWORD
            End With
            If (sender Is BtnHelpEmp) Then
                strSQL = " select Employee_code As [EmployeeCode], Name As [EmployeeName] from Employee_mst(NOLOCK) where UNIT_CODE ='" & gstrUNITID & "'" & _
                        " and Employee_code not in (select employee_code from user_mst where unit_code='" + gstrUNITID + "' and User_id='" + mP_User + "')"
                strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select Employee", 1, 0, txtEmpName.Text.Trim)
                If Not IsNothing(strHelp) Then
                    If strHelp.Length > 0 Then
                        txtEmpCode.Text = strHelp(0)
                        txtEmpName.Text = strHelp(1)
                    Else
                        txtEmpCode.Text = String.Empty
                        txtEmpName.Text = String.Empty
                    End If
                End If
            ElseIf (sender Is BtnHelpCustCode) Then
                If rbtn_NormalInvoice.Checked = True Then
                    strSQL = "SELECT DISTINCT CM.CUSTOMER_CODE, CM.CUST_NAME [Customer_Name], CM.KAMCODE [KAM_Code], EM.Name [KAM_Name]" & _
                             " FROM CUSTOMER_MST CM left join Employee_mst EM" & _
                             " on CM.UNIT_CODE=EM.UNIT_CODE and CM.KAMCODE=EM.Employee_code" & _
                             " WHERE CM.UNIT_CODE='" & gstrUNITID & "'" & _
                             " AND CM.CUSTOMER_CODE IN (SELECT DISTINCT ACCOUNT_CODE FROM SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "'" & _
                             " AND INVOICE_DATE BETWEEN '" + dtpFrm.Value.ToString("dd MMM yyyy") + "' AND '" + dtpToDt.Value.ToString("dd MMM yyyy") + "'" & _
                             " AND INVOICE_TYPE='INV' AND BILL_FLAG=1 AND CANCEL_FLAG=0" + ")"
                    strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select Customer Code", 1, 0, txtCustCode.Text.Trim)
                    If Not IsNothing(strHelp) Then
                        If strHelp.Length = 4 Then
                            txtCustCode.Text = strHelp(0)
                            lblCustDesc.Text = strHelp(1)
                            lblKAMName.Text = strHelp(3)
                            lblKAMName.Tag = strHelp(2)
                            BtnFetch.Enabled = True
                            ClearDataGridView(dgvDispatchDtl)
                            clearFarGrid()
                            ClearAllTaxes()
                        Else
                            txtCustCode.Text = String.Empty
                            lblCustDesc.Text = String.Empty
                            lblKAMName.Text = String.Empty
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub btnFreeze_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub btnSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim flag As Boolean = False
        Try
            If Not String.IsNullOrEmpty(txtProvDocNo.Text) Then
                If DialogResult.Yes = MessageBox.Show("Are you sure, you want to submit?", ResolveResString(100), MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) Then

                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub txtSplChars_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustCode.KeyPress
        If e.KeyChar = "'" Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtProvDocNo_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtProvDocNo.KeyPress
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    Private Sub BtnFetch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFetch.Click
        Dim strQuery As String = String.Empty
        Dim strHelp As String()
        Dim selectedItems As String = String.Empty
        Try
            If rbtn_NormalInvoice.Checked = True Then
                ClearTmpTable()
                FillTmpTable()
                If cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    'strQuery = " SELECT TMP.ITEM_CODE, TMP.VARMODEL" & _
                    '        " , CAST(SUM(TMP.INVOICE_QTY) AS NUMERIC(18,2)) INVOICE_QTY, CAST(SUM(TMP.ACTUALINVOICEQTY) AS NUMERIC(18,2)) ACTUALINVOICEQTY" & _
                    '        " , ISNULL(IM.DESCRIPTION,'') [ITEM_DESC], ISNULL(BM.MODEL_DESC,'') MODELDESC" & _
                    '        " FROM SALES_PROV_TMPPARTDETAIL TMP" & _
                    '        " INNER JOIN(" & _
                    '        "  SELECT DISTINCT ITEM_CODE, VARMODEL, UNIT_CODE FROM SALES_PROV_RATEWISEDETAIL WHERE CAST(PROV_DOCNO AS VARCHAR(20))='" + txtProvDocNo.Text.Trim() + "' AND UNIT_CODE='" + gstrUNITID + "' " & _
                    '        " ) SRQ " & _
                    '        " ON TMP.ITEM_CODE=SRQ.ITEM_CODE AND TMP.VARMODEL=SRQ.VARMODEL AND TMP.UNIT_CODE=SRQ.UNIT_CODE " & _
                    '        " LEFT JOIN ITEM_MST IM" & _
                    '        " ON TMP.ITEM_CODE=IM.ITEM_CODE AND TMP.UNIT_CODE=IM.UNIT_CODE" & _
                    '        " LEFT JOIN BUDGET_MODEL_MST BM" & _
                    '        " ON TMP.VARMODEL=BM.MODEL_CODE AND TMP.UNIT_CODE=BM.UNIT_CODE AND BM.ACTIVE=1" & _
                    '        " WHERE TMP.IPADDRESS='" + gstrIpaddressWinSck + "' AND TMP.UNIT_CODE='" + gstrUNITID + "' " & _
                    '        " GROUP BY TMP.CUSTOMER_CODE, TMP.ITEM_CODE, TMP.VARMODEL,CASE ISNULL(SRQ.ITEM_CODE,'') WHEN '' THEN CAST(0 AS BIT) ELSE CAST(1 AS BIT) END" & _
                    '        " , ISNULL(IM.DESCRIPTION,''), ISNULL(BM.MODEL_DESC,'')"
                    strQuery = " SELECT TMP.ITEM_CODE,cust .Cust_Drgno CustomeritemCode ,cust.Drg_Desc CustomerItemDesc, TMP.VARMODEL" & _
                           " , CAST(SUM(TMP.INVOICE_QTY) AS NUMERIC(18,2)) INVOICE_QTY, CAST(SUM(TMP.ACTUALINVOICEQTY) AS NUMERIC(18,2)) ACTUALINVOICEQTY" & _
                           " , ISNULL(IM.DESCRIPTION,'') [ITEM_DESC], ISNULL(BM.MODEL_DESC,'') MODELDESC" & _
                           " FROM SALES_PROV_TMPPARTDETAIL TMP" & _
                           " INNER JOIN(" & _
                           "  SELECT DISTINCT ITEM_CODE, VARMODEL, UNIT_CODE FROM SALES_PROV_RATEWISEDETAIL WHERE CAST(PROV_DOCNO AS VARCHAR(20))='" + txtProvDocNo.Text.Trim() + "' AND UNIT_CODE='" + gstrUNITID + "' " & _
                           " ) SRQ " & _
                           " ON TMP.ITEM_CODE=SRQ.ITEM_CODE AND TMP.VARMODEL=SRQ.VARMODEL AND TMP.UNIT_CODE=SRQ.UNIT_CODE " & _
                           " LEFT JOIN ITEM_MST IM" & _
                           " ON TMP.ITEM_CODE=IM.ITEM_CODE AND TMP.UNIT_CODE=IM.UNIT_CODE" & _
                           " LEFT JOIN BUDGET_MODEL_MST BM" & _
                           " ON TMP.VARMODEL=BM.MODEL_CODE AND TMP.UNIT_CODE=BM.UNIT_CODE AND BM.ACTIVE=1" & _
                           " LEFT JOIN CUSTITEM_MST CUST ON TMP.ITEM_CODE=CUST.ITEM_CODE AND " & _
                           " TMP.CUSTOMER_CODE=CUST.ACCOUNT_CODE AND TMP.UNIT_CODE=CUST.UNIT_CODE  AND CUST.ACTIVE =1 " & _
                           " WHERE TMP.IPADDRESS='" + gstrIpaddressWinSck + "' AND TMP.UNIT_CODE='" + gstrUNITID + "' " & _
                           " GROUP BY TMP.CUSTOMER_CODE, TMP.ITEM_CODE,cust .Cust_Drgno ,cust.Drg_Desc, TMP.VARMODEL,CASE ISNULL(SRQ.ITEM_CODE,'') WHEN '' THEN CAST(0 AS BIT) ELSE CAST(1 AS BIT) END" & _
                           " , ISNULL(IM.DESCRIPTION,''), ISNULL(BM.MODEL_DESC,'')"
                Else
                    'strQuery = " SELECT CASE ISNULL(SRQ.ITEM_CODE,'') WHEN '' THEN CAST(0 AS BIT) ELSE CAST(1 AS BIT) END [SELECT], TMP.ITEM_CODE, TMP.VARMODEL, ISNULL(BM.MODEL_DESC,'') MODELDESC" & _
                    '      " , CAST(SUM(TMP.INVOICE_QTY) AS NUMERIC(18,2)) INVOICE_QTY, CAST(SUM(TMP.ACTUALINVOICEQTY) AS NUMERIC(18,2)) ACTUALINVOICEQTY" & _
                    '      " , ISNULL(IM.DESCRIPTION,'') [ITEM_DESC]" & _
                    '      " FROM SALES_PROV_TMPPARTDETAIL TMP" & _
                    '      " LEFT JOIN(" & _
                    '      "  SELECT DISTINCT ITEM_CODE, VARMODEL, UNIT_CODE FROM SALES_PROV_RATEWISEDETAIL WHERE CAST(PROV_DOCNO AS VARCHAR(20))='" + txtProvDocNo.Text.Trim() + "' AND UNIT_CODE='" + gstrUNITID + "' " & _
                    '      " ) SRQ " & _
                    '      " ON TMP.ITEM_CODE=SRQ.ITEM_CODE AND TMP.VARMODEL=SRQ.VARMODEL AND TMP.UNIT_CODE=SRQ.UNIT_CODE " & _
                    '      " LEFT JOIN ITEM_MST IM" & _
                    '      " ON TMP.ITEM_CODE=IM.ITEM_CODE AND TMP.UNIT_CODE=IM.UNIT_CODE" & _
                    '      " LEFT JOIN BUDGET_MODEL_MST BM" & _
                    '      " ON TMP.VARMODEL=BM.MODEL_CODE AND TMP.UNIT_CODE=BM.UNIT_CODE AND BM.ACTIVE=1" & _
                    '      " WHERE TMP.IPADDRESS='" + gstrIpaddressWinSck + "' AND TMP.UNIT_CODE='" + gstrUNITID + "'" & _
                    '      " GROUP BY TMP.CUSTOMER_CODE, TMP.ITEM_CODE, TMP.VARMODEL,CASE ISNULL(SRQ.ITEM_CODE,'') WHEN '' THEN CAST(0 AS BIT) ELSE CAST(1 AS BIT) END" & _
                    '      " , ISNULL(IM.DESCRIPTION,''), ISNULL(BM.MODEL_DESC,'')"
                    strQuery = " SELECT CASE ISNULL(SRQ.ITEM_CODE,'') WHEN '' THEN CAST(0 AS BIT) ELSE CAST(1 AS BIT) END [SELECT], TMP.ITEM_CODE,cust .Cust_Drgno CustomeritemCode ,cust.Drg_Desc CustomerItemDesc,TMP.VARMODEL, ISNULL(BM.MODEL_DESC,'') MODELDESC" & _
                         " , CAST(SUM(TMP.INVOICE_QTY) AS NUMERIC(18,2)) INVOICE_QTY, CAST(SUM(TMP.ACTUALINVOICEQTY) AS NUMERIC(18,2)) ACTUALINVOICEQTY" & _
                         " , ISNULL(IM.DESCRIPTION,'') [ITEM_DESC]" & _
                         " FROM SALES_PROV_TMPPARTDETAIL TMP" & _
                         " LEFT JOIN(" & _
                         "  SELECT DISTINCT ITEM_CODE, VARMODEL, UNIT_CODE FROM SALES_PROV_RATEWISEDETAIL WHERE CAST(PROV_DOCNO AS VARCHAR(20))='" + txtProvDocNo.Text.Trim() + "' AND UNIT_CODE='" + gstrUNITID + "' " & _
                         " ) SRQ " & _
                         " ON TMP.ITEM_CODE=SRQ.ITEM_CODE AND TMP.VARMODEL=SRQ.VARMODEL AND TMP.UNIT_CODE=SRQ.UNIT_CODE " & _
                         " LEFT JOIN ITEM_MST IM" & _
                         " ON TMP.ITEM_CODE=IM.ITEM_CODE AND TMP.UNIT_CODE=IM.UNIT_CODE" & _
                         " LEFT JOIN BUDGET_MODEL_MST BM" & _
                         " ON TMP.VARMODEL=BM.MODEL_CODE AND TMP.UNIT_CODE=BM.UNIT_CODE AND BM.ACTIVE=1" & _
                         " LEFT JOIN CUSTITEM_MST CUST ON TMP.ITEM_CODE=CUST.ITEM_CODE AND " & _
                         " TMP.CUSTOMER_CODE=CUST.ACCOUNT_CODE AND TMP.UNIT_CODE=CUST.UNIT_CODE  AND CUST.ACTIVE =1 " & _
                         " WHERE TMP.IPADDRESS='" + gstrIpaddressWinSck + "' AND TMP.UNIT_CODE='" + gstrUNITID + "'" & _
                         " GROUP BY TMP.CUSTOMER_CODE, TMP.ITEM_CODE,cust .Cust_Drgno ,cust.Drg_Desc, TMP.VARMODEL,CASE ISNULL(SRQ.ITEM_CODE,'') WHEN '' THEN CAST(0 AS BIT) ELSE CAST(1 AS BIT) END" & _
                         " , ISNULL(IM.DESCRIPTION,''), ISNULL(BM.MODEL_DESC,'')"
                    If Not IsNothing(dtSelItems) Then
                        If dtSelItems.Rows.Count > 0 Then
                            For Each dr As DataRow In dtSelItems.Rows
                                selectedItems = selectedItems + dr("ITEM_CODE") + "|"
                            Next
                            selectedItems = selectedItems.Remove(selectedItems.LastIndexOf("|"), 1)
                        End If
                    End If
                    With ctlHelp
                        .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                        .ConnectAsUser = gstrCONNECTIONUSER
                        .ConnectThroughDSN = gstrCONNECTIONDSN
                        .ConnectWithPWD = gstrCONNECTIONPASSWORD
                    End With

                    strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Select Part(s)", 0, True, selectedItems)
                    If Not IsNothing(strHelp) Then
                        strQuery = String.Empty
                        If strHelp.Length = 10 Then
                            If Not IsNothing(strHelp(9)) Then
                                strQuery = strHelp(9).Replace("|", "','")
                            End If
                        End If
                        strQuery = " SELECT TMP.ITEM_CODE,cust .Cust_Drgno CustomeritemCode ,cust.Drg_Desc CustomerItemDesc, TMP.VARMODEL" & _
                        " , CAST(SUM(TMP.INVOICE_QTY) AS NUMERIC(18,2)) INVOICE_QTY, CAST(SUM(TMP.ACTUALINVOICEQTY) AS NUMERIC(18,2)) ACTUALINVOICEQTY" & _
                        " , ISNULL(IM.DESCRIPTION,'') [ITEM_DESC], ISNULL(BM.MODEL_DESC,'') MODELDESC" & _
                        " FROM SALES_PROV_TMPPARTDETAIL TMP" & _
                        " LEFT JOIN(" & _
                        "  SELECT DISTINCT ITEM_CODE, VARMODEL, UNIT_CODE FROM SALES_PROV_RATEWISEDETAIL WHERE CAST(PROV_DOCNO AS VARCHAR(20))='" + txtProvDocNo.Text.Trim() + "' AND UNIT_CODE='" + gstrUNITID + "' " & _
                        " ) SRQ " & _
                        " ON TMP.ITEM_CODE=SRQ.ITEM_CODE AND TMP.VARMODEL=SRQ.VARMODEL AND TMP.UNIT_CODE=SRQ.UNIT_CODE " & _
                        " LEFT JOIN ITEM_MST IM" & _
                        " ON TMP.ITEM_CODE=IM.ITEM_CODE AND TMP.UNIT_CODE=IM.UNIT_CODE" & _
                        " LEFT JOIN BUDGET_MODEL_MST BM" & _
                        " ON TMP.VARMODEL=BM.MODEL_CODE AND TMP.UNIT_CODE=BM.UNIT_CODE AND BM.ACTIVE=1" & _
                        " LEFT JOIN CUSTITEM_MST CUST ON TMP.ITEM_CODE=CUST.ITEM_CODE AND " & _
                         " TMP.CUSTOMER_CODE=CUST.ACCOUNT_CODE AND TMP.UNIT_CODE=CUST.UNIT_CODE  AND CUST.ACTIVE =1 " & _
                        " WHERE TMP.IPADDRESS='" + gstrIpaddressWinSck + "' AND TMP.UNIT_CODE='" + gstrUNITID + "' and TMP.ITEM_COde in ('" + strQuery + "')" & _
                        " GROUP BY TMP.CUSTOMER_CODE, TMP.ITEM_CODE, cust .Cust_Drgno ,cust.Drg_Desc, TMP.VARMODEL,CASE ISNULL(SRQ.ITEM_CODE,'') WHEN '' THEN CAST(0 AS BIT) ELSE CAST(1 AS BIT) END" & _
                        " , ISNULL(IM.DESCRIPTION,''), ISNULL(BM.MODEL_DESC,'')"
                    End If
                End If
                dtSelItems = New DataTable
                dtSelItems = SqlConnectionclass.GetDataTable(strQuery)
                If (Not IsNothing(dtSelItems)) Then
                    dgvDispatchDtl.DataSource = dtSelItems
                    If (dtSelItems.Rows.Count = 0) Then
                        MessageBox.Show("No Record found.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                End If
                clearFarGrid()
                bindRateItemfpGrid()
            End If

            If rbtn_MarutiFile.Checked = True Then
                If cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    If txt_marutiFile.Text.Trim.ToString() <> "" Then
                        If txtCustCode.Text.Trim.ToString <> "" Then
                            strQuery = ""
                            strQuery = "SELECT TMP.ITEM_CODE,MST.ITEM_DESC,cust .Cust_Drgno CustomeritemCode ,ISNULL(BM.MODEL_CODE,'') AS MODELCODE ,ISNULL(BM.MODEL_DESC,'') AS MODELDESC ,SUM(TMP.STRSHP) AS Invoice_Qty ,SUM(TMP.STRACP) AS ActualInvoiceQty" & _
                            " FROM  TMP_SALESFILEDATA TMP INNER JOIN CUSTITEM_MST MST ON TMP.STRPARTNO =MST.CUST_DRGNO AND TMP.UNIT_CODE=MST.UNIT_CODE " & _
                            " AND TMP.ACCOUNT_CODE =MST.ACCOUNT_CODE AND TMP.ITEM_CODE =MST.ITEM_CODE" & _
                            " INNER JOIN MARUTI_FILE_HDR HDR ON TMP.UNIT_CODE=HDR.UNITCODE AND TMP.BATCH_NO=HDR.BATCH_NO" & _
                            " LEFT OUTER JOIN BUDGET_MODEL_MST BM ON BM.MODEL_CODE=MST.VARMODEL AND BM.UNIT_CODE =TMP.UNIT_CODE AND BM.ACTIVE=1 " & _
                            " LEFT JOIN CUSTITEM_MST CUST ON TMP.ITEM_CODE=CUST.ITEM_CODE AND " & _
                            " TMP.ACCOUNT_CODE=CUST.ACCOUNT_CODE AND TMP.UNIT_CODE=CUST.UNIT_CODE  AND CUST.ACTIVE =1 " & _
                            " WHERE TMP.UNIT_CODE ='" + gstrUNITID + "' AND TMP.ACCOUNT_CODE ='" + txtCustCode.Text.Trim.ToString() + "' and hdr.filename='" + txt_marutiFile.Text.Trim.ToString() + "'" & _
                            " GROUP BY TMP.ITEM_CODE,MST.ITEM_DESC,cust .Cust_Drgno ,cust.Drg_Desc,BM.MODEL_CODE,BM.MODEL_DESC"

                            With ctlHelp
                                .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                                .ConnectAsUser = gstrCONNECTIONUSER
                                .ConnectThroughDSN = gstrCONNECTIONDSN
                                .ConnectWithPWD = gstrCONNECTIONPASSWORD
                            End With

                            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Select Part(s)")
                            If Not IsNothing(strHelp) Then
                                If Not IsNothing(strHelp(0)) Then
                                    getdata(strHelp(0).Trim.ToString())
                                End If
                            End If

                        Else
                            MsgBox("Kindly Select Customer First!!")
                        End If
                    Else
                        MsgBox("Kindly Select File First!!")
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub dgvDispatchDtl_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDispatchDtl.CellContentClick
        Try
            dgvDispatchDtl.CommitEdit(DataGridViewDataErrorContexts.Commit)
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub fpSpreadRateWiseDtl_ComboSelChange(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ComboSelChangeEvent) Handles fpSpreadRateWiseDtl.ComboSelChange
        Try
            If rbtn_MarutiFile.Checked = False Then
                With fpSpreadRateWiseDtl
                    .Row = e.row
                    If e.col = RateWiseDtlGrid.Col_PriceChange Then
                        .Col = RateWiseDtlGrid.Col_NewRate
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .Value = ""
                        .Col = RateWiseDtlGrid.Col_Change
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .Value = ""
                        .Col = RateWiseDtlGrid.Col_ReasonChange
                        .Value = ""
                        .Col = RateWiseDtlGrid.Col_NewInvRate
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .Value = ""
                        .Col = RateWiseDtlGrid.Col_TotEffVal
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .Value = ""
                        .Col = RateWiseDtlGrid.Col_CorrectionNature
                        .Value = ""
                        LockUnlockRateWiseGrid(e.row, e.row, False)
                    ElseIf e.col = RateWiseDtlGrid.Col_ChangeEff Then
                        CalculateRateEffect(e.row)
                        GetTotalBasicValEffect()
                        CalculateTaxes()
                    End If
                End With
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub fpSpreadRateWiseDtl_LeaveCell(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles fpSpreadRateWiseDtl.LeaveCell
        Try
            If e.row > 0 Then
                With fpSpreadRateWiseDtl
                    .Row = e.row
                    .Col = RateWiseDtlGrid.Col_Select
                    If .Value = 1 Then
                        If e.col = RateWiseDtlGrid.Col_Change Or e.col = RateWiseDtlGrid.Col_NewRate Then
                            CalculateRateEffect(e.row)
                            GetTotalBasicValEffect()
                            CalculateTaxes()
                        End If
                    End If
                    'fpSpreadRateWiseDtl.SetActiveCell(e.col, e.row)
                    'fpSpreadRateWiseDtl_Enter(fpSpreadRateWiseDtl, New System.EventArgs)
                End With
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
                            ' mayur
                            If rbtn_MarutiFile.Checked = True Then
                                CalculateRateEffect(e.row)
                                GetTotalBasicValEffect()
                                CalculateTaxes()
                            End If
                            ' mayur
                        ElseIf .Value = 0 And rbtn_MarutiFile.Checked = True Then
                            CalculateRateEffect(e.row)
                            GetTotalBasicValEffect()
                            CalculateTaxes()
                        Else
                            If rbtn_MarutiFile.Checked = False And cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
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
                        End If
                    ElseIf e.col = RateWiseDtlGrid.Col_ShowInv Then
                        .Col = RateWiseDtlGrid.Col_Part_Code
                        .Row = e.row
                        Dim frmObj As New FRMMKTTRN0084A
                        frmObj.gProvDocNo = txtProvDocNo.Text.Trim()
                        frmObj.gItem_Code = Convert.ToString(.Value).Trim()
                        .Col = RateWiseDtlGrid.Col_InvRate
                        frmObj.gInvoiceRate = Convert.ToString(.Value).Trim()
                        If rbtn_MarutiFile.Checked = True Then
                            frmObj.gfilemode = True
                            .Col = RateWiseDtlGrid.Col_NewRate
                            frmObj.gNewRate = Convert.ToString(.Value).Trim()
                        End If
                        If rbtn_MarutiFile.Checked = False Then
                            frmObj.gfilemode = False
                        End If
                        frmObj.ShowDialog()
                    End If
                End With
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub fpSpreadRateWiseDtl_KeyDownEvent(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles fpSpreadRateWiseDtl.KeyDownEvent
        Dim strHelp As String()
        Dim strQuery As String = String.Empty
        Try
            If e.keyCode = Keys.F1 Then
                If fpSpreadRateWiseDtl.ActiveCol = RateWiseDtlGrid.Col_CorrectionNature And fpSpreadRateWiseDtl.ActiveRow > 0 And (cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Or cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD) Then
                    fpSpreadRateWiseDtl.Row = fpSpreadRateWiseDtl.ActiveRow
                    fpSpreadRateWiseDtl.Col = RateWiseDtlGrid.Col_Select
                    If fpSpreadRateWiseDtl.Value = 1 Then
                        strQuery = "select CorrectionID ID, CorrectionDesc Description from Sales_Prov_NatureofCorrection"
                        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Select Nature of Correction")
                        If Not IsNothing(strHelp) Then
                            If strHelp.Length > 0 Then
                                fpSpreadRateWiseDtl.Col = RateWiseDtlGrid.Col_CorrectionNature
                                fpSpreadRateWiseDtl.Value = strHelp(1)
                                fpSpreadRateWiseDtl.CellTag = strHelp(0)
                            End If
                        End If
                    End If
                End If

                ' Code Added By Mayur Against issue ID 10816097 
                If fpSpreadRateWiseDtl.ActiveCol = RateWiseDtlGrid.Col_NewRate And fpSpreadRateWiseDtl.ActiveRow > 0 And (cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD) And rbtn_MarutiFile.Checked = False Then
                    fpSpreadRateWiseDtl.Row = fpSpreadRateWiseDtl.ActiveRow
                    fpSpreadRateWiseDtl.Col = RateWiseDtlGrid.Col_Select
                    If fpSpreadRateWiseDtl.Value = 1 Then
                        fpSpreadRateWiseDtl.Col = RateWiseDtlGrid.Col_PriceChange
                        fpSpreadRateWiseDtl.Row = fpSpreadRateWiseDtl.ActiveRow
                        If (fpSpreadRateWiseDtl.Text = "Value") Then
                            fpSpreadRateWiseDtl.Col = RateWiseDtlGrid.Col_Part_Code
                            fpSpreadRateWiseDtl.Row = fpSpreadRateWiseDtl.ActiveRow
                            part_code = fpSpreadRateWiseDtl.Value.Trim.ToString()
                            fpSpreadRateWiseDtl.Col = RateWiseDtlGrid.Col_InvRate
                            fpSpreadRateWiseDtl.Row = fpSpreadRateWiseDtl.ActiveRow
                            rate = fpSpreadRateWiseDtl.Value
                            If (part_code.Trim().ToString() <> "") Then
                                GetPriceChangeDetails(part_code.Trim().ToString(), rate, "ADD")
                                fpSpreadRateWiseDtl.Col = RateWiseDtlGrid.Col_NewRate
                                fpSpreadRateWiseDtl.Row = fpSpreadRateWiseDtl.ActiveRow
                                fpSpreadRateWiseDtl.Lock = True
                            End If
                        End If
                    End If
                End If

                If txtProvDocNo.Text <> "" Then
                    If fpSpreadRateWiseDtl.ActiveCol = RateWiseDtlGrid.Col_NewRate And fpSpreadRateWiseDtl.ActiveRow > 0 And (cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Or cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT) Then

                        fpSpreadRateWiseDtl.Col = RateWiseDtlGrid.Col_PriceChange
                        fpSpreadRateWiseDtl.Row = fpSpreadRateWiseDtl.ActiveRow

                        If (fpSpreadRateWiseDtl.Text = "Value") Then

                            fpSpreadRateWiseDtl.Col = RateWiseDtlGrid.Col_InvRate
                            fpSpreadRateWiseDtl.Row = fpSpreadRateWiseDtl.ActiveRow
                            rate = fpSpreadRateWiseDtl.Value
                            fpSpreadRateWiseDtl.Col = RateWiseDtlGrid.Col_Part_Code
                            fpSpreadRateWiseDtl.Row = fpSpreadRateWiseDtl.ActiveRow
                            part_code = fpSpreadRateWiseDtl.Value.Trim.ToString()

                            If (cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT) Then
                                If (part_code.Trim().ToString() <> "") Then
                                    GetPriceChangeDetails(part_code.Trim().ToString(), rate, "EDIT")
                                End If
                            End If
                            If (cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW) Then
                                If (part_code.Trim().ToString() <> "") Then
                                    GetPriceChangeDetails(part_code.Trim().ToString(), rate, "VIEW")
                                End If
                            End If

                        End If

                    End If

                End If
                ' Code Added By Mayur Against issue ID 10816097 
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
    Private Sub BtnHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExciseHelp.Click, BtnHelpEmp.Click, BtnHelpSurcharge.Click, BtnHelpSECESS.Click, BtnHelpSalesTax.Click, btnHelpCess.Click, BtnHelpAED.Click, BtnHelpAddVAT.Click
        Dim strQry As String = String.Empty
        Dim strhelp As String()
        Try
            If Not cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                If txtPositiveVal.Text.Trim <> "" Then
                    With ctlHelp
                        .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                        .ConnectAsUser = gstrCONNECTIONUSER
                        .ConnectThroughDSN = gstrCONNECTIONDSN
                        .ConnectWithPWD = gstrCONNECTIONPASSWORD
                    End With
                    If sender Is BtnExciseHelp Then
                        strQry = " SELECT DISTINCT TXRT_RATE_NO,CAST(TXRT_PERCENTAGE AS NUMERIC(9,2)) TXRT_PERCENTAGE FROM GEN_TAXRATE " & _
                             " WHERE UNIT_CODE='" + gstrUNITID + "' AND TX_TAXEID ='EXC' AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) OR (CAST(GETDATE() AS DATE)<= DEACTIVE_DATE))"
                        strhelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry, "Select Excise")
                        If Not IsNothing(strhelp) Then
                            If Not IsNothing(strhelp(1)) Then
                                txtExcise.Text = strhelp(0).Trim()
                                txtPerExcise.Text = strhelp(1).Trim()
                            Else
                                txtExcise.Text = ""
                                txtPerExcise.Text = ""
                            End If
                        End If
                    ElseIf sender Is btnHelpCess Then
                        strQry = "Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where Unit_code='" & gstrUNITID & "' and tx_TaxeID ='ECS' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                        strhelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry, "Select Excise")
                        If Not IsNothing(strhelp) Then

                            If Not IsNothing(strhelp(1)) Then
                                txtCESS.Text = strhelp(0).Trim()
                                txtPerCESS.Text = strhelp(1).Trim()
                            Else
                                txtCESS.Text = ""
                                txtPerCESS.Text = ""
                            End If

                        End If
                    ElseIf sender Is BtnHelpSalesTax Then
                        strQry = "Select TxRT_Rate_No,TxRt_Percentage from Gen_taxRate where Unit_code='" & gstrUNITID & "' and (Tx_TaxeID ='CST' OR Tx_TaxeID ='LST' OR Tx_TaxeID ='VAT') and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                        strhelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry, "Sales Tax Help")
                        If Not IsNothing(strhelp) Then
                            If Not IsNothing(strhelp(1)) Then
                                txtSalesTaxCode.Text = strhelp(0).Trim()
                                txtPerSalesTax.Text = strhelp(1).Trim()
                            Else
                                txtSalesTaxCode.Text = ""
                                txtPerSalesTax.Text = ""
                            End If

                        End If
                    ElseIf sender Is BtnHelpAED Then
                        strQry = " SELECT DISTINCT TXRT_RATE_NO,CAST(TXRT_PERCENTAGE AS NUMERIC(9,2)) TXRT_PERCENTAGE FROM GEN_TAXRATE " & _
                             " WHERE UNIT_CODE='" + gstrUNITID + "' AND TX_TAXEID ='AED' AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) OR (CAST(GETDATE() AS DATE)<= DEACTIVE_DATE))"
                        strhelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry, "Select Excise")
                        If Not IsNothing(strhelp) Then
                            If IsNothing(strhelp(1)) Then
                                txtAEDCode.Text = ""
                                txtPerAED.Text = ""
                            Else
                                txtAEDCode.Text = strhelp(0).Trim()
                                txtPerAED.Text = strhelp(1).Trim()
                            End If
                        End If
                    ElseIf sender Is BtnHelpSECESS Then
                        strQry = "Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where Unit_code='" & gstrUNITID & "' and tx_TaxeID ='ECSSH' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                        strhelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry, "SECESS  Help")
                        If Not IsNothing(strhelp) Then
                            If Not IsNothing(strhelp(1)) Then
                                txtSECESS.Text = strhelp(0).Trim()
                                txtPerSECESS.Text = strhelp(1).Trim()
                            Else
                                txtSECESS.Text = ""
                                txtPerSECESS.Text = ""
                            End If
                        End If
                    ElseIf sender Is BtnHelpSurcharge Then
                        strQry = "Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where Unit_code='" & gstrUNITID & "' and tx_TaxeID ='SST' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                        strhelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry, "Surcharge Help")
                        If Not IsNothing(strhelp) Then
                            If Not IsNothing(strhelp(1)) Then
                                txtSurcharge.Text = strhelp(0).Trim()
                                txtPerSurcharge.Text = strhelp(1).Trim()
                            Else
                                txtSurcharge.Text = ""
                                txtPerSurcharge.Text = ""
                            End If
                        End If
                    ElseIf sender Is BtnHelpAddVAT Then
                        strQry = "Select TxRT_Rate_No,TxRt_Percentage from Gen_taxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  Tx_TaxeID in('ADVAT','ADCST')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                        strhelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry, "ADD VAT")
                        If Not IsNothing(strhelp) Then
                            If Not IsNothing(strhelp(1)) Then
                                txtAddVAT.Text = strhelp(0).Trim()
                                txtPerAddVAT.Text = strhelp(1).Trim()
                            Else
                                txtAddVAT.Text = ""
                                txtPerAddVAT.Text = ""
                            End If
                        End If
                    End If
                End If
                CalculateTaxes()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub txtHelp_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtExcise.KeyDown, txtCESS.KeyDown, txtEmpCode.KeyDown, txtCustCode.KeyDown, txtProvDocNo.KeyDown, txtSurcharge.KeyDown, txtSECESS.KeyDown, txtSalesTaxCode.KeyDown, txtAEDCode.KeyDown, txtAddVAT.KeyDown, txtEmpName.KeyDown
        Try
            If e.KeyCode = Keys.F1 Then
                If sender Is txtExcise Then
                    BtnHelp_Click(BtnExciseHelp, New EventArgs())
                ElseIf sender Is txtCESS Then
                    BtnHelp_Click(btnHelpCess, New EventArgs())
                ElseIf sender Is txtSalesTaxCode Then
                    BtnHelp_Click(BtnHelpSalesTax, New EventArgs())
                ElseIf sender Is txtAEDCode Then
                    BtnHelp_Click(BtnHelpAED, New EventArgs())
                ElseIf sender Is txtSECESS Then
                    BtnHelp_Click(BtnHelpSECESS, New EventArgs())
                ElseIf sender Is txtSurcharge Then
                    BtnHelp_Click(BtnHelpSurcharge, New EventArgs())
                ElseIf sender Is txtAddVAT Then
                    BtnHelp_Click(BtnHelpAddVAT, New EventArgs())
                ElseIf (sender Is txtCustCode) Then
                    BtnHelpHdr_Click(BtnHelpCustCode, New EventArgs())
                ElseIf (sender Is txtEmpCode Or sender Is txtEmpName) Then
                    BtnHelpHdr_Click(BtnHelpEmp, New EventArgs())
                ElseIf (sender Is txtProvDocNo) Then
                    BtnHelpProvDocNo_Click(BtnHelpProvDocNo, New EventArgs())
                End If
            Else
                e.SuppressKeyPress = True

            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub fpSpreadRateWiseDtl_KeyPressEvent(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles fpSpreadRateWiseDtl.KeyPressEvent
        Try
            If e.keyAscii = 39 Then
                e.keyAscii = 0
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub BtnSubmitforAuth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSubmitforAuth.Click
        Dim sqlCmd As New SqlCommand
        Try
            If MsgBox("Are You Sure, you want to submit?", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                With sqlCmd
                    .CommandType = CommandType.StoredProcedure
                    .Connection = SqlConnectionclass.GetConnection()
                    .CommandText = "USP_SALES_PROV_Generation"
                    .Transaction = .Connection.BeginTransaction
                    .Parameters.Clear()
                    .Parameters.Add(New SqlParameter("@p_ProvDocNo", SqlDbType.VarChar, 20, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                    .Parameters("@p_ProvDocNo").Value = txtProvDocNo.Text.Trim
                    .Parameters.AddWithValue("@p_CreditNoteValue", txtNegativeVal.Text.Trim())
                    .Parameters.AddWithValue("@p_SupplInvoiceValue", txtPositiveVal.Text.Trim())

                    .Parameters.AddWithValue("@p_Excise", txtExciseVal.Text.Trim())
                    .Parameters.AddWithValue("@p_Cess", txtCESSVal.Text.Trim())
                    .Parameters.AddWithValue("@p_SalesTax", txtSalesTaxVal.Text.Trim())
                    .Parameters.AddWithValue("@p_AED", txtAEDVal.Text.Trim())
                    .Parameters.AddWithValue("@p_SEcess", txtSECESSVal.Text.Trim())
                    .Parameters.AddWithValue("@p_Surcharge", txtSurchargeVal.Text.Trim())
                    .Parameters.AddWithValue("@p_AddVAT", txtAddVATVal.Text.Trim())
                    .Parameters.AddWithValue("@p_TotalTaxableVal", txtTotAssVal.Text.Trim())
                    .Parameters.AddWithValue("@p_TotalNetVal", txtNetVal.Text.Trim())
                    .Parameters.AddWithValue("@p_RoundOff_Diff", txtRoundOffBy.Text.Trim())

                    .Parameters.AddWithValue("@p_Excise_CR", txtExciseVal_CR.Text.Trim())
                    .Parameters.AddWithValue("@p_Cess_CR", txtCESSVal_CR.Text.Trim())
                    .Parameters.AddWithValue("@p_SalesTax_CR", txtSalesTaxVal_CR.Text.Trim())
                    .Parameters.AddWithValue("@p_AED_CR", txtAEDVal_CR.Text.Trim())
                    .Parameters.AddWithValue("@p_SEcess_CR", txtSECESSVal_CR.Text.Trim())
                    .Parameters.AddWithValue("@p_Surcharge_CR", txtSurchargeVal_Neg.Text.Trim())
                    .Parameters.AddWithValue("@p_AddVAT_CR", txtAddVATVal_CR.Text.Trim())
                    .Parameters.AddWithValue("@p_TotalTaxableVal_CR", txtTotAssVal_Neg.Text.Trim())
                    .Parameters.AddWithValue("@p_TotalNetVal_CR", txtNetVal_neg.Text.Trim())
                    .Parameters.AddWithValue("@p_RoundOff_Diff_CR", txtRoundOffBy_Neg.Text.Trim())

                    .Parameters.AddWithValue("@p_TRANTYPE", "S")
                    .Parameters.AddWithValue("@p_UserId", mP_User)
                    .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                    .Parameters.AddWithValue("@p_IPAddress", gstrIpaddressWinSck)
                    .Parameters.Add(New SqlParameter("@p_ERROR", SqlDbType.VarChar, 200, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                    .ExecuteNonQuery()
                    If String.IsNullOrEmpty(Convert.ToString(sqlCmd.Parameters("@p_ERROR").Value)) Then
                        .Transaction.Commit()
                        MessageBox.Show("Sale Provision Doc No " + Convert.ToString(sqlCmd.Parameters("@p_ProvDocNo").Value) + " submitted successfully.")
                        InitializeForm(1)
                    Else
                        .Transaction.Rollback()
                        MessageBox.Show(Convert.ToString(sqlCmd.Parameters("@p_ERROR").Value), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    End If
                End With
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub dtp_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpFrm.ValueChanged, dtpToDt.ValueChanged
        Try
            If cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                ' InitializeForm(2)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

#End Region

#Region "Retrieve Docs"

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
            fdt = SqlConnectionclass.ExecuteReader("select DocData from Sales_Prov_DocList where cast(Prov_DocNo as varchar(20))='" + txtProvDocNo.Text.Trim() + "' and UNIT_CODE='" + gstrUNITID + "' and DocName='" + imavar1 + "'")
            fdt.Read()
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
            RaiseException(ex)
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
        Try
            If dgvDocList.ColumnCount <= 0 Then
                AddColumnDocListGrid()
            End If
            If dgvDocList.Rows.Count > 0 Then
                dgvDocList.DataSource = Nothing
            End If
            If IsNothing(dtDocTable) Then
                strQry = "select Prov_DocNo, Unit_Code, DocName, '' DocPath, DocExt, 'Show' Show, 'Remove' Remove from Sales_Prov_DocList where convert(varchar(20), Prov_DocNo)='" + txtProvDocNo.Text.Trim() + "' and Unit_Code='" + gstrUNITID + "'"
                dtDocTable = New DataTable
                dtDocTable = SqlConnectionclass.GetDataTable(strQry)
            End If
            dgvDocList.DataSource = dtDocTable
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub BtnUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnUpload.Click, btnDocFormExit.Click
        Dim strFileName As String = String.Empty
        Try
            If sender Is BtnUpload Then
                Dim odr As DataRow
                ofd_SelectDoc.Title = "Open File Dialog"
                ofd_SelectDoc.InitialDirectory = "C:\"
                ofd_SelectDoc.Filter = "All files (*.pdf)|*.pdf|All files (*.xls)|*.xls|All files (*.doc)|*.doc|All files (*.jpg)|*.jpg"
                ofd_SelectDoc.FilterIndex = 4
                ofd_SelectDoc.RestoreDirectory = True
                If ofd_SelectDoc.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    strFileName = ofd_SelectDoc.FileName
                    If dtDocTable.Select("DocName='" + Path.GetFileName(ofd_SelectDoc.FileName) + "'").Length > 0 Then
                        MessageBox.Show("File already exists.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Return
                    End If
                    odr = dtDocTable.NewRow
                    odr("DocName") = Path.GetFileName(ofd_SelectDoc.FileName)
                    odr("DocPath") = Path.GetFullPath(ofd_SelectDoc.FileName)
                    odr("DocExt") = Path.GetExtension(ofd_SelectDoc.FileName)
                    odr("Show") = "Show"
                    odr("Remove") = "Remove"
                    dtDocTable.Rows.Add(odr)
                    dgvDocList.DataSource = dtDocTable
                End If
            ElseIf sender Is btnDocFormExit Then
                If Not IsNothing(DocFrm) Then
                    DocFrm.Close()
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub BtnUploadDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnUploadDoc.Click
        Try
            If cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                BtnUpload.Enabled = False
            Else
                BtnUpload.Enabled = True
            End If
            DocFrm = New Form()
            DocFrm.StartPosition = FormStartPosition.CenterParent
            DocFrm.Width = 665
            DocFrm.Height = 260
            DocFrm.BackColor = Me.BackColor
            DocFrm.Text = "Upload Document"
            DocFrm.Controls.Add(DocUploadPanel)
            DocUploadPanel.Visible = True
            DocUploadPanel.Dock = DockStyle.Fill
            DocFrm.MaximizeBox = False
            DocFrm.MinimizeBox = False
            DocFrm.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
            DocFrm.ShowDialog()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub dgvDocList_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDocList.CellContentClick
        Try
            If e.ColumnIndex = dgvDocList.Columns("Show").Index Then
                If Not String.IsNullOrEmpty(Convert.ToString(dgvDocList.Rows(e.RowIndex).Cells("FilePath").Value)) Then
                    Process.Start(Convert.ToString(dgvDocList.Rows(e.RowIndex).Cells("FilePath").Value))
                Else
                    RetrieveFile(dgvDocList.Rows(e.RowIndex).Cells("FileName").Value, "", True)
                End If
            ElseIf e.ColumnIndex = dgvDocList.Columns("Remove").Index And Not cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                dtDocTable.Select("DocName='" + dgvDocList.Rows(e.RowIndex).Cells("FileName").Value + "'")(0).Delete()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

#End Region

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
            If mode = "EDIT" Then
                frmObj_PCD.g_ProvDoc_No = txtProvDocNo.Text.Trim.ToString
            End If
            frmObj_PCD.ShowDialog()
            If mode.ToString() = "ADD" Or mode.ToString() = "EDIT" Then
                fpSpreadRateWiseDtl.Col = RateWiseDtlGrid.Col_NewRate
                fpSpreadRateWiseDtl.Row = fpSpreadRateWiseDtl.ActiveRow
                fpSpreadRateWiseDtl.Value = frmObj_PCD.g_newrate
            End If
            frmObj_PCD.Dispose()
        Catch Ex As Exception
            RaiseException(Ex)
        End Try
    End Sub
    Private Sub chk_UploadDoc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            If cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If rbtn_MarutiFile.Checked = True Then
                    dtpFrm.Enabled = False
                    dtpToDt.Enabled = False
                    BtnHelpCustCode.Enabled = False
                    BtnFetch.Enabled = False
                    btn_fileName.Enabled = True
                    BtnHelpEmp.Enabled = False
                Else
                    rbtn_MarutiFile.Checked = False
                    dtpFrm.Enabled = True
                    dtpToDt.Enabled = True
                    BtnHelpCustCode.Enabled = True
                    BtnFetch.Enabled = False
                    btn_fileName.Enabled = False
                    BtnHelpEmp.Enabled = True
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub getdata(ByRef item_code As String)
        Try
            If rbtn_MarutiFile.Checked = True And cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then

                Dim strQry As String = String.Empty
                Dim sqlCmd As New SqlCommand()
                Dim odata As New SqlDataAdapter
                Dim odatatable As New DataSet

                With sqlCmd
                    .CommandText = "USP_SALES_PROV_Generation"
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 0
                    .Connection = SqlConnectionclass.GetConnection()
                    .Parameters.Clear()
                    If String.IsNullOrEmpty(txtProvDocNo.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_ProvDocNo", 0.0)
                    End If
                    .Parameters.AddWithValue("@p_CUSTOMERCODE", txtCustCode.Text.Trim.ToString())
                    .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                    .Parameters.AddWithValue("@p_IPAddress", gstrIpaddressWinSck)
                    .Parameters.AddWithValue("@p_UserId", mP_User)
                    .Parameters.AddWithValue("@p_FileName", txt_marutiFile.Text.Trim.ToString())
                    .Parameters.AddWithValue("@p_ItemCode", item_code)
                    .Parameters.AddWithValue("@p_TRANTYPE", "M")  'fill file Table
                    .Parameters.Add(New SqlParameter("@p_ERROR", SqlDbType.VarChar, 200, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                    odata.SelectCommand = sqlCmd
                    odata.Fill(odatatable)
                    .Dispose()
                    If Not String.IsNullOrEmpty(.Parameters("@p_ERROR").Value) Then
                        MessageBox.Show(Convert.ToString(sqlCmd.Parameters("@p_ERROR").Value), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    End If
                End With
                clearFarGrid()
                If odatatable.Tables.Count > 0 Then
                    If Not IsNothing(odatatable.Tables(0)) Then
                        dgvDispatchDtl.DataSource = odatatable.Tables(0)
                    End If
                    If Not IsNothing(odatatable.Tables(1)) Then
                        With fpSpreadRateWiseDtl
                            .MaxRows = 0
                            For Each row As DataRow In odatatable.Tables(1).Rows
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
                            GetTotalBasicValEffect()
                            CalculateTaxes()
                        End With
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub btn_MarutiUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_MarutiUpload.Click
        Try
            If cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then

                Dim strQuery As String = String.Empty
                Dim FileName As String = String.Empty
                Dim BATCHNO As String = String.Empty

                Dim fd As OpenFileDialog = New OpenFileDialog()

                If fd.ShowDialog() = DialogResult.OK Then

                    txtFilePath.Text = System.IO.Path.GetFileName(fd.FileName)
                    strQuery = ""

                    If txtFilePath.Text.Trim <> "" Then
                        
                        FileName = Convert.ToString(SqlConnectionclass.ExecuteScalar("SELECT ISNULL(FILENAME,'') FILENAME FROM MARUTI_FILE_HDR WHERE UNITCODE ='" & gstrUNITID & "' AND FILENAME ='" & txtFilePath.Text.Trim.ToString() & "'"))
                        If Not IsNothing(FileName) Then
                            If FileName = txtFilePath.Text.Trim.ToString() Then
                                MsgBox("Selected File Already Uploaded!!")
                                Exit Sub
                            End If
                        End If

                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)

                        mP_Connection.BeginTrans()

                        Using sr As StreamReader = New StreamReader(fd.FileName.ToString())
                            Dim line As String
                            Dim strOriLine As String
                            Dim strDataLine As String
                            Dim strVal() As String = String.Empty.Split(",")
                            Dim strsql As String
                            Dim batch_no As String = String.Empty
                            Dim batch As String = String.Empty
                            Dim sequence_no As String = String.Empty
                            Dim serial_no As Integer = 0

                            ' Read and display lines from the file until the end of  
                            ' the file is reached. q
                            While (Not sr.EndOfStream)
                                line = sr.ReadLine()
                                strOriLine = line
                                batch_no = line.Replace("BATCH NO  :", "§§")
                                i = batch_no.IndexOf("§§")
                                If (i > 0) Then
                                    batch = batch_no.Substring(i + 2, 6).Trim()
                                End If

                                sequence_no = line.Replace("SERIAL NO :", "§§§")
                                i = sequence_no.IndexOf("§§§")
                                If (i > 0) Then
                                    serial_no = Convert.ToInt32(sequence_no.Substring(i + 3, 4).Trim())
                                End If

                                line = line.Replace("Part No. :", "§").Replace("STX :", "»")
                                If (strVal.Length = 1) Then
                                    If line.Contains("§") = True And line.Contains("»") = True Then

                                        line = line.Replace(" ", ",").Replace(",,,,,,,,,,", ",").Replace(",,,,,,,,,", ",").Replace(",,,,,,,,", ",").Replace(",,,,,,,", ",").Replace(",,,,,,", ",").Replace(",,,,,", ",").Replace(",,,,", ",").Replace(",,,", ",").Replace(",,", ",")
                                        i = line.IndexOf("§")
                                        strPartNo = line.Substring(i + 2, 15)

                                        strDataLine = sr.ReadLine()
                                        strDataLine = sr.ReadLine()
                                        strDataLine = sr.ReadLine()
                                        strDataLine = sr.ReadLine()

                                        strDataLine = strDataLine.Replace(" ", ",").Replace(",,,,,,,,,,", ",").Replace(",,,,,,,,,", ",").Replace(",,,,,,,,", ",").Replace(",,,,,,,", ",").Replace(",,,,,,", ",").Replace(",,,,,", ",").Replace(",,,,", ",").Replace(",,,", ",").Replace(",,", ",")
                                        strVal = strDataLine.Split(",")

                                        strBillNo = Convert.ToInt32(strVal(2))
                                        strPricDt = strVal(3).ToString()
                                        strshp = Convert.ToInt16(strVal(6))
                                        stracp = Convert.ToInt16(strVal(7))
                                        strOldRate = Convert.ToDecimal(strVal(8))
                                        strBilldate = strVal(15).ToString()
                                        strNewRate = Convert.ToDecimal(strVal(19))

                                        strsql = ""
                                        strsql = "Insert Into TMP_SalesFileData (strPartNo,strBillNo,strPricDt,strshp,stracp,strOldRate,strNewRate,Unit_code,Ipaddress,invoiceNo,batch_no,serial_no,strBilldate) Values ('" & strPartNo & "','" & strBillNo & "','" & strPricDt & "','" & strshp & "','" & stracp & "','" & strOldRate & "','" & strNewRate & "','" & gstrUNITID & "','" & gstrIpaddressWinSck & "',CONVERT(INTEGER,(Substring(convert(varchar(20)," & strBillNo & "),1,2)+Substring(convert(varchar(20)," & strBillNo & "),4,10))),'" & batch & "'," & serial_no & ",'" & strBilldate & "')"

                                        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                                    End If
                                Else
                                    strDataLine = line
                                    strDataLine = strDataLine.Replace(" ", ",").Replace(",,,,,,,,,,", ",").Replace(",,,,,,,,,", ",").Replace(",,,,,,,,", ",").Replace(",,,,,,,", ",").Replace(",,,,,,", ",").Replace(",,,,,", ",").Replace(",,,,", ",").Replace(",,,", ",").Replace(",,", ",")
                                    strVal = strDataLine.Split(",")

                                    If strVal.Length = 23 Then

                                        strBillNo = Convert.ToInt32(strVal(2))
                                        strPricDt = strVal(3).ToString()
                                        strshp = Convert.ToInt16(strVal(6))
                                        stracp = Convert.ToInt16(strVal(7))
                                        strOldRate = Convert.ToDecimal(strVal(8))
                                        strBilldate = strVal(15).ToString()
                                        strNewRate = Convert.ToDecimal(strVal(19))
                                        strsql = ""
                                        strsql = "Insert Into TMP_SalesFileData (strPartNo,strBillNo,strPricDt,strshp,stracp,strOldRate,strNewRate,Unit_code,Ipaddress,invoiceNo,batch_no,serial_no,strBilldate) Values ('" & strPartNo & "','" & strBillNo & "','" & strPricDt & "','" & strshp & "','" & stracp & "','" & strOldRate & "','" & strNewRate & "','" & gstrUNITID & "','" & gstrIpaddressWinSck & "',CONVERT(INTEGER,(Substring(convert(varchar(20)," & strBillNo & "),1,2)+Substring(convert(varchar(20)," & strBillNo & "),4,10))),'" & batch & "'," & serial_no & ",'" & strBilldate & "')"

                                        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                                    End If

                                    If strVal.Length <> 23 Then
                                        strVal = String.Empty.Split(",")
                                    End If

                                End If
                            End While
                            strsql = ""
                            strsql = "UPDATE A SET A.ACCOUNT_CODE=B.ACCOUNT_CODE FROM TMP_SALESFILEDATA A INNER JOIN SALESCHALLAN_DTL (NOLOCK) B ON A.UNIT_CODE=B.UNIT_CODE AND A.INVOICENO =B.DOC_NO WHERE A.IPADDRESS ='" & gstrIpaddressWinSck & "' AND A.UNIT_CODE='" & gstrUNITID & "'"
                            mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                            strsql = ""
                            strsql = "UPDATE A SET A.ITEM_CODE =B.ITEM_CODE FROM TMP_SALESFILEDATA A INNER JOIN SALES_DTL (NOLOCK) B ON A.UNIT_CODE=B.UNIT_CODE AND A.INVOICENO =B.DOC_NO AND A.STRPARTNO=B.CUST_ITEM_CODE WHERE A.IPADDRESS ='" & gstrIpaddressWinSck & "' AND A.UNIT_CODE='" & gstrUNITID & "'"
                            mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                            strsql = ""
                            BATCHNO = Convert.ToString(SqlConnectionclass.ExecuteScalar("SELECT ISNULL(BATCH_NO,'') BATCH_NO FROM MARUTI_FILE_HDR WHERE UNITCODE ='" & gstrUNITID & "' AND BATCH_NO ='" & batch & "'"))
                            If Not IsNothing(BATCHNO) Then
                                If BATCHNO = batch Then
                                    MsgBox("Selected File Already Uploaded!!")
                                    mP_Connection.RollbackTrans()
                                    Exit Sub
                                End If
                            End If

                            strsql = ""
                            strsql = "INSERT INTO MARUTI_FILE_HDR (FILENAME,UNITCODE,BATCH_NO,ENT_DT) Values ('" & txtFilePath.Text.Trim.ToString() & "','" & gstrUNITID & "','" & batch & "',getdate())"
                            mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                            mP_Connection.CommitTrans()

                            txtFilePath.Text = ""

                        End Using
                        txtCustCode.Text = Convert.ToString(SqlConnectionclass.ExecuteScalar("SELECT DISTINCT TOP(1) Account_Code  FROM TMP_SALESFILEDATA WHERE IPADDRESS ='" & gstrIpaddressWinSck & "' AND UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE IS NOT NULL"))
                        MessageBox.Show("Uploading Completed!! For " + txtCustCode.Text.Trim.ToString())
                        txtCustCode.Text = ""
                    Else

                    End If
                End If

            Else

            End If

        Catch ex As Exception
            mP_Connection.RollbackTrans()
            RaiseException(ex)
        End Try
    End Sub
    Private Sub rbtn_NormalInvoice_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtn_NormalInvoice.CheckedChanged
        Try
            If rbtn_NormalInvoice.Checked = True Then
                txtEmpCode.Enabled = True
                txtEmpName.Enabled = True
                txtCustCode.Enabled = True
                BtnHelpCustCode.Enabled = True
                BtnHelpEmp.Enabled = True
                dtpFrm.Enabled = True
                dtpToDt.Enabled = True
                BtnFetch.Enabled = True
                btn_fileName.Enabled = False
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub rbtn_MarutiFile_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtn_MarutiFile.CheckedChanged
        Try
            If rbtn_MarutiFile.Checked = True Then
                txtEmpCode.Enabled = False
                txtEmpName.Enabled = False
                txtCustCode.Enabled = False
                BtnHelpCustCode.Enabled = False
                BtnHelpEmp.Enabled = False
                dtpFrm.Enabled = False
                dtpToDt.Enabled = False
                btn_fileName.Enabled = True

            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub btn_fileName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_fileName.Click
        Try
            If cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If rbtn_MarutiFile.Checked = True Then
                    Dim strQry As String = String.Empty
                    Dim strhelp As String()
                    Dim odata As DataTable
                    With ctlHelp
                        .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                        .ConnectAsUser = gstrCONNECTIONUSER
                        .ConnectThroughDSN = gstrCONNECTIONDSN
                        .ConnectWithPWD = gstrCONNECTIONPASSWORD
                    End With
                    strQry = "SELECT FILENAME AS FILENAME, BATCH_NO AS BATCH,CONVERT(VARCHAR(20),ENT_DT,106) AS UPLOADDATE FROM MARUTI_FILE_HDR  WHERE UNITCODE='" + gstrUNITID + "'"
                    strhelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry, "Select File")
                    If Not IsNothing(strhelp) Then
                        If Not IsNothing(strhelp(1)) Then
                            txt_marutiFile.Text = strhelp(0).Trim()
                            txtEmpCode.Enabled = True
                            txtEmpName.Enabled = True
                            BtnHelpEmp.Enabled = True
                            txtCustCode.Text = Convert.ToString(SqlConnectionclass.ExecuteScalar("SELECT DISTINCT TOP(1) ACCOUNT_CODE  FROM TMP_SALESFILEDATA WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE IS NOT NULL AND BATCH_NO IN (SELECT BATCH_NO FROM MARUTI_FILE_HDR WHERE UNITCODE ='" & gstrUNITID & "' AND FILENAME ='" & strhelp(0).Trim() & "')"))
                            strQry = ""
                            strQry = "SELECT DISTINCT CM.CUSTOMER_CODE, CM.CUST_NAME [Customer_Name], CM.KAMCODE [KAM_Code], EM.Name [KAM_Name] FROM CUSTOMER_MST CM left join Employee_mst EM on CM.UNIT_CODE=EM.UNIT_CODE and CM.KAMCODE=EM.Employee_code WHERE CM.UNIT_CODE='" & gstrUNITID & "' AND CM.CUSTOMER_CODE ='" & txtCustCode.Text.Trim.ToString() & "'"
                            odata = SqlConnectionclass.GetDataTable(strQry)
                            If odata.Rows.Count > 0 Then
                                txtCustCode.Text = odata.Rows(0)(0).ToString()
                                lblCustDesc.Text = odata.Rows(0)(1).ToString()
                                lblKAMName.Text = odata.Rows(0)(3).ToString()
                                lblKAMName.Tag = odata.Rows(0)(2).ToString()
                                BtnFetch.Enabled = True
                                rbtn_NormalInvoice.Enabled = False
                                ClearDataGridView(dgvDispatchDtl)
                                clearFarGrid()
                                ClearAllTaxes()
                            End If
                        Else
                            MsgBox("No Files Available to Select!!")
                        End If

                    Else
                        MsgBox("No Files Available to Select!!")
                    End If
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Function SaveFileData() As Boolean
        Try
            Dim SqlCmd As New SqlCommand
            Dim strQuery As String = String.Empty
            Dim fs As FileStream
            Dim rawData As Object
            Dim odt As DataTable
            Dim strProvDocNo As String = String.Empty

            If ValidateSave() Then
               

                If cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then

                    SqlCmd = New SqlCommand
                   

                    strQuery = " UPDATE TMP_SALESFILEDATA SET CHANGEREASON=@CHANGEREASON, NATUREOFCORRECTION=@NATUREOFCORRECTION, IsSelected=1" & _
                               " FROM TMP_SALESFILEDATA TMP INNER JOIN MARUTI_FILE_HDR HDR ON TMP.UNIT_CODE=HDR.UNITCODE AND TMP.BATCH_NO=HDR.BATCH_NO" & _
                               " WHERE TMP.ACCOUNT_CODE=@CUSTOMER_CODE AND TMP.ITEM_CODE=@ITEM_CODE AND TMP.STROLDRATE=@RATE AND TMP.STRNEWRATE=@NEWRATE AND" & _
                               " HDR.UNITCODE=@UNIT_CODE AND HDR.FILENAME=@P_FILENAME"

                    With SqlCmd
                        .Connection = SqlConnectionclass.GetConnection()
                        .CommandText = strQuery
                        .CommandType = CommandType.Text
                        .CommandTimeout = 0
                    End With

                    With fpSpreadRateWiseDtl
                        For Count As Integer = 1 To .MaxRows
                            SqlCmd.Parameters.Clear()
                            .Row = Count
                            .Col = RateWiseDtlGrid.Col_Select
                            If .Value = True Then

                                .Col = RateWiseDtlGrid.Col_NewRate
                                SqlCmd.Parameters.AddWithValue("@NEWRATE", Convert.ToString(.Value))
                                .Col = RateWiseDtlGrid.Col_ReasonChange
                                SqlCmd.Parameters.AddWithValue("@CHANGEREASON", Convert.ToString(.Value))
                                .Col = RateWiseDtlGrid.Col_CorrectionNature
                                SqlCmd.Parameters.AddWithValue("@NATUREOFCORRECTION", .CellTag)
                                .Col = RateWiseDtlGrid.Col_Part_Code
                                SqlCmd.Parameters.AddWithValue("@ITEM_CODE", Convert.ToString(.Value))
                                .Col = RateWiseDtlGrid.Col_InvRate
                                SqlCmd.Parameters.AddWithValue("@RATE", Convert.ToDecimal(.Value).ToString("0.00"))

                                SqlCmd.Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                                SqlCmd.Parameters.AddWithValue("@CUSTOMER_CODE", txtCustCode.Text.Trim())
                                SqlCmd.Parameters.AddWithValue("@P_FILENAME", txt_marutiFile.Text.Trim())

                                SqlCmd.ExecuteNonQuery()
                            End If
                        Next
                    End With

                End If


                If cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                    SqlCmd = New SqlCommand

                   
                    strQuery = " UPDATE SALES_PROV_RATEWISEDETAIL SET CHANGEREASON=@CHANGEREASON, NATUREOFCORRECTION=@NATUREOFCORRECTION, UPD_DT=GETDATE(), UPD_USERID =@UPD_USERID" & _
                          " WHERE UNIT_CODE=@UNIT_CODE AND PROV_DOCNO=@p_ProvDocNo AND RATE=@RATE AND NEWRATE=@NEWRATE AND CUSTOMER_CODE=@CUSTOMER_CODE AND ITEM_CODE=@ITEM_CODE"

                    With SqlCmd
                        .Connection = SqlConnectionclass.GetConnection()
                        .CommandText = strQuery
                        .CommandType = CommandType.Text
                        .CommandTimeout = 0
                    End With

                    With fpSpreadRateWiseDtl
                        For Count As Integer = 1 To .MaxRows
                            SqlCmd.Parameters.Clear()
                            .Row = Count
                            .Col = RateWiseDtlGrid.Col_Select
                            'If .Value = True Then

                            .Col = RateWiseDtlGrid.Col_NewRate
                            SqlCmd.Parameters.AddWithValue("@NEWRATE", Convert.ToString(.Value))
                            .Col = RateWiseDtlGrid.Col_ReasonChange
                            SqlCmd.Parameters.AddWithValue("@CHANGEREASON", Convert.ToString(.Value))
                            .Col = RateWiseDtlGrid.Col_CorrectionNature
                            SqlCmd.Parameters.AddWithValue("@NATUREOFCORRECTION", .CellTag)
                            .Col = RateWiseDtlGrid.Col_Part_Code
                            SqlCmd.Parameters.AddWithValue("@ITEM_CODE", Convert.ToString(.Value))
                            .Col = RateWiseDtlGrid.Col_InvRate
                            SqlCmd.Parameters.AddWithValue("@RATE", Convert.ToDecimal(.Value).ToString("0.00"))

                            SqlCmd.Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                            SqlCmd.Parameters.AddWithValue("@CUSTOMER_CODE", txtCustCode.Text.Trim())
                            SqlCmd.Parameters.AddWithValue("@UPD_USERID", gstrUserIDSelected)
                            SqlCmd.Parameters.AddWithValue("@p_ProvDocNo", txtProvDocNo.Text.Trim())

                            SqlCmd.ExecuteNonQuery()
                            'End If
                        Next
                    End With
                End If

                If cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then

                    SqlCmd = New SqlCommand

                   
                    strQuery = " UPDATE SALES_PROV_PARTINVOICEDETAIL SET CHANGEREASON=@CHANGEREASON, NATUREOFCORRECTION=@NATUREOFCORRECTION, UPD_DT=GETDATE(), UPD_USERID =@UPD_USERID" & _
                          " WHERE UNIT_CODE=@UNIT_CODE AND PROV_DOCNO=@p_ProvDocNo AND RATE=@RATE AND NEWRATE=@NEWRATE AND CUSTOMER_CODE=@CUSTOMER_CODE AND ITEM_CODE=@ITEM_CODE"

                    With SqlCmd
                        .Connection = SqlConnectionclass.GetConnection()
                        .CommandType = CommandType.Text
                        .CommandText = strQuery
                        .CommandTimeout = 0
                    End With

                    With fpSpreadRateWiseDtl
                        For Count As Integer = 1 To .MaxRows
                            SqlCmd.Parameters.Clear()
                            .Row = Count
                            .Col = RateWiseDtlGrid.Col_Select
                            'If .Value = True Then

                            .Col = RateWiseDtlGrid.Col_NewRate
                            SqlCmd.Parameters.AddWithValue("@NEWRATE", Convert.ToString(.Value))
                            .Col = RateWiseDtlGrid.Col_ReasonChange
                            SqlCmd.Parameters.AddWithValue("@CHANGEREASON", Convert.ToString(.Value))
                            .Col = RateWiseDtlGrid.Col_CorrectionNature
                            SqlCmd.Parameters.AddWithValue("@NATUREOFCORRECTION", .CellTag)
                            .Col = RateWiseDtlGrid.Col_Part_Code
                            SqlCmd.Parameters.AddWithValue("@ITEM_CODE", Convert.ToString(.Value))
                            .Col = RateWiseDtlGrid.Col_InvRate
                            SqlCmd.Parameters.AddWithValue("@RATE", Convert.ToDecimal(.Value).ToString("0.00"))

                            SqlCmd.Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                            SqlCmd.Parameters.AddWithValue("@CUSTOMER_CODE", txtCustCode.Text.Trim())
                            SqlCmd.Parameters.AddWithValue("@UPD_USERID", gstrUserIDSelected)
                            SqlCmd.Parameters.AddWithValue("@p_ProvDocNo", txtProvDocNo.Text.Trim())

                            SqlCmd.ExecuteNonQuery()
                            'End If
                        Next
                    End With
                End If

                With SqlCmd
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "USP_SALES_PROV_Generation"
                    .Transaction = .Connection.BeginTransaction
                    .Parameters.Clear()
                    If String.IsNullOrEmpty(txtProvDocNo.Text.Trim) Then
                        .Parameters.Add(New SqlParameter("@p_ProvDocNo", SqlDbType.VarChar, 20, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                        .Parameters("@p_ProvDocNo").Value = 0
                        .Parameters.AddWithValue("@p_TRANTYPE", "FA")
                    Else
                        .Parameters.Add(New SqlParameter("@p_ProvDocNo", SqlDbType.VarChar, 20, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                        .Parameters("@p_ProvDocNo").Value = txtProvDocNo.Text.Trim
                        .Parameters.AddWithValue("@p_TRANTYPE", "FE")
                    End If

                    .Parameters.AddWithValue("@p_PERSONTO_AUTH", txtEmpCode.Text.Trim())
                    .Parameters.AddWithValue("@p_KAMCODE", Convert.ToString(lblKAMName.Tag))
                    .Parameters.AddWithValue("@p_ExciseCode", txtExcise.Text.Trim())
                    If String.IsNullOrEmpty(txtPerExcise.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_Excise", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_Excise", txtPerExcise.Text.Trim())
                    End If
                    .Parameters.AddWithValue("@p_CessCode", txtCESS.Text.Trim())
                    If String.IsNullOrEmpty(txtPerCESS.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_Cess", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_Cess", txtPerCESS.Text.Trim())
                    End If
                    .Parameters.AddWithValue("@p_SalesTaxCode", txtSalesTaxCode.Text.Trim())
                    If String.IsNullOrEmpty(txtPerSalesTax.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_SalesTax", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_SalesTax", txtPerSalesTax.Text.Trim())
                    End If

                    .Parameters.AddWithValue("@p_AEDCode", txtAEDCode.Text.Trim())
                    If String.IsNullOrEmpty(txtPerAED.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_AED", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_AED", txtPerAED.Text.Trim())
                    End If

                    .Parameters.AddWithValue("@p_SEcessCode", txtSECESS.Text.Trim())
                    If String.IsNullOrEmpty(txtPerSECESS.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_SEcess", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_SEcess", txtPerSECESS.Text.Trim())
                    End If

                    .Parameters.AddWithValue("@p_SurchargeCode", txtSurcharge.Text.Trim())
                    If String.IsNullOrEmpty(txtPerSurcharge.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_Surcharge", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_Surcharge", txtPerSurcharge.Text.Trim())
                    End If

                    .Parameters.AddWithValue("@p_AddVATCode", txtAddVAT.Text.Trim())
                    If String.IsNullOrEmpty(txtPerAddVAT.Text.Trim()) Then
                        .Parameters.AddWithValue("@p_AddVAT", "0.0")
                    Else
                        .Parameters.AddWithValue("@p_AddVAT", txtPerAddVAT.Text.Trim())
                    End If

                    With fpSpreadRateWiseDtl
                        .Row = 1
                        .Col = RateWiseDtlGrid.Col_Part_Code
                        SqlCmd.Parameters.AddWithValue("@p_ItemCode", Convert.ToString(.Value))
                    End With

                    .Parameters.AddWithValue("@P_FILENAME", txt_marutiFile.Text.Trim())
                    .Parameters.AddWithValue("@p_UserId", mP_User)
                    .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                    .Parameters.AddWithValue("@p_CUSTOMERCODE", txtCustCode.Text.Trim())
                    .Parameters.AddWithValue("@p_IPAddress", gstrIpaddressWinSck)
                    .Parameters.Add(New SqlParameter("@p_ERROR", SqlDbType.VarChar, 200, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                    .ExecuteNonQuery()


                    If String.IsNullOrEmpty(Convert.ToString(SqlCmd.Parameters("@p_ERROR").Value)) Then
                        If String.IsNullOrEmpty(txtProvDocNo.Text.Trim()) Then
                            strProvDocNo = Convert.ToString(SqlCmd.Parameters("@p_ProvDocNo").Value)
                        Else
                            strProvDocNo = txtProvDocNo.Text.Trim()
                        End If
                        strQuery = "SELECT DOCNAME, CAST(0 AS BIT) ISSAVED FROM SALES_PROV_DOCLIST WHERE CAST(PROV_DOCNO AS VARCHAR(20))='" + txtProvDocNo.Text.Trim() + "' AND UNIT_CODE='" + gstrUNITID + "'"
                        odt = SqlConnectionclass.GetDataTable(strQuery)
                        If Not IsNothing(dtDocTable) Then
                            If dtDocTable.Rows.Count > 0 Then
                                strQuery = " DELETE FROM SALES_PROV_DOCLIST WHERE CAST(PROV_DOCNO AS VARCHAR(20))=@PROV_DOCNO AND UNIT_CODE=@UNIT_CODE AND DOCNAME=@DOCNAME" & _
                                           " INSERT INTO SALES_PROV_DOCLIST(PROV_DOCNO, UNIT_CODE, DOCNAME, DOCDATA, DOCEXT)" & _
                                           " SELECT @PROV_DOCNO, @UNIT_CODE, @DOCNAME, @DOCDATA, @DOCEXT "
                                .CommandText = strQuery
                                .CommandType = CommandType.Text
                                For Each odr As DataRow In dtDocTable.Rows
                                    If Not String.IsNullOrEmpty(Convert.ToString(odr("DocPath"))) Then
                                        fs = New FileStream(Convert.ToString(odr("DocPath")), FileMode.Open, FileAccess.Read)
                                        rawData = New Byte(fs.Length) {}
                                        fs.Read(rawData, 0, fs.Length)
                                        fs.Close()
                                        .Parameters.Clear()
                                        .Parameters.AddWithValue("@PROV_DOCNO", strProvDocNo)
                                        .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                                        .Parameters.AddWithValue("@DOCNAME", Convert.ToString(odr("DocName")))
                                        .Parameters.AddWithValue("@DOCDATA", rawData)
                                        .Parameters.AddWithValue("@DOCEXT", Convert.ToString(odr("DocExt")))
                                        .ExecuteNonQuery()
                                    End If
                                    If odt.Select("DocName='" + Convert.ToString(odr("DOCName")) + "'").Length > 0 Then
                                        odt.Select("DocName='" + Convert.ToString(odr("DOCName")) + "'")(0)("ISSAVED") = True
                                    End If
                                Next
                                .CommandText = " DELETE FROM SALES_PROV_DOCLIST WHERE CAST(PROV_DOCNO AS VARCHAR(20))=@PROV_DOCNO AND UNIT_CODE=@UNIT_CODE AND DOCNAME=@DOCNAME"
                                .CommandType = CommandType.Text
                                For Each odr As DataRow In odt.Select("ISSAVED=0")
                                    .Parameters.Clear()
                                    .Parameters.AddWithValue("@PROV_DOCNO", strProvDocNo)
                                    .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                                    .Parameters.AddWithValue("@DOCNAME", Convert.ToString(odr("DocName")))
                                    .ExecuteNonQuery()
                                Next
                            End If
                        ElseIf odt.Rows.Count > 0 Then
                            SqlConnectionclass.ExecuteNonQuery("DELETE FROM SALES_PROV_DOCLIST WHERE CAST(PROV_DOCNO AS VARCHAR(20))='" + strProvDocNo + "' AND UNIT_CODE='" + gstrUNITID + "'")
                        End If
                        .Transaction.Commit()
                        If String.IsNullOrEmpty(txtProvDocNo.Text.Trim()) Then
                            txtProvDocNo.Text = strProvDocNo
                            MessageBox.Show("Sale Provision Doc No " + txtProvDocNo.Text.Trim() + " created successfully.")
                        Else
                            MessageBox.Show("Record updated successfully.")
                        End If
                        Return True
                    Else
                        .Transaction.Rollback()
                        MessageBox.Show(Convert.ToString(SqlCmd.Parameters("@p_ERROR").Value), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Return False
                    End If
                End With
            Else
                Return False


            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function
    Private Sub getdetaildata(ByRef Prov_No As String)
        Try

            Dim strQry As String = String.Empty
            Dim sqlCmd As New SqlCommand()
            Dim odata As New SqlDataAdapter
            Dim odatatable As New DataSet

            With sqlCmd
                .CommandText = "USP_SALES_PROV_Generation"
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .Connection = SqlConnectionclass.GetConnection()

                .Parameters.Clear()
                .Parameters.AddWithValue("@p_ProvDocNo", Prov_No)
                .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                .Parameters.AddWithValue("@p_UserId", mP_User)
                .Parameters.AddWithValue("@p_IPAddress", gstrIpaddressWinSck)
                .Parameters.AddWithValue("@p_TRANTYPE", "VIEWMARUTI")

                .Parameters.Add(New SqlParameter("@p_ERROR", SqlDbType.VarChar, 200, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                odata.SelectCommand = sqlCmd
                odata.Fill(odatatable)
                .Dispose()
                If Not String.IsNullOrEmpty(.Parameters("@p_ERROR").Value) Then
                    MessageBox.Show(Convert.ToString(sqlCmd.Parameters("@p_ERROR").Value), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                End If
            End With
            clearFarGrid()
            If odatatable.Tables.Count > 0 Then
                If Not IsNothing(odatatable.Tables(0)) Then
                    dgvDispatchDtl.DataSource = odatatable.Tables(0)
                End If
                If Not IsNothing(odatatable.Tables(1)) Then
                    With fpSpreadRateWiseDtl
                        .MaxRows = 0
                        For Each row As DataRow In odatatable.Tables(1).Rows
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
                        GetTotalBasicValEffect()
                        CalculateTaxes()
                        rbtn_MarutiFile.Checked = True
                    End With
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    ' Code Added By Mayur Against issue ID 10816097    
    ' Negative Taxes
    Private Sub BtnHelp_CR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ExciseCode_CR.Click, btn_AEDCODE_CR.Click, btn_CESS_CR.Click, btn_SECESS_CR.Click, btn_ST_CR.Click, btn_VAT_CR.Click, BtnHelpSurcharge_Neg.Click
        Dim strQry As String = String.Empty
        Dim strhelp As String()
        Try
            If Not cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                If txtNegativeVal.Text.Trim <> "" Then
                    With ctlHelp
                        .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                        .ConnectAsUser = gstrCONNECTIONUSER
                        .ConnectThroughDSN = gstrCONNECTIONDSN
                        .ConnectWithPWD = gstrCONNECTIONPASSWORD
                    End With
                    If sender Is btn_ExciseCode_CR Then
                        strQry = " SELECT DISTINCT TXRT_RATE_NO,CAST(TXRT_PERCENTAGE AS NUMERIC(9,2)) TXRT_PERCENTAGE FROM GEN_TAXRATE " & _
                             " WHERE UNIT_CODE='" + gstrUNITID + "' AND TX_TAXEID ='EXC' AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) OR (CAST(GETDATE() AS DATE)<= DEACTIVE_DATE))"
                        strhelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry, "Select Excise")
                        If Not IsNothing(strhelp) Then
                            If Not IsNothing(strhelp(1)) Then
                                txtExcise_CR.Text = strhelp(0).Trim()
                                txtExcisePER_CR.Text = strhelp(1).Trim()
                            Else
                                txtExcise_CR.Text = ""
                                txtExcisePER_CR.Text = ""
                            End If
                        End If
                    ElseIf sender Is btn_CESS_CR Then
                        strQry = "Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where Unit_code='" & gstrUNITID & "' and tx_TaxeID ='ECS' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                        strhelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry, "Select CESS")
                        If Not IsNothing(strhelp) Then

                            If Not IsNothing(strhelp(1)) Then
                                txtCESS_CR.Text = strhelp(0).Trim()
                                txtPerCESS_CR.Text = strhelp(1).Trim()
                            Else
                                txtCESS_CR.Text = ""
                                txtPerCESS_CR.Text = ""
                            End If

                        End If
                    ElseIf sender Is btn_ST_CR Then
                        strQry = "Select TxRT_Rate_No,TxRt_Percentage from Gen_taxRate where Unit_code='" & gstrUNITID & "' and (Tx_TaxeID ='CST' OR Tx_TaxeID ='LST' OR Tx_TaxeID ='VAT') and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                        strhelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry, "Sales Tax Help")
                        If Not IsNothing(strhelp) Then
                            If Not IsNothing(strhelp(1)) Then
                                txtSalesTaxCode_CR.Text = strhelp(0).Trim()
                                txtPerSalesTax_CR.Text = strhelp(1).Trim()
                            Else
                                txtSalesTaxCode_CR.Text = ""
                                txtPerSalesTax_CR.Text = ""
                            End If

                        End If
                    ElseIf sender Is btn_AEDCODE_CR Then
                        strQry = " SELECT DISTINCT TXRT_RATE_NO,CAST(TXRT_PERCENTAGE AS NUMERIC(9,2)) TXRT_PERCENTAGE FROM GEN_TAXRATE " & _
                             " WHERE UNIT_CODE='" + gstrUNITID + "' AND TX_TAXEID ='AED' AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) OR (CAST(GETDATE() AS DATE)<= DEACTIVE_DATE))"
                        strhelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry, "Select AED")
                        If Not IsNothing(strhelp) Then
                            If IsNothing(strhelp(1)) Then
                                txtAEDCode_CR.Text = ""
                                txtPerAED_CR.Text = ""
                            Else
                                txtAEDCode_CR.Text = strhelp(0).Trim()
                                txtPerAED_CR.Text = strhelp(1).Trim()
                            End If
                        End If
                    ElseIf sender Is btn_SECESS_CR Then
                        strQry = "Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where Unit_code='" & gstrUNITID & "' and tx_TaxeID ='ECSSH' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                        strhelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry, "SECESS  Help")
                        If Not IsNothing(strhelp) Then
                            If Not IsNothing(strhelp(1)) Then
                                txtSECESS_CR.Text = strhelp(0).Trim()
                                txtPerSECESS_CR.Text = strhelp(1).Trim()
                            Else
                                txtSECESS_CR.Text = ""
                                txtPerSECESS_CR.Text = ""
                            End If
                        End If
                    ElseIf sender Is BtnHelpSurcharge_Neg Then
                        strQry = "Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where Unit_code='" & gstrUNITID & "' and tx_TaxeID ='SST' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                        strhelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry, "Surcharge Help")
                        If Not IsNothing(strhelp) Then
                            If Not IsNothing(strhelp(1)) Then
                                txtSurcharge_Neg.Text = strhelp(0).Trim()
                                txtPerSurcharge_neg.Text = strhelp(1).Trim()
                            Else
                                txtSurcharge_Neg.Text = ""
                                txtPerSurcharge_neg.Text = ""
                            End If
                        End If
                    ElseIf sender Is btn_VAT_CR Then
                        strQry = "Select TxRT_Rate_No,TxRt_Percentage from Gen_taxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  Tx_TaxeID in('ADVAT','ADCST')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                        strhelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry, "ADD VAT")
                        If Not IsNothing(strhelp) Then
                            If Not IsNothing(strhelp(1)) Then
                                txtAddVAT_CR.Text = strhelp(0).Trim()
                                txtAddPerVAT_CR.Text = strhelp(1).Trim()
                            Else
                                txtAddVAT_CR.Text = ""
                                txtAddPerVAT_CR.Text = ""
                            End If
                        End If
                    End If
                End If
                CalculateTaxes()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub txtHelp_CR_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtExcise_CR.KeyDown, txtAEDCode_CR.KeyDown, txtCESS_CR.KeyDown, txtSECESS_CR.KeyDown, txtSalesTaxCode_CR.KeyDown, txtAddVAT_CR.KeyDown, txtSurcharge_Neg.KeyDown
        Try
            If e.KeyCode = Keys.F1 Then
                If sender Is txtExcise_CR Then
                    BtnHelp_CR_Click(BtnExciseHelp, New EventArgs())
                ElseIf sender Is txtCESS_CR Then
                    BtnHelp_CR_Click(btnHelpCess, New EventArgs())
                ElseIf sender Is txtSalesTaxCode_CR Then
                    BtnHelp_CR_Click(BtnHelpSalesTax, New EventArgs())
                ElseIf sender Is txtAEDCode_CR Then
                    BtnHelp_CR_Click(BtnHelpAED, New EventArgs())
                ElseIf sender Is txtSurcharge_Neg Then
                    BtnHelp_CR_Click(BtnHelpSurcharge, New EventArgs())
                ElseIf sender Is txtAddVAT_CR Then
                    BtnHelp_CR_Click(BtnHelpAddVAT, New EventArgs())
                ElseIf (sender Is txtSECESS_CR) Then
                    BtnHelp_CR_Click(BtnHelpCustCode, New EventArgs())
                End If
            Else
                e.SuppressKeyPress = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
End Class