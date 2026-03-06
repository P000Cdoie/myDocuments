Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports VB = Microsoft.VisualBasic
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine


Public Class FRMMKTTRN0098
    '========================================================================================
    'COPYRIGHT          :   MOTHERSONSUMI INFOTECH & DESIGN LTD.
    'AUTHOR             :   PRASHANT RAJPAL
    'CREATION DATE      :   14- SEP 2017- 28-SEP-2017
    'DESCRIPTION        :   SUPPLEMENTARY INVOICE FORM 
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
    Public pdblqty As Double
    Private strSelInvoices As String 'to store the selected invoice numbers
    Private strNotSelInvoices As String ' to store the deselected invoice numbers
    Dim bool_check As Boolean = False
    Dim pstrChallanNo As String
    'Dim intUNLOCKED_INVOICES As Integer
    '======================RoundoffSetting=====================
    Dim mstrsalesorder As String
    Enum RateWiseDtlGrid
        Col_Part_Code = 1
        Col_NewInvRate = 2
        Col_TotEffVal = 3
    End Enum
    Private Enum ENUM_Gridinvoicedetails
        BASE_INVOICENO = 1
        CURR_INVOICENO
        InternalPartNo
        CustPartNo
        ItemDesc
        CustPartDesc
        SalesOrder
        Amendmentno
        InvoiceQty
        Rate
        NewRate
        TaxableAmount
        HSNSACCODE
        CGST_TAX_TYPE
        CGST_TAX
        SGST_TAX_TYPE
        SGST_TAX
        IGST_TAX_TYPE
        IGST_TAX

    End Enum
#End Region

#Region "Form Events"
    Private Sub FRMMKTTRN0084_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Call FitToClient(Me, GrpMain, ctlFormHeader, GrpCmdBtn, 600)
            Me.MdiParent = mdifrmMain
            GrpCmdBtn.Left = GrpCmdBtn.Left + 50
            dtpToDt.MaxDate = GetServerDate()
            InitializeForm(1)
            Me.opt_all_SO.Checked = True
            OptDrNote.Checked = True
            Me.BringToFront()
            Call FN_Spread_Settings()
            cmbsuppinvtype.SelectedIndex = 0
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
                'dtpFrm.Value = GetServerDate().AddMonths(-1)
                dtpFrm.Value = GetServerDate()
                dtpToDt.Value = GetServerDate()
                txtCustCode.Text = String.Empty
                lblCustDesc.Text = String.Empty
                txtCustCode.Enabled = False
                BtnHelpProvDocNo.Enabled = True
                BtnFetchItem.Enabled = False
                optMultipleInvoice.Enabled = False
                optsingleInvoice.Enabled = False
                AddColumnDispatchDtlGrid()
                AddRateWiseGridColumn()

                ClearAllTaxes()
                dtDocTable = Nothing
                bindDocList()
                txtDocNo.Text = ""
                txtDocNo.Enabled = True
                dtpFrm.Enabled = False
                dtpToDt.Enabled = False
                txtTCSTaxCode.Text = ""
                lblTCSTaxPerDes.Text = ""
                BtnHelpTCS.Enabled = False
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
                opt_all_SO.Checked = True
                clearFarGrid()
                OptDrNote.Checked = True
            ElseIf Form_Status_flag = 2 Then    'Add Mode
                txtDocNo.Text = ""
                txtDocNo.Enabled = False
                'dtpFrm.Enabled = False
                'dtpToDt.Enabled = False
                optMultipleInvoice.Enabled = True
                optsingleInvoice.Enabled = True
                optsingleInvoice.Checked = True
                txtCustCode.Enabled = True
                BtnHelpCustCode.Enabled = True
                dtpFrm.Enabled = True
                dtpToDt.Enabled = True
                BtnFetchItem.Enabled = True
                txtCustCode.Text = ""
                lblCustDesc.Text = String.Empty
                txtDocNo.Text = ""
                txtCustCode.Enabled = False

                BtnHelpProvDocNo.Enabled = False
                txtDocNo.Enabled = False
                txtTCSTaxCode.Text = ""
                lblTCSTaxPerDes.Text = ""
                BtnHelpTCS.Enabled = False
                dtDocTable = Nothing
                bindDocList()
                clearFarGrid()
                ClearAllTaxes()
                dtpFrm.Focus()
                opt_all_SO.Checked = True
                clearFarGrid()
                OptDrNote.Checked = True
                ''InitializeForm(2)
            ElseIf Form_Status_flag = 3 Then    'Sales Prov View Mode for not submiited Doc No

                txtDocNo.Enabled = True
                dtpFrm.Enabled = False
                dtpToDt.Enabled = False
                txtCustCode.Enabled = False

                'txtTotBasValEff.Text = String.Empty
                BtnHelpProvDocNo.Enabled = True
                optsingleInvoice.Enabled = False
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
                BtnHelpProvDocNo.Enabled = False
                BtnFetchItem.Enabled = True
                txtDocNo.Enabled = False
                optMultipleInvoice.Enabled = False
                optsingleInvoice.Enabled = False
                LockUnlockRateWiseGrid(1, fpSpreadRateWiseDtl.MaxRows, False)
            ElseIf Form_Status_flag = 5 Then    'Sales Prov View Mode for submiited Doc No


                txtDocNo.Enabled = True
                dtpFrm.Enabled = False
                dtpToDt.Enabled = False
                txtCustCode.Enabled = False
                BtnHelpCustCode.Enabled = False
                BtnHelpProvDocNo.Enabled = True

                optMultipleInvoice.Enabled = False
                optsingleInvoice.Enabled = False
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


    '''<summary>To clear BOM exploded items from FAR grid</summary>
    Private Function clearFarGrid() As Boolean
        Dim flag As Boolean = True
        Try
            fpSpreadRateWiseDtl.MaxRows = 0
            SSinvoiceDetail.MaxRows = 0
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
        Dim strsqlHdr As String
        Dim strSqlDtl As String
        Dim inti As Integer
        Dim strrefno As Object
        Dim stramendmentno As Object
        Dim strinternalCode As Object
        Dim strInvoiceRate As Object
        Dim strNewInvoiceRate As Object
        Dim strCustPartCode As Object
        Dim strtaxableamount As Double
        Dim strinvoicequantity As Object
        Dim strCGSTTAXTYPE As Object
        Dim strCGSTTAXPERCENT As Object
        Dim strCGSTTAXAMT As Object
        Dim strSGSTTAXTYPE As Object
        Dim strSGSTTAXPERCENT As Object
        Dim strSGSTTAXAMT As Object
        Dim strIGSTTAXTYPE As Object
        Dim strIGSTTAXPERCENT As Object
        Dim strIGSTTAXAMT As Object
        Dim strbaseinvoiceno As Object
        Dim strprevbaseinvoiceno As Object
        Dim dbltotalamount As Double
        Dim strInvoicenewRate As Object
        Dim strFinalexecutablesql As String
        Dim strupdatesql As String
        Dim strsupp_documentnumber As Integer
        Dim STRDRCR As String
        Dim strsuppinvtype As String
        Dim dblTCSTaxAmount As Double
        Dim strTCSTaxPerDes As Object
        Dim strtotalamount As Double
        Try
            If ValidateSave() Then
                strsupp_documentnumber = Find_Value("select max(SERIES_SUPPINVOICE )+1 from saleconf WHERE UNIT_CODE='" + gstrUNITID + "'AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE AND INVOICE_TYPE='INV' AND SUB_TYPE='F'")
                If optsingleInvoice.Checked = True Then
                    With SSinvoiceDetail
                        For inti = 1 To .MaxRows

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.Rate
                            strInvoiceRate = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.NewRate
                            strInvoicenewRate = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.SalesOrder
                            strrefno = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.Amendmentno
                            stramendmentno = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.InternalPartNo
                            strinternalCode = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.CustPartNo
                            strCustPartCode = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.TaxableAmount
                            strtaxableamount = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.InvoiceQty
                            strinvoicequantity = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.CGST_TAX_TYPE
                            strCGSTTAXTYPE = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.CGST_TAX
                            strCGSTTAXPERCENT = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.SGST_TAX_TYPE
                            strSGSTTAXTYPE = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.SGST_TAX
                            strSGSTTAXPERCENT = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.IGST_TAX_TYPE
                            strIGSTTAXTYPE = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.IGST_TAX
                            strIGSTTAXPERCENT = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.BASE_INVOICENO
                            strbaseinvoiceno = .Value

                            dbltotalamount = CDbl(strtaxableamount)
                            If OptDrNote.Checked = True Then
                                STRDRCR = "DR"
                            Else
                                STRDRCR = "CR"
                            End If
                            dbltotalamount = System.Math.Abs(dbltotalamount)
                            strtaxableamount = System.Math.Abs(strtaxableamount)
                            If strCGSTTAXPERCENT.ToString.Length > 0 Then
                                dbltotalamount = dbltotalamount + System.Math.Round((strtaxableamount * strCGSTTAXPERCENT) / 100, 2)
                            End If
                            If strSGSTTAXPERCENT.ToString.Length > 0 Then
                                dbltotalamount = dbltotalamount + System.Math.Round((strtaxableamount * strSGSTTAXPERCENT) / 100, 2)
                            End If
                            If strIGSTTAXPERCENT.ToString.Length > 0 Then
                                dbltotalamount = dbltotalamount + System.Math.Round((strtaxableamount * strIGSTTAXPERCENT) / 100, 2)
                            End If

                            Call SelectChallanNoFromSupplementatryInvHdr()
                            If cmbsuppinvtype.Text = "Price Variance" Then
                                strsuppinvtype = "Price Variance"
                            Else
                                strsuppinvtype = "Retro"
                            End If
                            If Val(lblTCSTaxPerDes.Text) > 0 Then
                                dblTCSTaxAmount = CalculateTCSTax(dbltotalamount, blnTCSTax, Val(lblTCSTaxPerDes.Text))
                                dbltotalamount = dbltotalamount + dblTCSTaxAmount
                            Else
                                lblTCSTaxPerDes.Text = 0
                            End If

                            strsqlHdr = "Insert into SupplementaryInv_hdr ("
                            strsqlHdr = strsqlHdr & "DRCR,SuppInvType,Unit_code,Rate,Location_Code,Account_Code,Cust_name,Cust_Ref,Amendment_No,Doc_No,Invoice_DateFrom,Invoice_DateTo,Invoice_Date,Bill_Flag,Cancel_flag,Item_Code,"
                            strsqlHdr = strsqlHdr & "Cust_Item_Code,Currency_Code,Basic_Amount,Accessible_amount,total_amount,SERIES_SUPPINVOICE,"
                            strsqlHdr = strsqlHdr & "Ent_dt,Ent_UserId,Upd_dt,Upd_Userid,TCSTax_Type,TCSTax_Per,TCSTaxAmount"
                            strsqlHdr = strsqlHdr & ") Values"
                            strsqlHdr = strsqlHdr & " ('" & STRDRCR & "','" & strsuppinvtype & "','" & gstrUNITID & "'," & strInvoiceRate & ",'" & gstrUNITID & "','" & Trim(txtCustCode.Text) & "','" & Trim(lblCustDesc.Text) & "','"
                            strsqlHdr = strsqlHdr & strrefno & "','" & stramendmentno & "'," & pstrChallanNo & ",'" & getDateForDB(dtpFrm.Value) & "','" & getDateForDB(dtpToDt.Value) & "','" & getDateForDB(GetServerDate()) & "',0,0,'" & strinternalCode
                            strsqlHdr = strsqlHdr & "','" & strCustPartCode & "','" & lblCurrencyDes.Text & "',"
                            strsqlHdr = strsqlHdr & strtaxableamount & "," & strtaxableamount & "," & dbltotalamount & ",'"
                            strsqlHdr = strsqlHdr & strsupp_documentnumber & "','"
                            strsqlHdr = strsqlHdr & getDateForDB(GetServerDate()) & "','"
                            strsqlHdr = strsqlHdr & mP_User & "','" & getDateForDB(GetServerDate()) & "','" & mP_User & "'"
                            strsqlHdr = strsqlHdr & ",'" & txtTCSTaxCode.Text & "'," & lblTCSTaxPerDes.Text & "," & dblTCSTaxAmount
                            strsqlHdr = strsqlHdr & ")"
                            If strCGSTTAXTYPE = Nothing Or strCGSTTAXTYPE = "" Then
                                strCGSTTAXTYPE = "CGST0"
                                strCGSTTAXPERCENT = "0"
                            End If
                            If strSGSTTAXTYPE = Nothing Or strSGSTTAXTYPE = "" Then
                                strSGSTTAXTYPE = "SGST0"
                                strSGSTTAXPERCENT = "0"
                            End If
                            If strIGSTTAXTYPE = Nothing Or strIGSTTAXTYPE = "" Then
                                strIGSTTAXTYPE = "IGST0"
                                strIGSTTAXPERCENT = "0"
                            End If
                            
                            strSqlDtl = "insert into SupplementaryInv_Dtl ("
                            strSqlDtl = strSqlDtl & "Location_Code,Doc_No,SuppInvDate,Item_code,Cust_Item_Code,"
                            strSqlDtl = strSqlDtl & "Rate_diff,Quantity,Basic_AmountDiff,Accessible_amountDiff,"
                            strSqlDtl = strSqlDtl & "total_amountDiff,Ent_dt,Ent_UserID,Upd_Dt,Upd_UserID,"
                            strSqlDtl = strSqlDtl & "Refdoc_No,Unit_Code,CGSTTXRT_TYPE,CGST_PERCENT,SGSTTXRT_TYPE,SGST_PERCENT,IGSTTXRT_TYPE,IGST_PERCENT,DIFF_CGST_AMT,DIFF_SGST_AMT,DIFF_IGST_AMT)"
                            strSqlDtl = strSqlDtl & "Values ('" & gstrUNITID & "'," & pstrChallanNo & ",'" & getDateForDB(GetServerDate()) & "','" & strinternalCode & "','"
                            strSqlDtl = strSqlDtl & strCustPartCode & "'," & System.Math.Abs(strInvoicenewRate - strInvoiceRate) & "," & strinvoicequantity & ","
                            strSqlDtl = strSqlDtl & strtaxableamount & "," & strtaxableamount & "," & dbltotalamount & ",'"
                            strSqlDtl = strSqlDtl & getDateForDB(GetServerDate()) & "','" & mP_User & "','" & getDateForDB(GetServerDate()) & "','" & mP_User & "'," & strbaseinvoiceno & ",'" & gstrUNITID & "','"
                            strSqlDtl = strSqlDtl & strCGSTTAXTYPE & "'," & strCGSTTAXPERCENT & ",'" & strSGSTTAXTYPE & "',"
                            strSqlDtl = strSqlDtl & strSGSTTAXPERCENT & ",'" & strIGSTTAXTYPE & "'," & strIGSTTAXPERCENT & ","

                            If strCGSTTAXPERCENT.ToString.Length > 0 Then
                                strSqlDtl = strSqlDtl & System.Math.Round((strtaxableamount * strCGSTTAXPERCENT) / 100, 2) & ","
                            End If

                            If strSGSTTAXPERCENT.ToString.Length > 0 Then
                                strSqlDtl = strSqlDtl & System.Math.Round((strtaxableamount * strSGSTTAXPERCENT) / 100, 2) & ","
                            End If

                            If strIGSTTAXPERCENT.ToString.Length > 0 Then
                                strSqlDtl = strSqlDtl & System.Math.Round((strtaxableamount * strIGSTTAXPERCENT) / 100, 2)
                            Else
                                strSqlDtl = strSqlDtl & "0"
                            End If

                            strSqlDtl = strSqlDtl & ")"
                            'strupdatesql = "UPDATE H SET TOTAL_AMOUNT=ISNULL(TCSTAXAMOUNT,0)+ XYZ.TOTALAMOUNT  FROM (SELECT H.DOC_NO,H.UNIT_CODE, SUM(D.TOTAL_AMOUNTDIFF )  "
                            'strupdatesql += " AS TOTALAMOUNT FROM  SUPPLEMENTARYINV_HDR H , SUPPLEMENTARYINV_DTL D  WHERE "
                            'strupdatesql += " H.UNIT_CODE = D.UNIT_CODE AND H.DOC_NO =D.DOC_NO AND H.UNIT_CODE ='" & gstrUNITID & "'"
                            'strupdatesql += " GROUP BY  H.DOC_NO,H.UNIT_CODE)XYZ ,SUPPLEMENTARYINV_HDR H"
                            'strupdatesql += " WHERE XYZ.UNIT_CODE = H.UNIT_CODE AND XYZ.Doc_No =H.DOC_NO AND H.DOC_NO ='" & pstrChallanNo & "'"

                            ResetDatabaseConnection()
                            mP_Connection.BeginTrans()
                            If strsqlHdr.ToString.Length > 0 Then
                                mP_Connection.Execute(strsqlHdr, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If

                            If strSqlDtl.ToString.Length > 0 Then
                                mP_Connection.Execute(strSqlDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                            'If strupdatesql.ToString.Length > 0 Then
                            'mP_Connection.Execute(strupdatesql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            'End If

                            mP_Connection.Execute("UPDATE SALECONF SET SERIES_SUPPINVOICE=" & strsupp_documentnumber & " WHERE UNIT_CODE='" + gstrUNITID + "'AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE AND INVOICE_TYPE='INV' AND SUB_TYPE='F'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                            mP_Connection.CommitTrans()
                        Next
                    End With
                Else
                    Dim sarrSelInvoices() As String
                    Dim intMaxRow As Integer
                    Dim blnnewInvoiceno As Boolean = False
                    With SSinvoiceDetail

                        sarrSelInvoices = Split(strSelInvoices, ",")
                        intMaxRow = UBound(sarrSelInvoices) + 1

                        For inti = 1 To .MaxRows

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.Rate
                            strInvoiceRate = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.NewRate
                            strInvoicenewRate = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.SalesOrder
                            strrefno = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.Amendmentno
                            stramendmentno = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.InternalPartNo
                            strinternalCode = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.CustPartNo
                            strCustPartCode = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.TaxableAmount
                            strtaxableamount = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.InvoiceQty
                            strinvoicequantity = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.CGST_TAX_TYPE
                            strCGSTTAXTYPE = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.CGST_TAX
                            strCGSTTAXPERCENT = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.SGST_TAX_TYPE
                            strSGSTTAXTYPE = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.SGST_TAX
                            strSGSTTAXPERCENT = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.IGST_TAX_TYPE
                            strIGSTTAXTYPE = .Value

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.IGST_TAX
                            strIGSTTAXPERCENT = .Value

                            If inti = 1 Then
                                strprevbaseinvoiceno = "0"
                            Else
                                strprevbaseinvoiceno = strbaseinvoiceno
                            End If

                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.BASE_INVOICENO
                            strbaseinvoiceno = .Value

                            dbltotalamount = CDbl(strtaxableamount)
                            If OptDrNote.Checked = True Then
                                STRDRCR = "DR"
                            Else
                                STRDRCR = "CR"
                            End If
                            dbltotalamount = System.Math.Abs(dbltotalamount)
                            strtaxableamount = System.Math.Abs(strtaxableamount)

                            If strCGSTTAXPERCENT.ToString.Length > 0 Then
                                dbltotalamount = dbltotalamount + System.Math.Round((strtaxableamount * strCGSTTAXPERCENT) / 100, 2)
                            End If

                            If strSGSTTAXPERCENT.ToString.Length > 0 Then
                                dbltotalamount = dbltotalamount + System.Math.Round((strtaxableamount * strSGSTTAXPERCENT) / 100, 2)
                            End If

                            If strIGSTTAXPERCENT.ToString.Length > 0 Then
                                dbltotalamount = dbltotalamount + System.Math.Round((strtaxableamount * strIGSTTAXPERCENT) / 100, 2)
                            End If

                            If Val(lblTCSTaxPerDes.Text) > 0 Then
                                dblTCSTaxAmount = CalculateTCSTax(dbltotalamount, blnTCSTax, Val(lblTCSTaxPerDes.Text))
                                dbltotalamount = dbltotalamount + dblTCSTaxAmount
                            Else
                                lblTCSTaxPerDes.Text = 0
                            End If


                            If strbaseinvoiceno = strprevbaseinvoiceno Then
                                strsqlHdr = ""
                                strSqlDtl = "insert into SupplementaryInv_Dtl ("
                                strSqlDtl = strSqlDtl & "Location_Code,Doc_No,SuppInvDate,Item_code,Cust_Item_Code,"
                                strSqlDtl = strSqlDtl & "Rate_diff,Quantity,Basic_AmountDiff,Accessible_amountDiff,"
                                strSqlDtl = strSqlDtl & "total_amountDiff,Ent_dt,Ent_UserID,Upd_Dt,Upd_UserID,"
                                strSqlDtl = strSqlDtl & "Refdoc_No,Unit_Code,CGSTTXRT_TYPE,CGST_PERCENT,SGSTTXRT_TYPE,SGST_PERCENT,IGSTTXRT_TYPE,IGST_PERCENT,DIFF_CGST_AMT,DIFF_SGST_AMT,DIFF_IGST_AMT,TCSTax_Type,TCSTax_Per,TCSTaxAmount)"
                                strSqlDtl = strSqlDtl & "Values ('" & gstrUNITID & "'," & pstrChallanNo & ",'" & getDateForDB(GetServerDate()) & "','" & strinternalCode & "','"
                                strSqlDtl = strSqlDtl & strCustPartCode & "'," & System.Math.Abs(strInvoicenewRate - strInvoiceRate) & "," & strinvoicequantity & ","
                                strSqlDtl = strSqlDtl & strtaxableamount & "," & strtaxableamount & "," & dbltotalamount & ",'"
                                strSqlDtl = strSqlDtl & getDateForDB(GetServerDate()) & "','" & mP_User & "','" & getDateForDB(GetServerDate()) & "','" & mP_User & "'," & strbaseinvoiceno & ",'" & gstrUNITID & "','"
                                strSqlDtl = strSqlDtl & strCGSTTAXTYPE & "'," & strCGSTTAXPERCENT & ",'" & strSGSTTAXTYPE & "',"
                                strSqlDtl = strSqlDtl & strSGSTTAXPERCENT & ",'" & strIGSTTAXTYPE & "'," & strIGSTTAXPERCENT & ","
                                If strCGSTTAXPERCENT.ToString.Length > 0 Then
                                    strSqlDtl = strSqlDtl & System.Math.Round((strtaxableamount * strCGSTTAXPERCENT) / 100, 2) & ","
                                End If
                                If strSGSTTAXPERCENT.ToString.Length > 0 Then
                                    strSqlDtl = strSqlDtl & System.Math.Round((strtaxableamount * strSGSTTAXPERCENT) / 100, 2) & ","
                                End If
                                If strIGSTTAXPERCENT.ToString.Length > 0 Then
                                    strSqlDtl = strSqlDtl & System.Math.Round((strtaxableamount * strIGSTTAXPERCENT) / 100, 2)
                                Else
                                    strSqlDtl = strSqlDtl & "0"
                                End If
                                strSqlDtl = strSqlDtl & ",'" & txtTCSTaxCode.Text & "'," & lblTCSTaxPerDes.Text & "," & dblTCSTaxAmount
                                strSqlDtl = strSqlDtl & ")"

                            Else
                                Call SelectChallanNoFromSupplementatryInvHdr()

                                strSqlDtl = "insert into SupplementaryInv_Dtl ("
                                strSqlDtl = strSqlDtl & "Location_Code,Doc_No,SuppInvDate,Item_code,Cust_Item_Code,"
                                strSqlDtl = strSqlDtl & "Rate_diff,Quantity,Basic_AmountDiff,Accessible_amountDiff,"
                                strSqlDtl = strSqlDtl & "total_amountDiff,Ent_dt,Ent_UserID,Upd_Dt,Upd_UserID,"
                                strSqlDtl = strSqlDtl & "Refdoc_No,Unit_Code,CGSTTXRT_TYPE,CGST_PERCENT,SGSTTXRT_TYPE,SGST_PERCENT,IGSTTXRT_TYPE,IGST_PERCENT,DIFF_CGST_AMT,DIFF_SGST_AMT,DIFF_IGST_AMT,TCSTax_Type,TCSTax_Per,TCSTaxAmount)"
                                strSqlDtl = strSqlDtl & "Values ('" & gstrUNITID & "'," & pstrChallanNo & ",'" & getDateForDB(GetServerDate()) & "','" & strinternalCode & "','"
                                strSqlDtl = strSqlDtl & strCustPartCode & "'," & System.Math.Abs(strInvoicenewRate - strInvoiceRate) & "," & strinvoicequantity & ","
                                strSqlDtl = strSqlDtl & strtaxableamount & "," & strtaxableamount & "," & dbltotalamount & ",'"
                                strSqlDtl = strSqlDtl & getDateForDB(GetServerDate()) & "','" & mP_User & "','" & getDateForDB(GetServerDate()) & "','" & mP_User & "'," & strbaseinvoiceno & ",'" & gstrUNITID & "','"
                                strSqlDtl = strSqlDtl & strCGSTTAXTYPE & "'," & strCGSTTAXPERCENT & ",'" & strSGSTTAXTYPE & "',"
                                strSqlDtl = strSqlDtl & strSGSTTAXPERCENT & ",'" & strIGSTTAXTYPE & "'," & strIGSTTAXPERCENT & ","
                                If strCGSTTAXPERCENT.ToString.Length > 0 Then
                                    strSqlDtl = strSqlDtl & System.Math.Round((strtaxableamount * strCGSTTAXPERCENT) / 100, 2) & ","
                                End If
                                If strSGSTTAXPERCENT.ToString.Length > 0 Then
                                    strSqlDtl = strSqlDtl & System.Math.Round((strtaxableamount * strSGSTTAXPERCENT) / 100, 2) & ","
                                End If
                                If strIGSTTAXPERCENT.ToString.Length > 0 Then
                                    strSqlDtl = strSqlDtl & System.Math.Round((strtaxableamount * strIGSTTAXPERCENT) / 100, 2)
                                Else
                                    strSqlDtl = strSqlDtl & "0"
                                End If
                                strSqlDtl = strSqlDtl & ",'" & txtTCSTaxCode.Text & "'," & lblTCSTaxPerDes.Text & "," & dblTCSTaxAmount

                                strSqlDtl = strSqlDtl & ")"

                                If cmbsuppinvtype.Text = "Price Variance" Then
                                    strsuppinvtype = "Price Variance"
                                Else
                                    strsuppinvtype = "Retro"
                                End If
                                'If Val(lblTCSTaxPerDes.Text) > 0 Then
                                '    dblTCSTaxAmount = CalculateTCSTax(dbltotalamount, blnTCSTax, Val(lblTCSTaxPerDes.Text))
                                '    dbltotalamount = dbltotalamount + dblTCSTaxAmount
                                'Else
                                '    lblTCSTaxPerDes.Text = 0
                                'End If

                                strsqlHdr = "Insert into SupplementaryInv_hdr ("
                                strsqlHdr = strsqlHdr & "DRCR,SuppInvType,Unit_code,Rate,Location_Code,Account_Code,Cust_name,Cust_Ref,Amendment_No,Doc_No,Invoice_DateFrom,Invoice_DateTo,Invoice_Date,Bill_Flag,Cancel_flag,Item_Code,"
                                strsqlHdr = strsqlHdr & "Cust_Item_Code,Currency_Code,Basic_Amount,Accessible_amount,total_amount,SERIES_SUPPINVOICE,"
                                strsqlHdr = strsqlHdr & "Ent_dt,Ent_UserId,Upd_dt,Upd_Userid,TCSTax_Type,TCSTax_Per,TCSTaxAmount"
                                strsqlHdr = strsqlHdr & ") Values"
                                strsqlHdr = strsqlHdr & " ('" & STRDRCR & "','" & strsuppinvtype & "','" & gstrUNITID & "'," & strInvoiceRate & ",'" & gstrUNITID & "','" & Trim(txtCustCode.Text) & "','" & Trim(lblCustDesc.Text) & "','"
                                strsqlHdr = strsqlHdr & strrefno & "','" & stramendmentno & "'," & pstrChallanNo & ",'" & getDateForDB(dtpFrm.Value) & "','" & getDateForDB(dtpToDt.Value) & "','" & getDateForDB(GetServerDate()) & "',0,0,'" & strinternalCode
                                strsqlHdr = strsqlHdr & "','" & strCustPartCode & "','" & lblCurrencyDes.Text & "',"
                                strsqlHdr = strsqlHdr & strtaxableamount & "," & strtaxableamount & "," & dbltotalamount & ",'"
                                strsqlHdr = strsqlHdr & strsupp_documentnumber & "','"
                                strsqlHdr = strsqlHdr & getDateForDB(GetServerDate()) & "','"
                                strsqlHdr = strsqlHdr & mP_User & "','" & getDateForDB(GetServerDate()) & "','" & mP_User & "'"
                                strsqlHdr = strsqlHdr & ",'" & txtTCSTaxCode.Text & "'," & lblTCSTaxPerDes.Text & "," & dblTCSTaxAmount
                                strsqlHdr = strsqlHdr & ")"

                            End If
                            strupdatesql = "UPDATE H SET TOTAL_AMOUNT=XYZ.TOTALAMOUNT ,TCSTAXAMOUNT=XYZ.TCSTAXAMOUNT FROM (SELECT H.DOC_NO,H.UNIT_CODE, SUM(D.TOTAL_AMOUNTDIFF )AS TOTALAMOUNT  ,SUM(D.TCSTAXAMOUNT ) as TCSTAXAMOUNT   "
                            strupdatesql += " FROM  SUPPLEMENTARYINV_HDR H , SUPPLEMENTARYINV_DTL D  WHERE "
                            strupdatesql += " H.UNIT_CODE = D.UNIT_CODE AND H.DOC_NO =D.DOC_NO AND H.UNIT_CODE ='" & gstrUNITID & "'"
                            strupdatesql += " GROUP BY  H.DOC_NO,H.UNIT_CODE)XYZ ,SUPPLEMENTARYINV_HDR H"
                            strupdatesql += " WHERE XYZ.UNIT_CODE = H.UNIT_CODE AND XYZ.Doc_No =H.DOC_NO AND H.DOC_NO ='" & pstrChallanNo & "'"

                            mP_Connection.BeginTrans()
                            If strsqlHdr.ToString.Length > 0 Then
                                mP_Connection.Execute(strsqlHdr, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                            If strSqlDtl.ToString.Length > 0 Then
                                mP_Connection.Execute(strSqlDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                            If strupdatesql.ToString.Length > 0 Then
                                mP_Connection.Execute(strupdatesql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                            mP_Connection.Execute("UPDATE SALECONF SET SERIES_SUPPINVOICE=" & strsupp_documentnumber & " WHERE UNIT_CODE='" + gstrUNITID + "'AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE AND INVOICE_TYPE='INV' AND SUB_TYPE='F'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            mP_Connection.CommitTrans()
                        Next
                    End With

                End If
                Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                cmdGrpSalesProv.Revert()
                cmdGrpSalesProv.Top = 10
                cmdGrpSalesProv.Left = 10
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = True
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = False
                cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = False
                optMultipleInvoice.Checked = False
                optsingleInvoice.Checked = True
                ClearTmpTable()
                clearFarGrid()
                ClearAllTaxes()
                txtCustCode.Text = ""
                lblCustDesc.Text = String.Empty
                InitializeForm(1)
                txtCustCode.Enabled = False
                BtnHelpCustCode.Enabled = False
            End If


        Catch ex As Exception
            mP_Connection.RollbackTrans()
            RaiseException(ex)
        Finally
            SqlCmd.Dispose()
        End Try
        Return True
    End Function
    Private Function SaveData_Multiple() As Boolean
        Dim SqlCmd As New SqlCommand
        Dim strQuery As String = String.Empty
        Dim fs As FileStream
        Dim rawData As Object
        Dim odt As DataTable
        Dim strProvDocNo As String = String.Empty
        Dim strsqlHdr As String
        Dim strSqlDtl As String
        Dim inti As Integer
        Dim strrefno As Object
        Dim stramendmentno As Object
        Dim strinternalCode As Object
        Dim strInvoiceRate As Object
        Dim strNewInvoiceRate As Object
        Dim strCustPartCode As Object
        Dim strtaxableamount As Double
        Dim strinvoicequantity As Object
        Dim strCGSTTAXTYPE As Object
        Dim strCGSTTAXPERCENT As Object
        Dim strCGSTTAXAMT As Object
        Dim strSGSTTAXTYPE As Object
        Dim strSGSTTAXPERCENT As Object
        Dim strSGSTTAXAMT As Object
        Dim strIGSTTAXTYPE As Object
        Dim strIGSTTAXPERCENT As Object
        Dim strIGSTTAXAMT As Object
        Dim strbaseinvoiceno As Object
        Dim strprevbaseinvoiceno As Object
        Dim dbltotalamount As Double
        Dim strInvoicenewRate As Object
        Dim strFinalexecutablesql As String
        Dim strupdatesql As String
        Dim strsupp_documentnumber As Integer
        Dim STRDRCR As String
        Dim strsuppinvtype As String
        Dim dblTCSTaxAmount As Double
        Dim strTCSTaxPerDes As Object
        Dim strtotalamount As Double
        Dim STRHSNCODE As Object
        Dim strSlqQuery As String
        Dim strSql As String
        Dim strupdateSlqQuery As String
        Dim blnMORETHAN_ONEITEM_SUPPINV As Boolean = False
        Dim pstrcurrentChallanNo As String
        Try
            If ValidateSave() Then
                strsupp_documentnumber = Find_Value("select max(SERIES_SUPPINVOICE )+1 from saleconf WHERE UNIT_CODE='" + gstrUNITID + "'AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE AND INVOICE_TYPE='INV' AND SUB_TYPE='F'")
                Call SelectChallanNoFromSupplementatryInvHdr()

                If cmbsuppinvtype.Text = "Price Variance" Then
                    strsuppinvtype = "Price Variance"
                Else
                    strsuppinvtype = "Retro"
                End If

                If OptDrNote.Checked = True Then
                    STRDRCR = "DR"
                Else
                    STRDRCR = "CR"
                End If

                
                strSql = "delete from TMP_supplementaryData where Unit_code='" & gstrUNITID & "' and Ipaddress='" & gstrIpaddressWinSck & "'"
                SqlConnectionclass.ExecuteNonQuery(strSql)

                
                Dim sarrSelInvoices() As String
                Dim intMaxRow As Integer
                Dim blnnewInvoiceno As Boolean = False
                With SSinvoiceDetail

                    sarrSelInvoices = Split(strSelInvoices, ",")
                    intMaxRow = UBound(sarrSelInvoices) + 1

                    For inti = 1 To .MaxRows

                        .Row = inti
                        .Col = ENUM_Gridinvoicedetails.Rate
                        strInvoiceRate = .Value

                        .Row = inti
                        .Col = ENUM_Gridinvoicedetails.NewRate
                        strInvoicenewRate = .Value

                        .Row = inti
                        .Col = ENUM_Gridinvoicedetails.SalesOrder
                        strrefno = .Value

                        .Row = inti
                        .Col = ENUM_Gridinvoicedetails.Amendmentno
                        stramendmentno = .Value

                        .Row = inti
                        .Col = ENUM_Gridinvoicedetails.InternalPartNo
                        strinternalCode = .Value

                        .Row = inti
                        .Col = ENUM_Gridinvoicedetails.CustPartNo
                        strCustPartCode = .Value

                        .Row = inti
                        .Col = ENUM_Gridinvoicedetails.TaxableAmount
                        strtaxableamount = .Value

                        .Row = inti
                        .Col = ENUM_Gridinvoicedetails.InvoiceQty
                        strinvoicequantity = .Value

                        .Row = inti
                        .Col = ENUM_Gridinvoicedetails.CGST_TAX_TYPE
                        strCGSTTAXTYPE = .Value

                        .Row = inti
                        .Col = ENUM_Gridinvoicedetails.CGST_TAX
                        strCGSTTAXPERCENT = .Value

                        .Row = inti
                        .Col = ENUM_Gridinvoicedetails.SGST_TAX_TYPE
                        strSGSTTAXTYPE = .Value

                        .Row = inti
                        .Col = ENUM_Gridinvoicedetails.SGST_TAX
                        strSGSTTAXPERCENT = .Value

                        .Row = inti
                        .Col = ENUM_Gridinvoicedetails.IGST_TAX_TYPE
                        strIGSTTAXTYPE = .Value

                        .Row = inti
                        .Col = ENUM_Gridinvoicedetails.IGST_TAX
                        strIGSTTAXPERCENT = .Value

                        .Row = inti
                        .Col = ENUM_Gridinvoicedetails.HSNSACCODE
                        STRHSNCODE = .Value

                        If inti = 1 Then
                            strprevbaseinvoiceno = "0"
                        Else
                            strprevbaseinvoiceno = strbaseinvoiceno
                        End If

                        .Row = inti
                        .Col = ENUM_Gridinvoicedetails.BASE_INVOICENO
                        strbaseinvoiceno = .Value

                        dbltotalamount = CDbl(strtaxableamount)
                        If OptDrNote.Checked = True Then
                            STRDRCR = "DR"
                        Else
                            STRDRCR = "CR"
                        End If
                        dbltotalamount = System.Math.Abs(dbltotalamount)
                        strtaxableamount = System.Math.Abs(strtaxableamount)

                        If strCGSTTAXPERCENT.ToString.Length > 0 Then
                            dbltotalamount = dbltotalamount + System.Math.Round((strtaxableamount * strCGSTTAXPERCENT) / 100, 2)
                        End If

                        If strSGSTTAXPERCENT.ToString.Length > 0 Then
                            dbltotalamount = dbltotalamount + System.Math.Round((strtaxableamount * strSGSTTAXPERCENT) / 100, 2)
                        End If

                        If strIGSTTAXPERCENT.ToString.Length > 0 Then
                            dbltotalamount = dbltotalamount + System.Math.Round((strtaxableamount * strIGSTTAXPERCENT) / 100, 2)
                        End If

                        If Val(lblTCSTaxPerDes.Text) > 0 Then
                            dblTCSTaxAmount = CalculateTCSTax(dbltotalamount, blnTCSTax, Val(lblTCSTaxPerDes.Text))
                            dbltotalamount = dbltotalamount + dblTCSTaxAmount
                        Else
                            lblTCSTaxPerDes.Text = 0
                        End If


                        strSqlDtl = "insert into TMP_supplementaryData ("
                        strSqlDtl = strSqlDtl & "strPartNo,stracp,strOldRate,strNewRate,Unit_code,"
                        strSqlDtl = strSqlDtl & "ipaddress,item_code,account_code,DIFF_RATE,invoiceno,"
                        strSqlDtl = strSqlDtl & "RefInvoice_no,basic_amount,CGSTTAX_TYPE,CGSTTAX_PER,SGSTTAX_TYPE,"
                        strSqlDtl = strSqlDtl & "SGSTTAX_PER,IGSTTAX_TYPE,IGSTTAX_PER,CGST_AMT,SGST_AMT,iGST_AMT,TOTAL_AMOUNT,HSNSACCODE,TCSTAXAMOUNT,TCSTAXTYPEPER)"
                        strSqlDtl = strSqlDtl & "Values ('" & strCustPartCode & "'," & strinvoicequantity & "," & strInvoiceRate & "," & strInvoicenewRate & ",'" & gstrUNITID & "'"
                        strSqlDtl = strSqlDtl & ",'" & gstrIpaddressWinSck & "','" & strinternalCode & "','" & Trim(txtCustCode.Text) & "'," & System.Math.Abs(strInvoicenewRate - strInvoiceRate) & ",0,'" & strbaseinvoiceno & "'"
                        strSqlDtl = strSqlDtl & "," & strtaxableamount & ",'" & strCGSTTAXTYPE & "'," & strCGSTTAXPERCENT & ",'" & strSGSTTAXTYPE & "'," & strSGSTTAXPERCENT & ",'" & strIGSTTAXTYPE & "'"
                        strSqlDtl = strSqlDtl & "," & strIGSTTAXPERCENT & ","
                        If strCGSTTAXPERCENT.ToString.Length > 0 Then
                            strSqlDtl = strSqlDtl & System.Math.Round((strtaxableamount * strCGSTTAXPERCENT) / 100, 2) & ","
                        End If
                        If strSGSTTAXPERCENT.ToString.Length > 0 Then
                            strSqlDtl = strSqlDtl & System.Math.Round((strtaxableamount * strSGSTTAXPERCENT) / 100, 2) & ","
                        End If
                        If strIGSTTAXPERCENT.ToString.Length > 0 Then
                            strSqlDtl = strSqlDtl & System.Math.Round((strtaxableamount * strIGSTTAXPERCENT) / 100, 2)
                        Else
                            strSqlDtl = strSqlDtl & "0"
                        End If
                        strSqlDtl = strSqlDtl & "," & dbltotalamount & ",'" & STRHSNCODE & "'"
                        strSqlDtl = strSqlDtl & "," & dblTCSTaxAmount & ",'" & lblTCSTaxPerDes.Text & "'"
                        strSqlDtl = strSqlDtl & ")"

                        If strSqlDtl.ToString.Length > 0 Then
                            mP_Connection.Execute(strSqlDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If
                    Next


                    strSql = "select TOP 1 1 FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustCode.Text & "' AND MORETHAN_ONEITEM_SUPPINV=1 "

                    If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
                        blnMORETHAN_ONEITEM_SUPPINV = True
                    Else
                        blnMORETHAN_ONEITEM_SUPPINV = False
                    End If

                    SqlConnectionclass.BeginTrans()
                    If blnMORETHAN_ONEITEM_SUPPINV = True Then
                        strSlqQuery = " INSERT INTO supplementaryData_batchwise (invoice_no,PartNo,item_code,hsnsaccode,actqty,OldRate,NewRate,DIFF_RATE,Unit_code,Account_Code,BASIC_AMOUNT,CGSTTAX_TYPE,CGSTTAX_PER,CGST_AMT,SGSTTAX_TYPE,SGSTTAX_PER,SGST_AMT,IGSTTAX_TYPE,IGSTTAX_PER,IGST_AMT,PART_DESC,INVOICE_DATE,TOTAL_AMOUNT,TCSTAXAMOUNT,CURRENT_INVOICE_STATUS) "
                        strSlqQuery += " SELECT  " + pstrChallanNo + "+ DENSE_RANK() OVER ( PARTITION BY ITEM_CODE ORDER BY STROLDRATE),strPartNo,ITEM_CODE,hsnsaccode,SUM(stracp),strOldRate,strNewRate,DIFF_RATE,Unit_code,account_code ,sum(BASIC_AMOUNT),CGSTTAX_TYPE,"
                        strSlqQuery += " CGSTTAX_PER,sum(CGST_AMT),SGSTTAX_TYPE,SGSTTAX_PER,sum(SGST_AMT),IGSTTAX_TYPE,IGSTTAX_PER,sum(IGST_AMT),"
                        strSlqQuery += " PART_DESC,'" & getDateForDB(GetServerDate()) & "',sum(total_amount ) ,SUM(TCSTAXAMOUNT), 1 from TMP_supplementaryData  where unit_code='" & gstrUNITID & "' and Ipaddress='" & gstrIpaddressWinSck & "'"
                        strSlqQuery += "GROUP BY strPartNo,ITEM_CODE,hsnsaccode,strOldRate,strNewRate,DIFF_RATE,Unit_code,account_code ,CGSTTAX_TYPE, "
                        strSlqQuery += " CGSTTAX_PER,SGSTTAX_TYPE,SGSTTAX_PER,IGSTTAX_TYPE,IGSTTAX_PER,PART_DESC"
                        SqlConnectionclass.ExecuteNonQuery(strSlqQuery)

                    Else

                        strSlqQuery = " INSERT INTO supplementaryData_batchwise (invoice_no,PartNo,item_code,hsnsaccode,actqty,OldRate,NewRate,DIFF_RATE,Unit_code,Account_Code,BASIC_AMOUNT,CGSTTAX_TYPE,CGSTTAX_PER,CGST_AMT,SGSTTAX_TYPE,SGSTTAX_PER,SGST_AMT,IGSTTAX_TYPE,IGSTTAX_PER,IGST_AMT,PART_DESC,INVOICE_DATE,TOTAL_AMOUNT,TCSTAXAMOUNT,CURRENT_INVOICE_STATUS) "
                        strSlqQuery += " SELECT  " + pstrChallanNo + "+ DENSE_RANK() OVER ( ORDER BY STROLDRATE),strPartNo,ITEM_CODE,hsnsaccode,SUM(stracp),strOldRate,strNewRate,DIFF_RATE,Unit_code,account_code ,sum(BASIC_AMOUNT),CGSTTAX_TYPE,"
                        strSlqQuery += " CGSTTAX_PER,sum(CGST_AMT),SGSTTAX_TYPE,SGSTTAX_PER,sum(SGST_AMT),IGSTTAX_TYPE,IGSTTAX_PER,sum(IGST_AMT),"
                        strSlqQuery += " PART_DESC,'" & getDateForDB(GetServerDate()) & "',sum(total_amount ),SUM(TCSTAXAMOUNT),1 from TMP_supplementaryData  where unit_code='" & gstrUNITID & "' and Ipaddress='" & gstrIpaddressWinSck & "'"
                        strSlqQuery += "GROUP BY strPartNo,ITEM_CODE,hsnsaccode,strOldRate,strNewRate,DIFF_RATE,Unit_code,account_code ,CGSTTAX_TYPE, "
                        strSlqQuery += " CGSTTAX_PER,SGSTTAX_TYPE,SGSTTAX_PER,IGSTTAX_TYPE,IGSTTAX_PER,PART_DESC "
                        SqlConnectionclass.ExecuteNonQuery(strSlqQuery)

                    End If
                    If blnMORETHAN_ONEITEM_SUPPINV = True And optsingleInvoice.Checked = True Then
                        pstrcurrentChallanNo = pstrChallanNo + 1

                        strupdateSlqQuery = "UPDATE M SET INVOICENO=MB.INVOICE_NO FROM TMP_supplementaryData M INNER JOIN supplementaryData_batchwise MB ON"
                        strupdateSlqQuery += " M.UNIT_CODE=MB.UNIT_CODE AND "
                        strupdateSlqQuery += " M.ITEM_CODE=MB.ITEM_CODE AND M.strPartNo=MB.PARTNO"
                        strupdateSlqQuery += " AND M.strOldRate=MB.OLDRATE AND M.strnewRate=MB.NEWRATE AND M.DIFF_RATE=MB.DIFF_RATE"
                        strupdateSlqQuery += " AND MB.INVOICE_NO LIKE '99%' "
                        strupdateSlqQuery += " where M.unit_code='" & gstrUNITID & "' and M.invoiceNo=0 and  m.Ipaddress='" & gstrIpaddressWinSck & "' "
                        strupdateSlqQuery += " and MB.INVOICE_NO=" & pstrcurrentChallanNo

                    Else
                        strupdateSlqQuery = "UPDATE M SET INVOICENO=MB.INVOICE_NO FROM TMP_supplementaryData M INNER JOIN supplementaryData_batchwise MB ON"
                        strupdateSlqQuery += " M.UNIT_CODE=MB.UNIT_CODE AND "
                        strupdateSlqQuery += " M.ITEM_CODE=MB.ITEM_CODE AND M.strPartNo=MB.PARTNO"
                        strupdateSlqQuery += " AND M.strOldRate=MB.OLDRATE AND M.strnewRate=MB.NEWRATE AND M.DIFF_RATE=MB.DIFF_RATE"
                        strupdateSlqQuery += " AND MB.INVOICE_NO LIKE '99%' "
                        strupdateSlqQuery += " where M.unit_code='" & gstrUNITID & "' and M.invoiceNo=0 and  m.Ipaddress='" & gstrIpaddressWinSck & "'"
                        strupdateSlqQuery += " and current_invoice_status=1"

                    End If
                    
                    SqlConnectionclass.ExecuteNonQuery(strupdateSlqQuery)



                    strSlqQuery = "INSERT INTO supplementaryData_detail (invoice_no,PartNo,item_code,hsnsaccode,actqty,OldRate,NewRate,DIFF_RATE,Unit_code,Ref_invNo,Account_Code,BASIC_AMOUNT,CGSTTAX_TYPE,CGSTTAX_PER,CGST_AMT,SGSTTAX_TYPE,SGSTTAX_PER,SGST_AMT,IGSTTAX_TYPE,IGSTTAX_PER,IGST_AMT,part_desc,[invoice_date],[TOTAL_AMOUNT],ent_dt,ent_userid ) "
                    strSlqQuery += " SELECT  invoiceno,strPartNo,ITEM_CODE,hsnsaccode,stracp,strOldRate,strNewRate,DIFF_RATE,Unit_code,refinvoice_no,account_code ,BASIC_AMOUNT,CGSTTAX_TYPE,"
                    strSlqQuery += " CGSTTAX_PER,CGST_AMT,SGSTTAX_TYPE,SGSTTAX_PER,SGST_AMT,IGSTTAX_TYPE,IGSTTAX_PER,IGST_AMT,"
                    strSlqQuery += " PART_DESC,'" & getDateForDB(GetServerDate()) & "',total_amount ," & " getdate(),'" & mP_User & "' from TMP_supplementaryData  where unit_code='" & gstrUNITID & "' and Ipaddress='" & gstrIpaddressWinSck & "'"
                    SqlConnectionclass.ExecuteNonQuery(strSlqQuery)

                    strsqlHdr = "Insert into SupplementaryInv_hdr ("
                    strsqlHdr = strsqlHdr & "DRCR,SuppInvType,Unit_code,Location_Code,Account_Code,Cust_name,Doc_No,cust_ref,"
                    strsqlHdr = strsqlHdr & "Invoice_DateFrom, Invoice_DateTo, Invoice_Date, Bill_Flag, Cancel_flag,"
                    strsqlHdr = strsqlHdr & "Currency_Code,"
                    strsqlHdr = strsqlHdr & "Ent_dt,Ent_UserId,Upd_dt,Upd_Userid,tcstax_type,tcstax_per)"
                    strsqlHdr = strsqlHdr & " select distinct '" & STRDRCR & "','" & strsuppinvtype & "','" & gstrUNITID & "','" & gstrUNITID & "',sb.Account_Code,'" & Trim(lblCustDesc.Text) & "',"
                    strsqlHdr = strsqlHdr & " invoice_no,invoice_no,invoice_date,invoice_date,invoice_date,0,0,'INR',"
                    strsqlHdr = strsqlHdr & " '" & getDateForDB(GetServerDate()) & "','"
                    strsqlHdr = strsqlHdr & mP_User & "','" & getDateForDB(GetServerDate()) & "','" & mP_User & "','" & txtTCSTaxCode.Text & "','" & lblTCSTaxPerDes.Text & "'"
                    strsqlHdr = strsqlHdr & " from supplementaryData_batchwise sb inner join TMP_supplementaryData TM on "
                    strsqlHdr = strsqlHdr & " sb.unit_code=tm.unit_code and sb.invoice_no=tm.invoiceno where tm.unit_code='" & gstrUNITID & "'"
                    strsqlHdr = strsqlHdr & " and ipaddress ='" & gstrIpaddressWinSck & "'"
                    If blnMORETHAN_ONEITEM_SUPPINV = True And optsingleInvoice.Checked = True Then
                        strsqlHdr = strsqlHdr & " and sb.invoice_no=" & pstrcurrentChallanNo
                    End If

                    SqlConnectionclass.ExecuteNonQuery(strsqlHdr)

                    strSqlDtl = "insert into supplementaryinv_dtl(Location_Code,Doc_No,LastSupplementary,SuppInvDate,Item_code,"
                    strSqlDtl = strSqlDtl & "Cust_Item_Code,PrevRate,Rate,Rate_diff,Quantity,"
                    strSqlDtl = strSqlDtl & " Basic_AmountDiff,Accessible_amountDiff,"
                    strSqlDtl = strSqlDtl & "total_amountDiff,Ent_dt,Ent_UserID,Upd_Dt,Upd_UserID,UNIT_CODE,"
                    strSqlDtl = strSqlDtl & "DIFF_CGST_AMT, DIFF_SGST_AMT,DIFF_IGST_AMT,CGSTTXRT_TYPE,CGST_PERCENT,SGSTTXRT_TYPE,"
                    strSqlDtl = strSqlDtl & "SGST_PERCENT,IGSTTXRT_TYPE,IGST_PERCENT)"
                    strSqlDtl = strSqlDtl & "select distinct '" & gstrUNITID & "',sb.invoice_no, 0,'" & getDateForDB(GetServerDate()) & "',SB.item_Code,"
                    strSqlDtl = strSqlDtl & "PartNo,SB.oldrate,SB.newrate,sb.diff_rate,actqty,sb.basic_amount,sb.basic_amount,sb.total_amount"
                    strSqlDtl = strSqlDtl & ",'" & getDateForDB(GetServerDate()) & " ','"
                    strSqlDtl = strSqlDtl & mP_User & "','" & getDateForDB(GetServerDate()) & "','" & mP_User & "','" & gstrUNITID & "',"
                    strSqlDtl = strSqlDtl & "SB.CGST_AMT ,SB.SGST_AMT,SB.IGST_AMT,SB.CGSTTAX_TYPE,SB.CGSTTAX_PER,SB.SGSTTAX_TYPE,SB.SGSTTAX_PER,SB.IGSTTAX_TYPE,SB.IGSTTAX_PER "
                    strSqlDtl = strSqlDtl & "from supplementaryData_batchwise SB inner join TMP_supplementaryData TM on "
                    strSqlDtl = strSqlDtl & " sb.unit_code=tm.unit_code and sb.invoice_no=tm.invoiceno where tm.unit_code='" & gstrUNITID & "'"
                    strSqlDtl = strSqlDtl & " and ipaddress ='" & gstrIpaddressWinSck & "'"
                    If blnMORETHAN_ONEITEM_SUPPINV = True And optsingleInvoice.Checked = True Then
                        strsqlHdr = strsqlHdr & " and sb.invoice_no=" & pstrcurrentChallanNo
                    End If

                    SqlConnectionclass.ExecuteNonQuery(strSqlDtl)

                    If blnMORETHAN_ONEITEM_SUPPINV = True Then
                        strupdateSlqQuery = " UPDATE SH SET SH.TOTAL_AMOUNT=XYZ.TOTAL_AMOUNT , BASIC_AMOUNT =XYZ.BASIC_AMOUNT, SH.TCSTAXAMOUNT= XYZ.TCSTAXAMOUNT FROM ("
                        strupdateSlqQuery += " SELECT SUM(BASIC_AMOUNT)BASIC_AMOUNT ,SUM(TOTAL_AMOUNT)TOTAL_AMOUNT ,SUM(TCSTAXAMOUNT)TCSTAXAMOUNT , INVOICE_NO ,UNIT_CODE FROM SUPPLEMENTARYDATA_BATCHWISE "
                        strupdateSlqQuery += " GROUP BY UNIT_CODE,INVOICE_NO )XYZ inner join SUPPLEMENTARYINV_HDR sh ON "
                        strupdateSlqQuery += "  SH.UNIT_CODE = XYZ.UNIT_CODE AND SH.DOC_NO=XYZ.INVOICE_NO"
                        strupdateSlqQuery += " INNER JOIN TMP_SUPPLEMENTARYDATA TP ON TP.UNIT_CODE=XYZ.UNIT_CODE "
                        strupdateSlqQuery += " AND SH.DOC_NO=TP.INVOICENO where TP.UNIT_CODE='" & gstrUNITID & "' AND TP.IPADDRESS ='" & gstrIpaddressWinSck & "'"
                        strupdateSlqQuery += " AND SH.BILL_FLAG=0 "
                        SqlConnectionclass.ExecuteNonQuery(strupdateSlqQuery)

                    Else
                        strupdateSlqQuery = "UPDATE SH SET SH.TOTAL_AMOUNT=M.TOTAL_AMOUNT , BASIC_AMOUNT =M.BASIC_AMOUNT , SH.TCSTAXAMOUNT= M.TCSTAXAMOUNT FROM "
                        strupdateSlqQuery += " SUPPLEMENTARYDATA_BATCHWISE M INNER JOIN SUPPLEMENTARYINV_HDR SH "
                        strupdateSlqQuery += " ON M.UNIT_CODE=SH.UNIT_CODE AND M.INVOICE_NO=SH.DOC_NO "
                        strupdateSlqQuery += " INNER JOIN TMP_SUPPLEMENTARYDATA TP ON TP.UNIT_CODE=M.UNIT_CODE "
                        strupdateSlqQuery += " AND SH.DOC_NO=TP.INVOICENO where TP.UNIT_CODE='" & gstrUNITID & "' AND TP.IPADDRESS ='" & gstrIpaddressWinSck & "'"
                        strupdateSlqQuery += " AND SH.BILL_FLAG=0 "
                        SqlConnectionclass.ExecuteNonQuery(strupdateSlqQuery)
                    End If
                    strupdateSlqQuery = "UPDATE MB SET CURRENT_INVOICE_STATUS=0 FROM TMP_supplementaryData M INNER JOIN supplementaryData_batchwise MB ON"
                    strupdateSlqQuery += " M.UNIT_CODE=MB.UNIT_CODE AND "
                    strupdateSlqQuery += " M.ITEM_CODE=MB.ITEM_CODE AND M.strPartNo=MB.PARTNO"
                    strupdateSlqQuery += " AND M.strOldRate=MB.OLDRATE AND M.strnewRate=MB.NEWRATE AND M.DIFF_RATE=MB.DIFF_RATE"
                    strupdateSlqQuery += " where M.unit_code='" & gstrUNITID & "' and  m.Ipaddress='" & gstrIpaddressWinSck & "'"
                    strupdateSlqQuery += " AND CURRENT_INVOICE_STATUS=1"

                    SqlConnectionclass.ExecuteNonQuery(strupdateSlqQuery)
                    SqlConnectionclass.CommitTran()

                End With

            End If

            Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            cmdGrpSalesProv.Revert()
            cmdGrpSalesProv.Top = 10
            cmdGrpSalesProv.Left = 10
            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = True
            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = False
            cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = False
            optMultipleInvoice.Checked = False
            optsingleInvoice.Checked = True
            ClearTmpTable()
            clearFarGrid()
            ClearAllTaxes()
            txtCustCode.Text = ""
            lblCustDesc.Text = String.Empty
            InitializeForm(1)
            txtCustCode.Enabled = False
            BtnHelpCustCode.Enabled = False



        Catch ex As Exception
            mP_Connection.RollbackTrans()
            RaiseException(ex)
        Finally
            SqlCmd.Dispose()
        End Try
        Return True
    End Function

    Public Function Find_Value(ByRef strField As String) As String

        On Error GoTo ErrHandler
        Dim rs As New ADODB.Recordset
        rs = New ADODB.Recordset
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.Open(strField, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
        If rs.RecordCount > 0 Then
            If IsDBNull(rs.Fields(0).Value) = False Then
                Find_Value = rs.Fields(0).Value
            Else
                Find_Value = ""
            End If
        Else
            Find_Value = ""
        End If
        rs.Close()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub SelectChallanNoFromSupplementatryInvHdr()
        Dim strChallanNo As String
        Dim rsChallanNo As ClsResultSetDB
        Dim blnMORETHAN_ONEITEM_SUPPINV As Boolean = False
        Dim strSql As String

        Try
            strSql = "select TOP 1 1 FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustCode.Text & "' AND MORETHAN_ONEITEM_SUPPINV=1 "

            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
                BLNMORETHAN_ONEITEM_SUPPINV = True
            Else
                BLNMORETHAN_ONEITEM_SUPPINV = False
            End If

            strChallanNo = "Select max(Doc_No) as Doc_No from SupplementaryInv_hdr where Unit_code='" & gstrUNITID & "' and bill_flag=0 and Doc_No>" & 99000000
            rsChallanNo = New ClsResultSetDB
            rsChallanNo.GetResult(strChallanNo, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsChallanNo.GetNoRows > 0 Then
                If Val(rsChallanNo.GetValue("Doc_No")) = 0 Then
                    pstrChallanNo = "99000001"
                Else
                    If optsingleInvoice.Checked = True Or (optMultipleInvoice.Checked = True And BLNMORETHAN_ONEITEM_SUPPINV = False) Then
                        pstrChallanNo = CStr(Val(rsChallanNo.GetValue("Doc_No")) + 1)
                    Else
                        pstrChallanNo = CStr(Val(rsChallanNo.GetValue("Doc_No")))

                    End If


                End If
            Else
                pstrChallanNo = "99000001"
            End If

            rsChallanNo.ResultSetClose()
            rsChallanNo = Nothing
            Exit Sub

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ''' <summary>Validate manadatory fields</summary>
    Private Function ValidateSave() As Boolean
        Dim strQry As String = String.Empty
        Dim isValid As Boolean = False
        Dim newrate As Double = 0.0
        Dim oldrate As Double = 0.0
        Dim inti As Integer
        Try
            If String.IsNullOrEmpty(txtCustCode.Text.Trim) Then
                MessageBox.Show("Customer cannot be blank.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                Return False
            End If

            If gblnGSTUnit = True And txtTCSTaxCode.Enabled = True Then
                If txtTCSTaxCode.Text.Trim.ToString.Length = 0 Then
                    MessageBox.Show("Entered TCS TAX .", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                    Return False
                End If
            End If

            With fpSpreadRateWiseDtl
                If .MaxRows = 0 Then
                    MessageBox.Show("Select atleast one Part Code.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                    Return False
                End If
                'vaildation
                With fpSpreadRateWiseDtl
                    For inti = 1 To .MaxRows
                        .Row = inti
                        .Col = RateWiseDtlGrid.Col_NewInvRate
                        If .Value <> "" Then
                            If .Value < 0 Then
                                .Value = "0"
                                MessageBox.Show("Rate cannot be in -ve.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                                Return False
                            End If
                        End If
                    Next
                End With

                'validation
                If OptDrNote.Checked = True Then
                    'DEBIT NOTE VALIDATION
                    With SSinvoiceDetail
                        For inti = 1 To .MaxRows
                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.TaxableAmount

                            If .Value.ToString.Length > 0 Then
                                If .Value < 0 Then
                                    MessageBox.Show("Kindly Enter Correct Value ,Value should be positive ", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Return False
                                End If
                            Else
                                MessageBox.Show("Kindly Enter some Value .", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                                Return False
                            End If
                        Next
                    End With
                Else 'CREDIT NOTE VALIDATION
                    With SSinvoiceDetail

                        For inti = 1 To .MaxRows
                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.TaxableAmount

                            If .Value.ToString.Length > 0 Then
                                If .Value > 0 Then
                                    MessageBox.Show("Kindly Enter Correct Value ! Current Rate should be less than Base Invoice Rate ", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    .Focus()
                                    Return False
                                End If
                            Else
                                MessageBox.Show("Kindly Enter some Value .", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                                Return False
                            End If
                        Next
                    End With
                End If

                'For row As Integer = 1 To .MaxRows
                '    .Row = row
                '    .Col = RateWiseDtlGrid.Col_Select
                '    If .Value = 1 Then                      'Is row selected
                '        isValid = True
                '        .Col = RateWiseDtlGrid.Col_TotEffVal
                '        If .Value = 0.0 Then
                '            MessageBox.Show("New Rate must be different from Old rate.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                '            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                '            .Focus()
                '            Return False
                '        End If



                '    End If
                'Next

            End With


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
                .Parameters("@p_ProvDocNo").Value = txtDocNo.Text.Trim
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

                .MaxCols = RateWiseDtlGrid.Col_Part_Code
                .Col = RateWiseDtlGrid.Col_Part_Code
                .Value = "Part Code"
                .set_ColWidth(RateWiseDtlGrid.Col_Part_Code, 15)

                .MaxCols = RateWiseDtlGrid.Col_NewInvRate
                .Col = RateWiseDtlGrid.Col_NewInvRate
                .Value = "New Invoice Rate"
                .set_ColWidth(RateWiseDtlGrid.Col_NewInvRate, 13)

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

                        '.Col = RateWiseDtlGrid.Col_Inv_Qty
                        '.Value = Convert.ToString(row("Invoice_Qty"))
                        '.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        '.TypeFloatMax = 9999999999.99
                        '.TypeFloatMin = 0
                        '.TypeFloatSeparator = False
                        '.TypeFloatDecimalPlaces = 2


                        '.Col = RateWiseDtlGrid.Col_InvRate
                        '.Value = Convert.ToString(row("Rate"))
                        '.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        '.TypeFloatMax = 9999999999.99
                        '.TypeFloatMin = 0
                        '.TypeFloatSeparator = False
                        '.TypeFloatDecimalPlaces = 2



                        ''If Convert.ToBoolean(row("SELECT")) Then
                        'CalculateRateEffect(.Row)
                        'End If



                        '.Col = RateWiseDtlGrid.Col_Select
                        '.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                        '.Value = 0

                        If Not cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                            LockUnlockRateWiseGrid(.Row, .Row, False)   'Unlock grid
                        Else
                            LockUnlockRateWiseGrid(.Row, .Row, True)  'lock grid
                        End If
                    Next
                    'GetTotalBasicValEffect()
                End With
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ''' <summary>Column addition in Dispatch Detail DataGridView</summary>
    Private Sub AddColumnDispatchDtlGrid()

    End Sub

    ''' <summary>Fill data in Tmp Table.</summary>
    Private Sub FillTmpTable()
        Dim strQry As String = String.Empty
        Try
            'For Each row In dt.Rows
            '            STRSQL = "INSERT INTO TMP_SUPPLEMENETARYINV_DATA(CUST_ITEM_CODE,ITEM_CODE,IP_ADDRESS,UNIT_CODE) "
            '           strSQL = strSQL + " VALUES('" & row("VariantCode").ToString() & "','" & row("ColorCode").ToString() & "','" & row("ColorDesc").ToString() & "'," & Double.Parse(row("SOBPer").ToString()) & "," & Double.Parse(row("SOBVol").ToString()) & ",'" & strIpAddress & "','" + gstrUNITID + "')"
            '          SqlConnectionclass.ExecuteNonQuery(strSQL)
            '         Next
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ''' <summary>Clear data in Tmp Table.</summary>
    Private Sub ClearTmpTable()
        Dim strQry As String = String.Empty
        Try
            strQry = "DELETE FROM TMP_SUPPLEMENETARYINV_DATA WHERE UNIT_CODE='" + gstrUNITID + "' AND IP_ADDRESS='" + gstrIpaddressWinSck + "'"

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
            'strQry = "SELECT CASE ISNULL(SRQ.ITEM_CODE,'') WHEN '' THEN 0 ELSE 1 END [SELECT],TMP.CUSTOMER_CODE, TMP.ITEM_CODE, TMP.VARMODEL" & _
            '        " , TMP.RATE, SUM(TMP.INVOICE_QTY) INVOICE_QTY, SUM(TMP.ACTUALINVOICEQTY) ACTUALINVOICEQTY, ISNULL(SRQ.PRICECHANGETYPE, TMP.PRICECHANGETYPE) PRICECHANGETYPE" & _
            '        " , ISNULL(SRQ.NEWRATE, TMP.NEWRATE) NEWRATE, ISNULL(SRQ.CHANGEEFFECT, TMP.CHANGEEFFECT) CHANGEEFFECT" & _
            '        " , ISNULL(SRQ.PERCENTAGECHANGE, TMP.PERCENTAGECHANGE) PERCENTAGECHANGE, ISNULL(SRQ.CHANGEREASON, TMP.CHANGEREASON) CHANGEREASON" & _
            '        " , ISNULL(SPNC.CORRECTIONDESC,'') CORRECTIONDESC, ISNULL(SPNC.CORRECTIONID,'') CORRECTIONID" & _
            '        " FROM SALES_PROV_TMPPARTDETAIL TMP LEFT JOIN SALES_PROV_RATEWISEDETAIL SRQ" & _
            '        " ON TMP.CUSTOMER_CODE=SRQ.CUSTOMER_CODE AND TMP.ITEM_CODE=SRQ.ITEM_CODE AND TMP.VARMODEL=SRQ.VARMODEL AND TMP.UNIT_CODE=SRQ.UNIT_CODE" & _
            '        " AND TMP.RATE=SRQ.RATE AND CAST(SRQ.PROV_DOCNO AS VARCHAR(20))='" + txtDocNo.Text.Trim() + "'" & _
            '        " LEFT JOIN SALES_PROV_NATUREOFCORRECTION SPNC ON SRQ.NATUREOFCORRECTION=SPNC.CORRECTIONID" & _
            '        " WHERE TMP.IPADDRESS='" + gstrIpaddressWinSck + "' AND TMP.UNIT_CODE='" + gstrUNITID + "'" & _
            '        " and tmp.Item_code in ('" + strQry.Trim() + "')" & _
            '        " GROUP BY TMP.CUSTOMER_CODE, TMP.ITEM_CODE, TMP.VARMODEL, CASE ISNULL(SRQ.ITEM_CODE,'') WHEN '' THEN 0 ELSE 1 END, TMP.RATE" & _
            '        " , ISNULL(SRQ.PRICECHANGETYPE, TMP.PRICECHANGETYPE), ISNULL(SRQ.NEWRATE, TMP.NEWRATE), ISNULL(SRQ.CHANGEEFFECT, TMP.CHANGEEFFECT)" & _
            '        " , ISNULL(SRQ.PERCENTAGECHANGE, TMP.PERCENTAGECHANGE), ISNULL(SRQ.CHANGEREASON, TMP.CHANGEREASON), ISNULL(SPNC.CORRECTIONDESC,'')" & _
            '        " , ISNULL(SPNC.CORRECTIONID,'')"
            strQry = "select item_code from item_mst where unit_code='" & gstrUNITID & "' and Item_code in ('" + strQry.Trim() + "')"
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

                'If Not status And Not cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then  'Unlock Grid
                '    For i As Integer = row1 To row2
                '        '.Col = RateWiseDtlGrid.Col_PriceChange
                '        .Col = RateWiseDtlGrid.Col_NewInvRate
                '        .Row = i
                '        val = .Value
                '        .Row2 = i

                '        If val = 0 Then
                '            .Col = RateWiseDtlGrid.Col_Change
                '            .Value = ""
                '            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                '            .Col = RateWiseDtlGrid.Col_NewInvRate
                '            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                '            .Value = ""
                '        Else
                '            .Col = RateWiseDtlGrid.Col_NewRate
                '            .Value = ""
                '            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                '        End If

                '        .BlockMode = True
                '        .Col = RateWiseDtlGrid.Col_NewRate
                '        .Col2 = RateWiseDtlGrid.Col_NewRate
                '        '.Lock = Not (val = 0)  ' Mayur
                '        .Lock = True  ' Mayur
                '        .Col = RateWiseDtlGrid.Col_ChangeEff
                '        .Col2 = RateWiseDtlGrid.Col_ChangeEff
                '        .Lock = (val = 0)
                '        .Col = RateWiseDtlGrid.Col_Change
                '        .Col2 = RateWiseDtlGrid.Col_Change
                '        .Lock = (val = 0)
                '        .BlockMode = False
                '    Next
                'End If
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub CalculateRateEffect(ByVal row As Integer)
        'Dim InvRate As Double = 0.0
        'Dim RateChange As Double = 0.0
        'Dim NewInvRate As Double = 0.0
        'Dim TotEffVal As Double = 0.0
        'Dim priceChangeType As Integer = 0
        'Dim changeEffect As Integer = 0
        'Dim ActInvQty As Double = 0.0
        'Try
        '    With fpSpreadRateWiseDtl
        '        .Row = row
        '        .Col = RateWiseDtlGrid.Col_InvRate
        '        InvRate = Convert.ToDecimal(.Value)
        '        .Col = RateWiseDtlGrid.Col_ChangeEff
        '        changeEffect = .Value
        '        .Col = RateWiseDtlGrid.Col_ActInvQty
        '        ActInvQty = Convert.ToDouble(.Value)
        '        If priceChangeType = 0 Then
        '            .Col = RateWiseDtlGrid.Col_NewRate
        '            If String.IsNullOrEmpty(.Value) Or .Value = "0.00" Then
        '                RateChange = 0.0
        '                .Value = ""
        '            Else
        '                RateChange = Math.Round(Convert.ToDouble(.Value), 2)
        '            End If
        '        Else
        '            .Col = RateWiseDtlGrid.Col_Change
        '            If String.IsNullOrEmpty(.Value) Or .Value = "0.00" Then
        '                RateChange = 0.0
        '                .Value = ""
        '            Else
        '                RateChange = Math.Round(Convert.ToDouble(.Value), 2)
        '            End If
        '            .Col = RateWiseDtlGrid.Col_NewInvRate
        '            If RateChange = 0.0 Then
        '                .Value = ""
        '            ElseIf changeEffect = 0 Then
        '                RateChange = Math.Round(InvRate * (1 - RateChange / 100), 2)
        '                .Value = RateChange
        '                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
        '                .TypeFloatMax = 9999999999.99
        '                .TypeFloatMin = 0
        '                .TypeFloatSeparator = False
        '                .TypeFloatDecimalPlaces = 2
        '            Else
        '                RateChange = Math.Round(InvRate * (1 + RateChange / 100), 2)
        '                .Value = RateChange
        '                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
        '                .TypeFloatMax = 9999999999.99
        '                .TypeFloatMin = 0
        '                .TypeFloatSeparator = False
        '                .TypeFloatDecimalPlaces = 2
        '            End If
        '        End If
        '        .Col = RateWiseDtlGrid.Col_TotEffVal
        '        TotEffVal = ActInvQty * (RateChange - InvRate)
        '        If RateChange = 0.0 Then
        '            .Value = ""
        '        Else
        '            .Value = TotEffVal
        '            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
        '            .TypeFloatMax = 9999999999.99
        '            .TypeFloatMin = 0
        '            .TypeFloatSeparator = False
        '            .TypeFloatDecimalPlaces = 2
        '        End If
        '    End With
        'Catch ex As Exception
        '    RaiseException(ex)
        'End Try
    End Sub

    Private Sub GetTotalBasicValEffect()
        Dim inti As Integer
        Dim strtaxablevalue As Object
        Dim intnewrate As Double
        Dim intrate As Double
        Dim intQuantity As Double

        With SSinvoiceDetail
            For inti = 1 To .MaxRows
                .Row = inti
                .Col = ENUM_Gridinvoicedetails.NewRate
                If .Value <> "" Then
                    intnewrate = .Value
                    .Col = ENUM_Gridinvoicedetails.Rate
                    intrate = .Value
                    .Col = ENUM_Gridinvoicedetails.InvoiceQty
                    intQuantity = .Value
                    .Col = ENUM_Gridinvoicedetails.TaxableAmount
                    strtaxablevalue = Nothing
                    strtaxablevalue = intQuantity * (intnewrate - intrate)
                    .SetText(ENUM_Gridinvoicedetails.TaxableAmount, inti, Convert.ToDecimal(strtaxablevalue).ToString("0.00"))
                    If strtaxablevalue < 0 Then
                        .BlockMode = True
                        .Row = inti
                        .Row2 = inti
                        .Col = ENUM_Gridinvoicedetails.TaxableAmount
                        .Col2 = ENUM_Gridinvoicedetails.TaxableAmount
                        .ForeColor = Color.Red
                        .FontBold = True
                        .Lock = True
                        .BlockMode = False
                    Else
                        .BlockMode = True
                        .Row = inti
                        .Row2 = inti
                        .Col = ENUM_Gridinvoicedetails.TaxableAmount
                        .Col2 = ENUM_Gridinvoicedetails.TaxableAmount
                        .ForeColor = Color.Green
                        .FontBold = True
                        .Lock = True
                        .BlockMode = False
                    End If
                End If

            Next
        End With

    End Sub

    Private Sub ClearAllTaxes()
        SSinvoiceDetail.MaxRows = 0
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
    Function FetchinvoiceRecords() As String
        Dim StrSql As String
        Dim StrgstSql As String
        Dim Row, Col As Integer
        Dim rsRecords As ClsResultSetDB
        Dim rsgsttaxes As ClsResultSetDB
        Dim varDespatchQty As Object = Nothing
        Dim rsPrevQty As ADODB.Recordset
        Dim objSQLConn As SqlConnection
        Dim objReader As SqlDataReader
        Dim objCommand As SqlCommand

        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        Call FN_Spread_Settings()
        Row = 1

        If cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            StrSql = "SELECT D.REFDOC_NO AS DOC_NO, D.DOC_NO as CURR_INVOICENO,D.ITEM_CODE,D.CUST_ITEM_CODE,CM.DRG_DESC AS CUST_ITEM_DESC ,CM.Item_Desc AS cust_item_desc,cust_ref,amendment_no,D.Quantity AS SALES_QUANTITY,D.RATE_DIFF AS RATE , d.CGSTTXRT_TYPE,d.CGST_PERCENT,d.SGSTTXRT_TYPE,d.SGST_PERCENT,d.IGSTTXRT_TYPE,d.IGST_PERCENT " & _
            " FROM SUPPLEMENTARYINV_HDR H (NOLOCK) INNER JOIN  SUPPLEMENTARYINV_DTL D(NOLOCK) " & _
           " ON H.UNIT_CODE=D.UNIT_CODE AND H.DOC_NO=D.DOC_NO " & _
            "INNER JOIN CUSTITEM_MST CM ON CM.UNIT_CODE=D.UNIT_CODE AND CM.ITEM_CODE=D.ITEM_CODE AND CM.Cust_Drgno =D.Cust_Item_Code AND CM.ACCOUNT_CODE=H.Account_Code " & _
            " AND H.UNIT_CODE=cm.UNIT_CODE AND H.DOC_NO=D.DOC_NO  " & _
           " AND H.UNIT_CODE = '" & gstrUNITID & "' AND H.SERIES_SUPPINVOICE IN(" & txtDocNo.Text & ")"
        Else
            StrSql = "select * from (" & _
                "SELECT d.DOC_NO, d.item_code,d.cust_item_code,d.cust_item_desc,d.cust_ref ,d.amendment_no,SALES_QUANTITY-(SELECT ISNULL(SUM(QUANTITY_REC),0) FROM GRN_INVOICE_DTL GI WHERE GI.UNIT_CODE=d.UNIT_CODE AND GI.INVOICE_NO =d.DOC_NO AND GI.ITEM_CODE =d.ITEM_CODE ) AS SALES_QUANTITY ,d.rate FROM SALESCHALLAN_DTL H (NOLOCK) INNER JOIN SALES_DTL D(NOLOCK) " & _
               "  ON H.UNIT_CODE=D.UNIT_CODE AND H.DOC_NO=D.DOC_NO AND H.UNIT_CODE = '" & gstrUNITID & "' AND H.DOC_NO IN(" & strSelInvoices & ")" & _
            "AND EXISTS (SELECT TOP 1 1 FROM TMP_SUPPLEMENETARYINV_DATA TS WHERE TS.UNIT_CODE=D.UNIT_CODE AND TS.ITEM_CODE=D.ITEM_CODE " & _
            " AND TS.CUST_ITEM_CODE=D.CUST_ITEM_CODE AND TS.IP_ADDRESS='" & gstrIpaddressWinSck & "'))a where a.SALES_QUANTITY >0 order by a.doc_no"
        End If
        rsRecords = New ClsResultSetDB
        rsRecords.GetResult(StrSql)

        If rsRecords.GetNoRows > 0 Then
            rsRecords.MoveFirst()
            While Not rsRecords.EOFRecord
                With SSinvoiceDetail
                    SSinvoiceDetail.MaxRows = SSinvoiceDetail.MaxRows + 1
                    .SetText(ENUM_Gridinvoicedetails.BASE_INVOICENO, .MaxRows, rsRecords.GetValue("DOC_NO"))

                    If cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                        .SetText(ENUM_Gridinvoicedetails.CURR_INVOICENO, .MaxRows, rsRecords.GetValue("CURR_INVOICENO"))
                        .ColHidden = False
                    End If

                    .SetText(ENUM_Gridinvoicedetails.InternalPartNo, .MaxRows, rsRecords.GetValue("item_code"))
                    .SetText(ENUM_Gridinvoicedetails.CustPartNo, .MaxRows, rsRecords.GetValue("cust_item_code"))
                    .SetText(ENUM_Gridinvoicedetails.ItemDesc, .MaxRows, rsRecords.GetValue("cust_item_desc"))
                    .SetText(ENUM_Gridinvoicedetails.CustPartDesc, .MaxRows, rsRecords.GetValue("cust_item_desc"))
                    .SetText(ENUM_Gridinvoicedetails.SalesOrder, .MaxRows, rsRecords.GetValue("cust_ref"))
                    .SetText(ENUM_Gridinvoicedetails.Amendmentno, .MaxRows, rsRecords.GetValue("amendment_no"))
                    .SetText(ENUM_Gridinvoicedetails.InvoiceQty, .MaxRows, rsRecords.GetValue("sales_quantity"))
                    .SetText(ENUM_Gridinvoicedetails.Rate, .MaxRows, rsRecords.GetValue("Rate"))
                    If cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                        .SetText(ENUM_Gridinvoicedetails.CGST_TAX_TYPE, .MaxRows, rsRecords.GetValue("CGSTTXRT_TYPE"))
                        .SetText(ENUM_Gridinvoicedetails.SGST_TAX_TYPE, .MaxRows, rsRecords.GetValue("sGSTTXRT_TYPE"))
                        .SetText(ENUM_Gridinvoicedetails.IGST_TAX_TYPE, .MaxRows, rsRecords.GetValue("IGSTTXRT_TYPE"))
                        .SetText(ENUM_Gridinvoicedetails.CGST_TAX, .MaxRows, rsRecords.GetValue("CGST_PERCENT"))
                        .SetText(ENUM_Gridinvoicedetails.SGST_TAX, .MaxRows, rsRecords.GetValue("SGST_PERCENT"))
                        .SetText(ENUM_Gridinvoicedetails.IGST_TAX, .MaxRows, rsRecords.GetValue("IGST_PERCENT"))
                    End If

                    StrgstSql = "set dateformat 'dmy' SELECT * FROM DBO.UFN_GST_ITEMWISETAXES('" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & rsRecords.GetValue("item_code") & "','" & GetServerDateNew() & "','" & GetServerDateNew() & "')"
                    objSQLConn = SqlConnectionclass.GetConnection()
                    objCommand = New SqlCommand(StrgstSql, objSQLConn)
                    objReader = objCommand.ExecuteReader()
                    If objReader.HasRows = True Then
                        objReader.Read()
                        .SetText(ENUM_Gridinvoicedetails.HSNSACCODE, .MaxRows, objReader.GetValue(1).ToString)
                        .SetText(ENUM_Gridinvoicedetails.CGST_TAX_TYPE, .MaxRows, objReader.GetValue(2).ToString)
                        .SetText(ENUM_Gridinvoicedetails.CGST_TAX, .MaxRows, GetTaxRate(objReader.GetValue(2).ToString, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='CGST'"))
                        .SetText(ENUM_Gridinvoicedetails.SGST_TAX_TYPE, .MaxRows, objReader.GetValue(3).ToString)
                        .SetText(ENUM_Gridinvoicedetails.SGST_TAX, .MaxRows, GetTaxRate(objReader.GetValue(3).ToString, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SGST'"))
                        .SetText(ENUM_Gridinvoicedetails.IGST_TAX_TYPE, .MaxRows, objReader.GetValue(5).ToString)
                        .SetText(ENUM_Gridinvoicedetails.IGST_TAX, .MaxRows, GetTaxRate(objReader.GetValue(5).ToString, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='IGST'"))
                    End If
                End With
                rsRecords.MoveNext()
            End While
        End If

        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Return Nothing
        Exit Function
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function GetTaxRate(ByRef pstrFieldText As String, ByRef pstrColumnName As String, ByRef pstrTableName As String, ByRef pstrFieldName_WhichValueRequire As String, Optional ByRef pstrCondition As String = "") As Double
        On Error GoTo ErrHandler
        GetTaxRate = 0
        Dim strTableSql As String 'Declared To Make Select Query
        Dim rsExistData As ClsResultSetDB
        If Len(Trim(pstrCondition)) > 0 Then
            strTableSql = "select " & Trim(pstrFieldName_WhichValueRequire) & " from " & Trim(pstrTableName) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and " & pstrCondition
        Else
            strTableSql = "select " & Trim(pstrFieldName_WhichValueRequire) & " from " & Trim(pstrTableName) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "'"
        End If
        rsExistData = New ClsResultSetDB
        rsExistData.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsExistData.GetNoRows > 0 Then

            GetTaxRate = rsExistData.GetValue(Trim(pstrFieldName_WhichValueRequire))
        Else
            GetTaxRate = 0
        End If
        rsExistData.ResultSetClose()
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
#End Region

#Region "Control Events"
    Private Sub cmdGrpSalesProv_ButtonClick(ByVal Sender As System.Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles cmdGrpSalesProv.ButtonClick
        Dim blnMORETHAN_ONEITEM_SUPPINV As Boolean = False
        Dim strsql As String 
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

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                    strSql = "select TOP 1 1 FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustCode.Text & "' AND MORETHAN_ONEITEM_SUPPINV=1 "

                    If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
                        BLNMORETHAN_ONEITEM_SUPPINV = True
                    Else
                        BLNMORETHAN_ONEITEM_SUPPINV = False
                    End If



                    If optsingleInvoice.Checked = True Or (optMultipleInvoice.Checked = True And BLNMORETHAN_ONEITEM_SUPPINV = False) Then
                        SaveData()
                    Else
                        SaveData_Multiple()

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
                    optMultipleInvoice.Checked = False
                    optsingleInvoice.Checked = True
                    dtpFrm.Enabled = True
                    dtpToDt.Enabled = True
                    ClearAllTaxes()
                    ClearTmpTable()
                    BtnFetchItem.Enabled = True

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                    InitializeForm(4)
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = True
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = False
                    cmdGrpSalesProv.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True

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
                            optMultipleInvoice.Checked = False
                            optsingleInvoice.Checked = False
                        End If
                    End If
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
                    'GenerateAnexure()
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
            strSQL = " select distinct series_suppinvoice , Invoice_DateFrom ,Invoice_DateTo,Account_Code,Cust_name,Currency_Code from supplementaryinv_hdr where Unit_Code='" + gstrUNITID + "' and series_suppinvoice >0  order by series_suppinvoice  desc"
            With ctlHelp
                .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                .ConnectAsUser = gstrCONNECTIONUSER
                .ConnectThroughDSN = gstrCONNECTIONDSN
                .ConnectWithPWD = gstrCONNECTIONPASSWORD
            End With
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select SUPP", 1)
            If Not IsNothing(strHelp) Then
                If strHelp.Length > 0 Then
                    txtDocNo.Text = strHelp(0).Trim
                    odt = SqlConnectionclass.GetDataTable(strSQL)
                    If odt.Rows.Count > 0 Then
                        dtpFrm.Value = Convert.ToDateTime(odt.Rows(0)("Invoice_DateFrom"))
                        dtpToDt.Value = Convert.ToDateTime(odt.Rows(0)("Invoice_DateTo"))
                        txtCustCode.Text = Convert.ToString(odt.Rows(0)("Account_Code"))
                        lblCustDesc.Text = Convert.ToString(odt.Rows(0)("CUST_NAME"))
                        lblCurrencyDes.Text = Convert.ToString(odt.Rows(0)("Currency_Code"))
                        'If strHelp(5).Trim.ToString = "NORMAL" Then
                        optsingleInvoice.Checked = False
                        optMultipleInvoice.Checked = False
                        'bindRateItemfpGrid()
                        Call FetchinvoiceRecords()
                        'End If
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
    Private Sub BtnHelpHdr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnHelpCustCode.Click
        Try
            Dim strSQL As String = String.Empty
            Dim blntcscheck As Boolean = False
            Dim Strdrcr As String
            Dim strHelp As String()
            With ctlHelp
                .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                .ConnectAsUser = gstrCONNECTIONUSER
                .ConnectThroughDSN = gstrCONNECTIONDSN
                .ConnectWithPWD = gstrCONNECTIONPASSWORD
            End With

            'If optsingleInvoice.Checked = True Then
            strSQL = "Select DISTINCT Account_code,Cust_Name from VW_SUPP_INVOICE_CUSTHELP Where " & _
             " UNIT_CODE='" & gstrUNITID & "'" & _
            "AND INVOICE_DATE BETWEEN '" + dtpFrm.Value.ToString("dd MMM yyyy") + "' AND '" + dtpToDt.Value.ToString("dd MMM yyyy") + "'"

            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "Select Customer Code")

            If Not IsNothing(strHelp) Then

                txtCustCode.Text = strHelp(0)
                lblCustDesc.Text = strHelp(1)
                clearFarGrid()
                ClearAllTaxes()
                If txtCustCode.Text <> "0" Then
                    FillCurrency()
                End If
            End If
            'End If

            If Len(Trim(txtCustCode.Text)) > 0 Then
                If OptDrNote.Checked = True Then
                    strdrcr = "DR"
                Else
                    strdrcr = "CR"
                End If
                strSQL = "select dbo.UDF_IRN_TCSREQUIRED_SUPP( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & strdrcr & "')"
                If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
                    txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtTCSTaxCode.Text = "" : BtnHelpTCS.Enabled = True
                Else
                    txtTCSTaxCode.Enabled = False : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtTCSTaxCode.Text = "" : BtnHelpTCS.Enabled = False
                End If

            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub FillCurrency()
        Try
            Dim strSQL As String = String.Empty

            lblCurrencyDes.Text = SqlConnectionclass.ExecuteScalar("Select isnull(Currency_code,'') from customer_mst Where UNIT_CODE='" & gstrUNITID & "' and customer_code= '" & txtCustCode.Text & "'")
            If lblCurrencyDes.Text = "" Then
                lblCurrencyDes.Text = gstrCURRENCYCODE
           
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
            If Not String.IsNullOrEmpty(txtDocNo.Text) Then
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
    Private Sub txtdocno_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDocNo.KeyPress
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub fpSpreadRateWiseDtl_LeaveCell(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles fpSpreadRateWiseDtl.LeaveCell
        Dim strpartcode As String
        Dim inti As Integer
        Dim strnewinvoicerate As String
        Try
            If e.row > 0 Then
                With fpSpreadRateWiseDtl

                    .Row = e.row
                    .Col = RateWiseDtlGrid.Col_Part_Code
                    strpartcode = .Value

                    .Row = e.row
                    .Col = RateWiseDtlGrid.Col_NewInvRate
                    strnewinvoicerate = .Value

                    With SSinvoiceDetail
                        For inti = 1 To .MaxRows
                            .Row = inti
                            .Col = ENUM_Gridinvoicedetails.InternalPartNo
                            If .Value = strpartcode Then
                                .SetText(ENUM_Gridinvoicedetails.NewRate, inti, strnewinvoicerate)
                            End If

                        Next
                    End With
                    GetTotalBasicValEffect()
                    '            If .Value = 1 Then
                    '                If e.col = RateWiseDtlGrid.Col_Change Or e.col = RateWiseDtlGrid.Col_NewRate Then
                    '                    CalculateRateEffect(e.row)
                    '                    GetTotalBasicValEffect()
                    '                    CalculateTaxes()
                    '                End If
                    '            End If
                    '            'fpSpreadRateWiseDtl.SetActiveCell(e.col, e.row)
                    '            'fpSpreadRateWiseDtl_Enter(fpSpreadRateWiseDtl, New System.EventArgs)
                End With
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub fpSpreadRateWiseDtl_KeyDownEvent(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles fpSpreadRateWiseDtl.KeyDownEvent
        Dim strHelp As String()
        Dim strQuery As String = String.Empty
        'Try
        '    If e.keyCode = Keys.F1 Then

        '        ' Code Added By Mayur Against issue ID 10816097 
        '        If fpSpreadRateWiseDtl.ActiveCol = RateWiseDtlGrid.Col_NewRate And fpSpreadRateWiseDtl.ActiveRow > 0 And (cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD) And optMultipleInvoice.Checked = False Then
        '            fpSpreadRateWiseDtl.Row = fpSpreadRateWiseDtl.ActiveRow
        '            fpSpreadRateWiseDtl.Col = RateWiseDtlGrid.Col_Select
        '            If fpSpreadRateWiseDtl.Value = 1 Then
        '                fpSpreadRateWiseDtl.Col = RateWiseDtlGrid.Col_Part_Code
        '                fpSpreadRateWiseDtl.Row = fpSpreadRateWiseDtl.ActiveRow
        '                part_code = fpSpreadRateWiseDtl.Value.Trim.ToString()
        '                fpSpreadRateWiseDtl.Col = RateWiseDtlGrid.Col_InvRate
        '                fpSpreadRateWiseDtl.Row = fpSpreadRateWiseDtl.ActiveRow
        '                rate = fpSpreadRateWiseDtl.Value
        '                If (part_code.Trim().ToString() <> "") Then
        '                    GetPriceChangeDetails(part_code.Trim().ToString(), rate, "ADD")
        '                    fpSpreadRateWiseDtl.Col = RateWiseDtlGrid.Col_NewRate
        '                    fpSpreadRateWiseDtl.Row = fpSpreadRateWiseDtl.ActiveRow
        '                    fpSpreadRateWiseDtl.Lock = True
        '                End If
        '            End If
        '        End If

        '        If txtDocNo.Text <> "" Then
        '            If fpSpreadRateWiseDtl.ActiveCol = RateWiseDtlGrid.Col_NewRate And fpSpreadRateWiseDtl.ActiveRow > 0 And (cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Or cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT) Then

        '                fpSpreadRateWiseDtl.Col = RateWiseDtlGrid.Col_InvRate
        '                fpSpreadRateWiseDtl.Row = fpSpreadRateWiseDtl.ActiveRow
        '                rate = fpSpreadRateWiseDtl.Value
        '                fpSpreadRateWiseDtl.Col = RateWiseDtlGrid.Col_Part_Code
        '                fpSpreadRateWiseDtl.Row = fpSpreadRateWiseDtl.ActiveRow
        '                part_code = fpSpreadRateWiseDtl.Value.Trim.ToString()

        '                If (cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT) Then
        '                    If (part_code.Trim().ToString() <> "") Then
        '                        GetPriceChangeDetails(part_code.Trim().ToString(), rate, "EDIT")
        '                    End If
        '                End If
        '                If (cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW) Then
        '                    If (part_code.Trim().ToString() <> "") Then
        '                        GetPriceChangeDetails(part_code.Trim().ToString(), rate, "VIEW")
        '                    End If
        '                End If

        '            End If

        '        End If
        '        ' Code Added By Mayur Against issue ID 10816097 
        '    End If
        'Catch ex As Exception
        '    RaiseException(ex)
        'End Try
    End Sub
    Private Sub fpSpreadRateWiseDtl_ClickEvent(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles fpSpreadRateWiseDtl.ClickEvent
        Try
            With fpSpreadRateWiseDtl
                If e.row > 0 Then
                    '.Col = RateWiseDtlGrid.Col_Model_Code
                    '.Row = e.row
                    'txtRModelDesc.Text = Convert.ToString(SqlConnectionclass.ExecuteScalar("SELECT MODEL_DESC FROM BUDGET_MODEL_MST(NOLOCK) WHERE UNIT_CODE='" + gstrUNITID + "' AND ACTIVE=1 AND MODEL_CODE='" + Convert.ToString(.Value).Trim() + "'"))
                    '.Col = RateWiseDtlGrid.Col_Part_Code
                    .Row = e.row
                    'txtRPartDesc.Text = Convert.ToString(SqlConnectionclass.ExecuteScalar("SELECT [DESCRIPTION] FROM ITEM_MST(NOLOCK) WHERE UNIT_CODE='" + gstrUNITID + "' AND ITEM_CODE='" + Convert.ToString(.Value).Trim() + "'"))
                End If
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtHelp_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustCode.KeyDown, txtDocNo.KeyDown
        Try
            If e.KeyCode = Keys.F1 Then
                If (sender Is txtCustCode) Then
                    BtnHelpHdr_Click(BtnHelpCustCode, New EventArgs())
                ElseIf (sender Is txtDocNo) Then
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
            fdt = SqlConnectionclass.ExecuteReader("select DocData from Sales_Prov_DocList where cast(Prov_DocNo as varchar(20))='" + txtDocNo.Text.Trim() + "' and UNIT_CODE='" + gstrUNITID + "' and DocName='" + imavar1 + "'")
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

    End Sub

    Private Sub BtnUploadDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try

            DocFrm = New Form()
            DocFrm.StartPosition = FormStartPosition.CenterParent
            DocFrm.Width = 665
            DocFrm.Height = 260
            DocFrm.BackColor = Me.BackColor
            DocFrm.Text = "Upload Document"
            DocFrm.MaximizeBox = False
            DocFrm.MinimizeBox = False
            DocFrm.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
            DocFrm.ShowDialog()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

#End Region

    ' Code Added By Mayur Against issue ID 10816097 
    Private Sub getdata(ByRef item_code As String)
        Try
            If optMultipleInvoice.Checked = True And cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then

                Dim strQry As String = String.Empty
                Dim sqlCmd As New SqlCommand()
                Dim odata As New SqlDataAdapter
                Dim odatatable As New DataSet

                clearFarGrid()
                If odatatable.Tables.Count > 0 Then
                    If Not IsNothing(odatatable.Tables(1)) Then
                        With fpSpreadRateWiseDtl
                            .MaxRows = 0
                            For Each row As DataRow In odatatable.Tables(1).Rows
                                .MaxRows = .MaxRows + 1
                                .Row = .MaxRows
                                .set_RowHeight(.Row, 12)

                                .Col = RateWiseDtlGrid.Col_Part_Code
                                .Value = Convert.ToString(row("ITEM_CODE"))

                                '.Col = RateWiseDtlGrid.Col_Model_Code
                                '.Value = Convert.ToString(row("VarModel"))


                                .Col = RateWiseDtlGrid.Col_NewInvRate
                                .Value = Convert.ToString(row("Rate"))
                                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                .TypeFloatMax = 9999999999.99
                                .TypeFloatMin = 0
                                .TypeFloatSeparator = False
                                .TypeFloatDecimalPlaces = 2


                                '.Col = RateWiseDtlGrid.Col_ChangeEff
                                '.CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
                                '.TypeComboBoxList = "-" + Chr(9) + "+"
                                'If Convert.ToBoolean(row("ChangeEffect")) Then   '0 for Subtraction and 1 for Addition;
                                '    .TypeComboBoxCurSel = 1
                                'Else
                                '    .TypeComboBoxCurSel = 0
                                'End If


                                '.Col = RateWiseDtlGrid.Col_NewRate
                                'If Not Convert.ToBoolean(row("PriceChangeType")) Then     '0 for Value and 1 for %
                                '    .Value = Convert.ToString(row("NewRate"))
                                'End If
                                '.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                '.TypeFloatMax = 9999999999.99
                                '.TypeFloatMin = 0
                                '.TypeFloatSeparator = False
                                '.TypeFloatDecimalPlaces = 2

                                '.Col = RateWiseDtlGrid.Col_Change
                                '.Value = Convert.ToString(row("PercentageChange"))
                                '.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                '.TypeFloatMax = 999.99
                                '.TypeFloatMin = 0
                                '.TypeFloatSeparator = False
                                '.TypeFloatDecimalPlaces = 2

                                CalculateRateEffect(.Row)
                                'End If


                                LockUnlockRateWiseGrid(.Row, .Row, True)  'lock grid

                            Next
                            GetTotalBasicValEffect()
                        End With
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub rbtn_NormalInvoice_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optsingleInvoice.CheckedChanged
        Try
            If optsingleInvoice.Checked = True Then
                txtCustCode.Enabled = True
                BtnHelpCustCode.Enabled = True
                dtpFrm.Enabled = True
                dtpToDt.Enabled = True
                BtnFetchItem.Enabled = True

            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub optMultipleInvoice_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optMultipleInvoice.CheckedChanged
        Try
            If optMultipleInvoice.Checked = True Then
                txtCustCode.Enabled = True
                BtnHelpCustCode.Enabled = True
                dtpFrm.Enabled = True
                dtpToDt.Enabled = True
                BtnFetchItem.Enabled = True
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub btn_fileName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            If cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If optMultipleInvoice.Checked = True Then
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

                            txtCustCode.Text = Convert.ToString(SqlConnectionclass.ExecuteScalar("SELECT DISTINCT TOP(1) ACCOUNT_CODE  FROM TMP_SALESFILEDATA WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE IS NOT NULL AND BATCH_NO IN (SELECT BATCH_NO FROM MARUTI_FILE_HDR WHERE UNITCODE ='" & gstrUNITID & "' AND FILENAME ='" & strhelp(0).Trim() & "')"))
                            strQry = ""
                            strQry = "SELECT DISTINCT CM.CUSTOMER_CODE, CM.CUST_NAME [Customer_Name], CM.KAMCODE [KAM_Code], EM.Name [KAM_Name] FROM CUSTOMER_MST CM left join Employee_mst EM on CM.UNIT_CODE=EM.UNIT_CODE and CM.KAMCODE=EM.Employee_code WHERE CM.UNIT_CODE='" & gstrUNITID & "' AND CM.CUSTOMER_CODE ='" & txtCustCode.Text.Trim.ToString() & "'"
                            odata = SqlConnectionclass.GetDataTable(strQry)
                            If odata.Rows.Count > 0 Then
                                txtCustCode.Text = odata.Rows(0)(0).ToString()
                                lblCustDesc.Text = odata.Rows(0)(1).ToString()
                                BtnFetchItem.Enabled = True
                                optsingleInvoice.Enabled = False
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

                If Not IsNothing(odatatable.Tables(1)) Then
                    With fpSpreadRateWiseDtl
                        .MaxRows = 0
                        For Each row As DataRow In odatatable.Tables(1).Rows
                            .MaxRows = .MaxRows + 1
                            .Row = .MaxRows
                            .set_RowHeight(.Row, 12)

                            .Col = RateWiseDtlGrid.Col_Part_Code
                            .Value = Convert.ToString(row("ITEM_CODE"))

                            '.Col = RateWiseDtlGrid.Col_Model_Code
                            '.Value = Convert.ToString(row("VarModel"))

                            .Col = RateWiseDtlGrid.Col_NewInvRate
                            .Value = Convert.ToString(row("Invoice_Qty"))
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                            .TypeFloatMax = 9999999999.99
                            .TypeFloatMin = 0
                            .TypeFloatSeparator = False
                            .TypeFloatDecimalPlaces = 2

                            '.Col = RateWiseDtlGrid.Col_ActInvQty
                            '.Value = Convert.ToString(row("ActualInvoiceQty"))
                            '.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                            '.TypeFloatMax = 9999999999.99
                            '.TypeFloatMin = 0
                            '.TypeFloatSeparator = False
                            '.TypeFloatDecimalPlaces = 2

                            .Col = RateWiseDtlGrid.Col_NewInvRate
                            .Value = Convert.ToString(row("Rate"))
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                            .TypeFloatMax = 9999999999.99
                            .TypeFloatMin = 0
                            .TypeFloatSeparator = False
                            .TypeFloatDecimalPlaces = 2


                            '.Col = RateWiseDtlGrid.Col_NewRate
                            'If Not Convert.ToBoolean(row("PriceChangeType")) Then     '0 for Value and 1 for %
                            '    .Value = Convert.ToString(row("NewRate"))
                            'End If
                            '.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                            '.TypeFloatMax = 9999999999.99
                            '.TypeFloatMin = 0
                            '.TypeFloatSeparator = False
                            '.TypeFloatDecimalPlaces = 2

                            '.Col = RateWiseDtlGrid.Col_Change
                            '.Value = Convert.ToString(row("PercentageChange"))
                            '.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                            '.TypeFloatMax = 999.99
                            '.TypeFloatMin = 0
                            '.TypeFloatSeparator = False
                            '.TypeFloatDecimalPlaces = 2

                            'CalculateRateEffect(.Row)
                            ''End If


                            '.Col = RateWiseDtlGrid.Col_Select
                            '.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                            '.Value = row("SELECT")

                            LockUnlockRateWiseGrid(.Row, .Row, True)  'lock grid

                        Next
                        GetTotalBasicValEffect()
                        optMultipleInvoice.Checked = True
                    End With
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    ' Code Added By Mayur Against issue ID 10816097    
    ' Negative Taxes


    Private Sub BtnFetchItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFetchItem.Click
        Dim strQuery As String = String.Empty
        Dim strHelp As String()
        Dim selectedItems As String = String.Empty
        Try
            If optsingleInvoice.Checked = True Or optMultipleInvoice.Checked = True Then
                ClearTmpTable()

                If cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    strQuery = ""
                Else
                    'If Not IsNothing(dtSelItems) Then
                    If opt_all_SO.Checked = True Then
                        strQuery = "Select Cust_Item_code,Item_code,cast(Cust_Item_desc as varchar(50)) as Cust_Item_desc  from Saleschallan_dtl a (nolock) ,sales_dtl b( nolock) " & _
                        " where a.Unit_code = b.Unit_code and a.doc_no= b.doc_no and a.Unit_code='" & gstrUNITID & "' and  Account_code = '" & txtCustCode.Text & "' and " & _
                        " invoice_date > = '" & Format(dtpFrm.Value, "dd MMM yyyy") & "' and invoice_date < = '" & Format(dtpToDt.Value, "dd MMM yyyy") & "' and " & _
                        " a.Doc_no = b.Doc_no and bill_flag = 1 and cancel_flag =0 group by Cust_Item_code,Item_code,cast(Cust_Item_desc as varchar(50)) "
                    Else
                        If mstrsalesorder = "" Then
                            MsgBox("Please select Sales Order.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                            Exit Sub
                        Else
                            strQuery = "Select Cust_Item_code,Item_code,cast(Cust_Item_desc as varchar(50)) as Cust_Item_desc  from Saleschallan_dtl a (nolock) ,sales_dtl b( nolock) " & _
                            " where a.Unit_code = b.Unit_code and a.doc_no= b.doc_no and a.Unit_code='" & gstrUNITID & "' and  Account_code = '" & txtCustCode.Text & "' and " & _
                            " invoice_date > = '" & Format(dtpFrm.Value, "dd MMM yyyy") & "' and invoice_date < = '" & Format(dtpToDt.Value, "dd MMM yyyy") & "' and " & _
                            " a.Doc_no = b.Doc_no and bill_flag = 1 and cancel_flag =0  and a.cust_ref in (" & mstrsalesorder & ") group by Cust_Item_code,Item_code,cast(Cust_Item_desc as varchar(50)) "
                        End If
                    End If
                    
                    If Not IsNothing(dtSelItems) Then

                        If dtSelItems.Rows.Count > 0 Then
                            For Each dr As DataRow In dtSelItems.Rows
                                selectedItems = selectedItems + dr("ITEM_CODE") + "|"

                            Next
                            selectedItems = selectedItems.Remove(selectedItems.LastIndexOf("|"), 1)
                            cmdSelectInvoice.Enabled = True

                        End If
                    End If
                    With ctlHelp
                        .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                        .ConnectAsUser = gstrCONNECTIONUSER
                        .ConnectThroughDSN = gstrCONNECTIONDSN
                        .ConnectWithPWD = gstrCONNECTIONPASSWORD
                    End With

                    strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Select Part(s)", , 1)
                    If Not IsNothing(strHelp) Then
                        strQuery = String.Empty
                        If strHelp.Length = 4 Then
                            If Not IsNothing(strHelp(3)) Then
                                strQuery = strHelp(3).Replace("|", "','")
                            End If
                        End If
                        SqlConnectionclass.ExecuteNonQuery("INSERT INTO TMP_SUPPLEMENETARYINV_DATA(CUST_ITEM_CODE,ITEM_CODE,IP_ADDRESS,UNIT_CODE)SELECT cust_drgno ,item_code,'" + gstrIpaddressWinSck + "','" + gstrUNITID + "'FROM CUSTITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE ='" & txtCustCode.Text & "' AND CUST_DRGNO IN('" + strQuery + "')")

                        strQuery = "Select DISTINCT Cust_Item_code,Item_code,cast(Cust_Item_desc as varchar(50)) as Cust_Item_desc  from Saleschallan_dtl a,sales_dtl b " & _
                            " where a.Unit_code = b.Unit_code and a.Unit_code='" & gstrUNITID & "' and  Account_code = '" & txtCustCode.Text & "' and " & _
                            " invoice_date > = '" & Format(dtpFrm.Value, "dd MMM yyyy") & "' and invoice_date < = '" & Format(dtpToDt.Value, "dd MMM yyyy") & "' and " & _
                            " a.Doc_no = b.Doc_no and bill_flag = 1 and cancel_flag =0 and b.cust_item_code in('" + strQuery + "')"
                    Else
                        strQuery = String.Empty
                        Exit Sub
                    End If
                End If
                dtSelItems = New DataTable
                dtSelItems = SqlConnectionclass.GetDataTable(strQuery)
                If strQuery.Length = 0 Then
                    Exit Sub
                End If
                If (Not IsNothing(dtSelItems)) Then
                    If (dtSelItems.Rows.Count = 0) Then
                        MessageBox.Show("No Record found.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Exit Sub
                    Else
                        cmdSelectInvoice.Enabled = True
                    End If
                End If
                clearFarGrid()
                bindRateItemfpGrid()
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub cmdSelectInvoice_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSelectInvoice.Click
        Dim objTempRs As New ADODB.Recordset
        If Len(Trim(txtCustCode.Text)) = 0 Then
            MsgBox("Please select the customer part code.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            Exit Sub
        End If
        If cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            If FillInvoiceNumber() = True Then
                lvwsalesorderlist.Enabled = False
                lvwsalesorderlist.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED_LIST_VIEW)
                fraInvoice.Left = VB6.TwipsToPixelsX(2625)
                fraInvoice.Width = VB6.TwipsToPixelsX(4215)
                fraInvoice.Visible = True
                fraInvoice.BringToFront()
            Else
                lvwsalesorderlist.Enabled = True
                lvwsalesorderlist.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                MsgBox("No invoice for the selected Customer part.", MsgBoxStyle.Information, "Empower")
                Exit Sub
            End If
        End If
    End Sub
    Private Function FillInvoiceNumber() As Boolean

        Dim strSQL As String
        Dim rsSalesDtl As New ClsResultSetDB
        Dim rsLastSupplementary As ClsResultSetDB
        Dim intMaxRow As Short
        Dim Intcounter As Short
        Dim intListCounter As Short
        Dim intListIndex As Short
        Dim sarrSelInvoices() As String
        Dim SOSTYPE As String
        lstInv.Items.Clear()
        lstInv.Columns.Item(0).Width = VB6.TwipsToPixelsX(2000)
        lstInv.Columns.Item(1).Width = VB6.TwipsToPixelsX(1000)

        lstInv.View = System.Windows.Forms.View.Details
        If cmdGrpSalesProv.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            sarrSelInvoices = Split(strSelInvoices, ",")
            intMaxRow = UBound(sarrSelInvoices) + 1
            lstInv.Columns.Item(1).Width = 0
            For Intcounter = 0 To intMaxRow - 1
                If sarrSelInvoices(Intcounter) <> "" Then
                    lstInv.Items.Insert(Intcounter, sarrSelInvoices(Intcounter))
                    lstInv.Items.Item(Intcounter).Checked = True
                End If
            Next Intcounter
            FillInvoiceNumber = True
            Exit Function
        End If
        If opt_all_SO.Checked = True Then
            SOSTYPE = "A"
        Else
            SOSTYPE = "S"
        End If
        strSQL = "select DISTINCT DOC_NO,INVOICE_DATE from (select * from SUPPLEMENTARYDATA_GST('" & gstrUNITID & "','" & Format(dtpFrm.Value, "dd MMM yyyy") & "','"
        strSQL = strSQL & Format(dtpToDt.Value, "dd MMM yyyy") & "','" & txtCustCode.Text & "',"
        strSQL = strSQL & "'" & cmdGrpSalesProv.Mode & "','" & txtDocNo.Text & "','" & SOSTYPE & "','" & gstrIpaddressWinSck & "' ))XYZ Order By Doc_No"
        rsSalesDtl.GetResult(strSQL)
        intMaxRow = rsSalesDtl.GetNoRows
        rsSalesDtl.MoveFirst()
        If intMaxRow > 0 Then
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            rsSalesDtl.MoveFirst()
            For Intcounter = 0 To intMaxRow - 1
                lstInv.Items.Insert(Intcounter, rsSalesDtl.GetValue("Doc_No"))
                If lstInv.Items.Item(Intcounter).SubItems.Count > 1 Then
                    lstInv.Items.Item(Intcounter).SubItems(1).Text = rsSalesDtl.GetValue("INVOICE_DATE")
                Else
                    lstInv.Items.Item(Intcounter).SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsSalesDtl.GetValue("INVOICE_DATE")))
                End If
                rsSalesDtl.MoveNext()
            Next Intcounter
            'If strSelInvoices <> "" Then
            '    chkUnCheckall.CheckState = System.Windows.Forms.CheckState.Checked
            '    chkUnCheckall.CheckState = System.Windows.Forms.CheckState.Unchecked
            '    chkSelectAll.CheckState = System.Windows.Forms.CheckState.Unchecked
            '    sarrSelInvoices = Split(strSelInvoices, ",")
            '    For intListCounter = 0 To UBound(sarrSelInvoices)
            '        If sarrSelInvoices(intListCounter) <> "" Then
            '            intListIndex = lstInv.FindItemWithText(sarrSelInvoices(intListCounter)).Index
            '            lstInv.Items.Item(intListIndex).Checked = True
            '        End If
            '    Next intListCounter
            'Else
            '    chkUnCheckall.CheckState = System.Windows.Forms.CheckState.Unchecked
            '    chkSelectAll.CheckState = System.Windows.Forms.CheckState.Checked
            'End If
            FillInvoiceNumber = True
        Else
            FillInvoiceNumber = False
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
    End Function

    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        fraInvoice.Visible = False
    End Sub
    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOk.Click
        Dim Intcounter As Short
        strSelInvoices = ""

        For Intcounter = 0 To lstInv.Items.Count - 1
            If lstInv.Items.Item(Intcounter).Checked = True Then
                strSelInvoices = strSelInvoices & "," & lstInv.Items.Item(Intcounter).Text
            Else
                strNotSelInvoices = strNotSelInvoices & "," & lstInv.Items.Item(Intcounter).Text
            End If
        Next Intcounter
        If VB.Left(strSelInvoices, 1) = "," Then
            strSelInvoices = VB.Right(strSelInvoices, Len(strSelInvoices) - 1)
        End If
        If VB.Left(strNotSelInvoices, 1) = "," Then
            strNotSelInvoices = VB.Right(strNotSelInvoices, Len(strNotSelInvoices) - 1)
        End If
        If strSelInvoices = "" Then
            MessageBox.Show("Select atleast Invoice .", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
            fraInvoice.Visible = False
            ClearTmpTable()
            clearFarGrid()
            ClearAllTaxes()
            BtnFetchItem.Enabled = True
        Else
            BtnFetchItem.Enabled = False
            fraInvoice.Visible = False
            Call FetchinvoiceRecords()
        End If
    End Sub

    Private Sub chkSelectAll_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkSelectAll.CheckStateChanged
        Dim intLoopCounter As Integer
        Dim intMaxLoop As Integer
        On Error GoTo ErrHandler
        If bool_check = True Then
            bool_check = False
            Exit Sub
        End If
        If chkSelectAll.CheckState = System.Windows.Forms.CheckState.Checked Then
            chkUnCheckall.CheckState = System.Windows.Forms.CheckState.Unchecked
            intMaxLoop = lstInv.Items.Count
            For intLoopCounter = 0 To intMaxLoop - 1

                With lstInv
                    bool_check = True
                    .Items.Item(intLoopCounter).Checked = True
                    bool_check = False
                End With

            Next
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub chkUnCheckall_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkUnCheckall.CheckStateChanged
        Dim intLoopCounter As Integer
        Dim intMaxLoop As Integer
        Dim lstItm As New ListViewItem
        On Error GoTo ErrHandler

        If bool_check = True Then
            bool_check = False
            Exit Sub
        End If
        If chkUnCheckall.CheckState = System.Windows.Forms.CheckState.Checked Then
            chkSelectAll.CheckState = System.Windows.Forms.CheckState.Unchecked
            intMaxLoop = lstInv.Items.Count
            'For intLoopCounter = 0 To intMaxLoop - 1
            '    With lstInv
            '        bool_check = True
            '        .Items.Item(intLoopCounter).Checked = False
            '        bool_check = False
            '    End With
            'Next
            For Each lstItm In lstInv.Items
                If IsNothing(lstItm) = True Then Continue For
                bool_check = True
                lstItm.Checked = False
                bool_check = False
            Next
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub FN_Spread_Settings()
        Dim Col As Short
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        With SSinvoiceDetail

            .MaxCols = ENUM_Gridinvoicedetails.IGST_TAX
            .MaxRows = 0
            .Row = 0
            .Col = ENUM_Gridinvoicedetails.BASE_INVOICENO : .Text = "BASE INV NO"
            .Col = ENUM_Gridinvoicedetails.CURR_INVOICENO : .Text = "INVOICE NO" ': .ColHidden = True
            .Col = ENUM_Gridinvoicedetails.InternalPartNo : .Text = "INT.PARTNO"
            .Col = ENUM_Gridinvoicedetails.CustPartNo : .Text = "CUSTOMER PARTNO"
            .Col = ENUM_Gridinvoicedetails.ItemDesc : .Text = "ITEM DESC"
            .Col = ENUM_Gridinvoicedetails.CustPartDesc : .Text = "CUSTPART DESC"
            .Col = ENUM_Gridinvoicedetails.SalesOrder : .Text = "SALES ORDER"
            .Col = ENUM_Gridinvoicedetails.Amendmentno : .Text = "AMENDMENT NO"
            .Col = ENUM_Gridinvoicedetails.InvoiceQty : .Text = "INVOICE QTY"
            .Col = ENUM_Gridinvoicedetails.Rate : .Text = "INVOICE RATE"
            .Col = ENUM_Gridinvoicedetails.NewRate : .Text = "NEW INVOICE RATE"
            .Col = ENUM_Gridinvoicedetails.TaxableAmount : .Text = "TAXABLE AMOUNT"
            .Col = ENUM_Gridinvoicedetails.HSNSACCODE : .Text = "HSN/SACCODE"
            .Col = ENUM_Gridinvoicedetails.CGST_TAX_TYPE : .Text = "CGST TYPE"
            .Col = ENUM_Gridinvoicedetails.CGST_TAX : .Text = "CGST PERCENT"
            .Col = ENUM_Gridinvoicedetails.SGST_TAX_TYPE : .Text = "SGST TYPE"
            .Col = ENUM_Gridinvoicedetails.SGST_TAX : .Text = "SGST PERCENT"
            .Col = ENUM_Gridinvoicedetails.IGST_TAX_TYPE : .Text = "IGST TYPE"
            .Col = ENUM_Gridinvoicedetails.IGST_TAX : .Text = "IGST PERCENT"

            .Row = -1
            .Col = ENUM_Gridinvoicedetails.BASE_INVOICENO
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Gridinvoicedetails.CURR_INVOICENO
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText

            .Col = ENUM_Gridinvoicedetails.InternalPartNo
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Gridinvoicedetails.CustPartNo
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Gridinvoicedetails.ItemDesc
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Gridinvoicedetails.CustPartDesc
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Gridinvoicedetails.SalesOrder
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Gridinvoicedetails.Amendmentno
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Gridinvoicedetails.InvoiceQty
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Gridinvoicedetails.Rate
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Gridinvoicedetails.TaxableAmount
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Gridinvoicedetails.CGST_TAX
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Gridinvoicedetails.SGST_TAX
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Gridinvoicedetails.IGST_TAX
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .set_RowHeight(0, 20)

            .set_ColWidth(ENUM_Gridinvoicedetails.BASE_INVOICENO, 10)
            .set_ColWidth(ENUM_Gridinvoicedetails.CURR_INVOICENO, 10)
            .set_ColWidth(ENUM_Gridinvoicedetails.InternalPartNo, 14)
            .set_ColWidth(ENUM_Gridinvoicedetails.CustPartNo, 14)
            .set_ColWidth(ENUM_Gridinvoicedetails.ItemDesc, 14)
            .set_ColWidth(ENUM_Gridinvoicedetails.CustPartDesc, 14)
            .set_ColWidth(ENUM_Gridinvoicedetails.SalesOrder, 0)
            .set_ColWidth(ENUM_Gridinvoicedetails.Amendmentno, 0)
            .set_ColWidth(ENUM_Gridinvoicedetails.InvoiceQty, 6)
            .set_ColWidth(ENUM_Gridinvoicedetails.Rate, 10)
            .set_ColWidth(ENUM_Gridinvoicedetails.NewRate, 0)
            .set_ColWidth(ENUM_Gridinvoicedetails.TaxableAmount, 10)
            .set_ColWidth(ENUM_Gridinvoicedetails.HSNSACCODE, 14)
            .set_ColWidth(ENUM_Gridinvoicedetails.CGST_TAX, 6)
            .set_ColWidth(ENUM_Gridinvoicedetails.CGST_TAX_TYPE, 8)
            .set_ColWidth(ENUM_Gridinvoicedetails.SGST_TAX, 7)
            .set_ColWidth(ENUM_Gridinvoicedetails.SGST_TAX_TYPE, 8)
            .set_ColWidth(ENUM_Gridinvoicedetails.IGST_TAX, 7)
            .set_ColWidth(ENUM_Gridinvoicedetails.IGST_TAX_TYPE, 8)

            .Row = 1
            .Row = .MaxRows
            .Col = ENUM_Gridinvoicedetails.BASE_INVOICENO
            .Col2 = .MaxCols
            .Lock = True
            .BlockMode = True
        End With

        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Description, Err.Source, mP_Connection)
        Exit Sub
    End Sub

    Private Sub fpSpreadRateWiseDtl_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles fpSpreadRateWiseDtl.Validating
        GetTotalBasicValEffect()
    End Sub
    Private Sub opt_all_SO_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles opt_all_SO.CheckedChanged
        If sender.Checked Then
            On Error GoTo ErrHandler
            Me.opt_all_SO.Checked = True
            Me.lvwsalesorderlist.Enabled = False
            Me.lvwsalesorderlist.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED_LIST_VIEW)
            Me.lvwsalesorderlist.Items.Clear()
            Me.lvwsalesorderlist.Columns.Clear()
            Me.lvwsalesorderlist.GridLines = False
            Me.lvwsalesorderlist.Enabled = False
            Me.opt_all_SO.Checked = True
            clearFarGrid()
            BtnFetchItem.Enabled = True
            cmdSelectInvoice.Enabled = False
            Exit Sub
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub

    Private Sub opt_sel_SO_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles opt_sel_SO.CheckedChanged
        If sender.Checked Then
            On Error GoTo ErrHandler
            Me.lvwsalesorderlist.Enabled = True
            Me.lvwsalesorderlist.View = System.Windows.Forms.View.Details
            Me.lvwsalesorderlist.CheckBoxes = True
            Me.lvwsalesorderlist.GridLines = True
            Me.lvwsalesorderlist.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            clearFarGrid()
            Call Populatesalesorder()
            BtnFetchItem.Enabled = True
            cmdSelectInvoice.Enabled = False
            Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If

    End Sub

    Private Sub Populatesalesorder()


        'Declarations
        Dim objsalesorder As ClsResultSetDB 'Class Object
        Dim strSQLSalesOrder As String 'Stores the SQL statement for getting the accounting locations
        Dim intLocationCount As Short 'Stores the total location count
        Dim lngsalesorderCtr As Integer 'For counter variable
        Dim objcustomers As ClsResultSetDB 'Class Object
        Dim strSQLcustomers As String 'Stores the SQL statement for getting the customers
        Dim intcustomerCount As Short 'Stores the total customer count
        Dim lngCustomerCtr As Integer

        Try

            lvwsalesorderlist.Items.Clear()
            With lvwsalesorderlist
                .Sort()
                ListViewColumnSorter.SortListView(lvwsalesorderlist, 0, SortOrder.Ascending)
                .LabelEdit = False
                .CheckBoxes = True
                .View = System.Windows.Forms.View.Details
                .Columns.Clear()
                .Columns.Insert(0, "", "SALES ORDER", 100)
                .Columns.Insert(1, "", "AMENDMENT NO", 189)
            End With

            strSQLSalesOrder = "SELECT distinct cust_ref ,amendment_no FROM saleschallan_Dtl WHERE unit_code = '" & gstrUNITID & "' and account_code='" & txtCustCode.Text.Trim & "'"
            strSQLSalesOrder += " and invoice_date > = '" & Format(dtpFrm.Value, "dd MMM yyyy") & "' and invoice_date < = '" & Format(dtpToDt.Value, "dd MMM yyyy") & "' and BILL_FLAG=1 AND CANCEL_FLAG=0 "

            objsalesorder = New ClsResultSetDB
            With objsalesorder
                Call .GetResult(strSQLSalesOrder, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                intLocationCount = .GetNoRows
                If txtCustCode.Text.Trim.Length = 0 Then
                    Call MsgBox("Please select Customer.", MsgBoxStyle.Information, "eMPro")
                    opt_all_SO.Checked = True
                    .ResultSetClose()
                    objsalesorder = Nothing
                    Exit Sub
                ElseIf intLocationCount <= 0 Then
                    Call MsgBox("No  Sales order  have been defined.", MsgBoxStyle.Information, "eMPro")
                    .ResultSetClose()
                    objsalesorder = Nothing
                    Exit Sub
                End If
                With lvwsalesorderlist
                    .Items.Clear()
                    objsalesorder.MoveFirst()
                    For lngsalesorderCtr = 0 To intLocationCount - 1
                        .Items.Insert(lngsalesorderCtr, CStr(objsalesorder.GetValue("cust_ref")))
                        If .Items.Item(lngsalesorderCtr).SubItems.Count > 1 Then
                            .Items.Item(lngsalesorderCtr).SubItems(1).Text = objsalesorder.GetValue("amendment_no").ToString
                        Else
                            .Items.Item(lngsalesorderCtr).SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, objsalesorder.GetValue("amendment_no")))
                        End If
                        objsalesorder.MoveNext()
                    Next
                End With
                .ResultSetClose()
                objsalesorder = Nothing
            End With
            Me.lvwsalesorderlist.Columns.Item(0).Width = 100
            Me.lvwsalesorderlist.Columns.Item(1).Width = 189
            Exit Sub
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub lvwsalesorderlist_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles lvwsalesorderlist.ItemChecked
        Dim Item As System.Windows.Forms.ListViewItem = lvwsalesorderlist.Items(e.Item.Index)
        Try

            Dim intLoopcount As Short
            Dim intMaxLoop As Short
            Dim inti As Short
            Dim strQry As String
            mstrsalesorder = ""
            intMaxLoop = lvwsalesorderlist.Items.Count

            strQry = "DELETE FROM TMP_SUPPLEMENETARYINV_SO_DATA WHERE UNIT_CODE='" + gstrUNITID + "' AND IP_ADDRESS='" + gstrIpaddressWinSck + "';"
            SqlConnectionclass.ExecuteNonQuery(strQry)
            For intLoopcount = 0 To intMaxLoop - 1
                If lvwsalesorderlist.Items.Item(intLoopcount).Checked = True Then
                    If Len(Trim(mstrsalesorder)) = 0 Then
                        mstrsalesorder = "'" & lvwsalesorderlist.Items.Item(intLoopcount).Text & "'"
                    Else
                        mstrsalesorder = mstrsalesorder & ",'" & lvwsalesorderlist.Items.Item(intLoopcount).Text & "'"
                    End If
                    SqlConnectionclass.ExecuteNonQuery("INSERT INTO TMP_SUPPLEMENETARYINV_SO_DATA(IP_ADDRESS,UNIT_CODE,SONO,AMENDMENT_NO,CUSTOMER_CODE) SELECT '" + gstrIpaddressWinSck + "','" + gstrUNITID + "','" & lvwsalesorderlist.Items.Item(intLoopcount).Text & "','" & lvwsalesorderlist.Items.Item(intLoopcount).SubItems(1).Text & "','" & txtCustCode.Text & "'")
                End If
            Next
            Exit Sub

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Public Function CalculateTCSTax(ByRef pdblTotalValue As Double, ByRef pblnTCSRoundOFF As Boolean, ByRef pintTCSPer As Double) As Double
        Dim dblTCSTax As Double
        Dim strsql As String
        Try
            If pblnTCSRoundOFF = True Then
                dblTCSTax = System.Math.Round((pdblTotalValue * pintTCSPer) / 100, 0)
            Else
                dblTCSTax = System.Math.Round((pdblTotalValue * pintTCSPer) / 100, 2)
            End If
            'CalculateTCSTax = dblTCSTax
            strsql = "select dbo.UFN_HIGHERVALUE_TCSROUNDING(" & dblTCSTax & ",'" & gstrUNITID & "'  )"
            CalculateTCSTax = SqlConnectionclass.ExecuteScalar(strsql)
            Exit Function

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function

    Private Sub BtnHelpTCS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnHelpTCS.Click
        Try
            Dim strSQL As String = String.Empty
            Dim strHelp As String()
            With ctlHelp
                .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                .ConnectAsUser = gstrCONNECTIONUSER
                .ConnectThroughDSN = gstrCONNECTIONDSN
                .ConnectWithPWD = gstrCONNECTIONPASSWORD
            End With

            'If optsingleInvoice.Checked = True Then
            strSQL = "Select DISTINCT TxRt_Rate_No,TxRt_Percentage from GEN_TAXRATE Where " & _
             " UNIT_CODE='" & gstrUNITID & "'" & _
            " and (Tx_TaxeID='TCS')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"

            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "TCS TAX")

            If Not IsNothing(strHelp) Then
                txtTCSTaxCode.Text = strHelp(0)
                lblTCSTaxPerDes.Text = strHelp(1)
            End If
            'End If


        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub OptDrNote_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptDrNote.CheckedChanged
        Dim STRDRCR As String
        Dim strSQL As String
        Try

        
            If Len(Trim(txtCustCode.Text)) > 0 Then
                If OptDrNote.Checked = True Then
                    STRDRCR = "DR"
                Else
                    STRDRCR = "CR"
                End If
                strSQL = "select dbo.UDF_IRN_TCSREQUIRED_SUPP( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & STRDRCR & "')"
                If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
                    txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtTCSTaxCode.Text = "" : BtnHelpTCS.Enabled = True : lblTCSTaxPerDes.Text = ""
                Else
                    txtTCSTaxCode.Enabled = False : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtTCSTaxCode.Text = "" : BtnHelpTCS.Enabled = False : lblTCSTaxPerDes.Text = ""
                End If

            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub OptCrNote_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptCrNote.CheckedChanged
        Dim STRDRCR As String
        Dim strSQL As String
        Try


            If Len(Trim(txtCustCode.Text)) > 0 Then
                If OptDrNote.Checked = True Then
                    STRDRCR = "DR"
                Else
                    STRDRCR = "CR"
                End If
                strSQL = "select dbo.UDF_IRN_TCSREQUIRED_SUPP( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & STRDRCR & "')"
                If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
                    txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtTCSTaxCode.Text = "" : BtnHelpTCS.Enabled = True : lblTCSTaxPerDes.Text = ""
                Else
                    txtTCSTaxCode.Enabled = False : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtTCSTaxCode.Text = "" : BtnHelpTCS.Enabled = False : lblTCSTaxPerDes.Text = ""
                End If

            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub cmdGrpSalesProv_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGrpSalesProv.Load

    End Sub
End Class