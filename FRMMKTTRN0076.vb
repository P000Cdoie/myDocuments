'========================================================================================
'Copyright          :   Mothersonsumi Infotech & Design Ltd.
'Module             :   Marketing
'Author             :   Amit Kumar [0670]
'Creation Date      :   19 Jul 2012
'Description        :   Form For Trading Invoice Entry.
'========================================================================================
Imports System.Data.SqlClient
Imports System.Text
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class FRMMKTTRN0076
    Dim mstrRefNo As String = String.Empty
    Dim mstrAmmNo As String
    Dim strInvType As String
    Dim strInvSubType As String
    Dim mstrLocationCode As String
    Dim mblnServiceInvoiceWithoutSO As Boolean
    Dim mblnSORequired As Boolean
    Dim mblnMultipleSOAllowed As Boolean
    Dim mstrCreditTermId As String
    Dim mblnCheckArray As Boolean
    Dim mstrStockLocation As String
    Dim mstrInvTypenew As String
    Dim mstrInvSubTypenew As String
    Dim mblnRejTracking As Boolean
    Dim mblnBatchTracking As Boolean
    Dim mblnbatchfifomode As Boolean
    Dim mblnBatchTrack As Boolean
    Dim mBatchData() As Batch_Details
    Dim mdblPrevQty() As Object
    Dim mdblToolCost() As Object
    Dim strBarCodeMainString() As String
    Dim strBarCodeSplitedString(10, 4) As String
    Dim strBRString As String
    Dim mstrInvoiceType As String
    Const MaxHdrGridCols As Short = 30

    '======================RoundoffSetting=====================
    Dim blnISInsExcisable As Boolean
    Dim blnEOUFlag As Boolean
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
    Dim intUNLOCKED_INVOICES As Integer
    '======================RoundoffSetting=====================

    Dim strSOSaleTaxType As String = String.Empty
    Dim blnBillFlag As Boolean = False
    Dim dblBasicValue As Decimal
    Public mstrItemCode As String
    Public dblGrinQuantityForSale As Double
    Public strGrinAllocationOKCancel As Boolean = False


    Private Enum EnumInv
        ENUMITEMCODE = 1
        CUSTPARTNO = 2
        RATE_PERUNIT = 3
        CUSTSUPPMAT_PERUNIT = 4
        SelectGrin = 5
        ENUMQUANTITY = 6
        PACKING = 7
        OTHERS_PERUNIT = 8
        FROMBOX = 9
        TOBOX = 10
        CUMULATIVEBOXES = 11
        BINQTY = 12
        BATCHCOL = 13
    End Enum
#Region "Functions And Subs"
    Private Sub GetDefaultTaxexFromSO()

        Dim Sqlcmd As New SqlCommand
        Dim strSql As String
        Dim ObjVal As Object = Nothing
        Dim SQLRd As SqlDataReader
        Dim builder As New StringBuilder

        Dim strDespatchQty As String = String.Empty
        Dim strCustDrgno As String = String.Empty
        Dim strInternalCode As String = String.Empty
        Try
            Sqlcmd.CommandTimeout = 0
            Sqlcmd.Connection = SqlConnectionclass.GetConnection
            Sqlcmd.CommandType = CommandType.Text
            If strSOSaleTaxType.Trim <> String.Empty Then
                Sqlcmd.CommandText = "SELECT TXRT_PERCENTAGE FROM GEN_TAXRATE WHERE UNIT_CODE='" + gstrUNITID + "' And LTRIM(RTRIM(TXRT_RATE_NO))='" + strSOSaleTaxType.Trim + "' AND (TX_TAXEID='CSTT' OR TX_TAXEID='LSTT' OR TX_TAXEID='VATT')  AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) OR (CAST(GETDATE() AS DATE) <= DEACTIVE_DATE))"
                ObjVal = Sqlcmd.ExecuteScalar()
                If IsNothing(ObjVal) = True Then ObjVal = String.Empty
                If ObjVal.ToString <> String.Empty Then
                    Me.txtSaleTaxType.Text = strSOSaleTaxType
                    Me.lblSaltax_Per.Text = ObjVal.ToString
                End If
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            If IsNothing(Sqlcmd.Connection.State) = False Then
                If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
            End If
            If IsNothing(Sqlcmd) = False Then
                Sqlcmd.Connection.Dispose()
                Sqlcmd.Dispose()
            End If
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
                MsgBox("No Data Define In Sales_Parameter Table", MsgBoxStyle.Critical, ResolveResString(100))
                Exit Sub
            End If
            If DataRd.IsClosed = False Then DataRd.Close()
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
    Private Sub GetSavedData(ByVal strInvoiceNo As String)
        Dim Sqlcmd As New SqlCommand
        Dim ObjVal As Object = Nothing
        Dim intX As Int16
        Dim builder As New StringBuilder
        Dim SQLRd As SqlDataReader
        Try

            BlankTaxDetails()
            Sqlcmd.CommandTimeout = 0
            Sqlcmd.Connection = SqlConnectionclass.GetConnection
            Sqlcmd.CommandType = CommandType.Text

            builder.Remove(0, builder.ToString.Length)
            builder.AppendLine(" SELECT  ")
            builder.AppendLine("Account_Code,")
            builder.AppendLine("CONSIGNEE_CODE,")
            builder.AppendLine("Bill_Flag,")
            builder.AppendLine("FIFO_flag,")
            builder.AppendLine("Print_Flag,")
            builder.AppendLine("Cancel_flag,")
            builder.AppendLine("dataPosted,")
            builder.AppendLine("ftp,")
            builder.AppendLine("ExciseExumpted,")
            builder.AppendLine("ServiceInvoiceformatExport,")
            builder.AppendLine("Discount_Type,")
            builder.AppendLine("RejectionPosting,")
            builder.AppendLine("FOC_Invoice,")
            builder.AppendLine("PrintExciseFormat,")
            builder.AppendLine("FreshCrRecd,")
            builder.AppendLine("Trans_Parameter_Flag,")
            builder.AppendLine("InvoiceAgainstMultipleSO,")
            builder.AppendLine("TextFileGenerated,")
            builder.AppendLine("sameunitloading,")
            builder.AppendLine("postingFlag,")
            builder.AppendLine("MULTIPLESO,")
            builder.AppendLine("Suffix,")
            builder.AppendLine("Transport_Type,")
            builder.AppendLine("Invoice_Type,")
            builder.AppendLine("Sub_Category,")
            builder.AppendLine("SalesTax_Type,")
            builder.AppendLine("SalesTax_FormNo,")
            builder.AppendLine("other_ref,")
            builder.AppendLine("Surcharge_salesTaxType,")
            builder.AppendLine("LoadingChargeTaxType,")
            builder.AppendLine("CustBankID,")
            builder.AppendLine("USLOC,")
            builder.AppendLine("TCSTax_Type,")
            builder.AppendLine("ECESS_Type,")
            builder.AppendLine("SRCESS_Type,")
            builder.AppendLine("CVDCESS_Type,")
            builder.AppendLine("TurnOverTaxType,")
            builder.AppendLine("SDTax_Type,")
            builder.AppendLine("ServiceTax_Type,")
            builder.AppendLine("SECESS_Type,")
            builder.AppendLine("CVDSECESS_Type,")
            builder.AppendLine("SRSECESS_Type,")
            builder.AppendLine("ISCHALLAN,")
            builder.AppendLine("ISCONSOLIDATE,")
            builder.AppendLine("Ecess_TotalDuty_Type,")
            builder.AppendLine("SEcess_TotalDuty_Type,")
            builder.AppendLine("ADDVAT_Type,")
            builder.AppendLine("Invoice_Date,")
            builder.AppendLine("Form3Date,")
            builder.AppendLine("Exchange_Date,")
            builder.AppendLine("Ent_dt,")
            builder.AppendLine("Upd_dt,")
            builder.AppendLine("Doc_No,")
            builder.AppendLine("NRGPNOIncaseOfServiceInvoice,")
            builder.AppendLine("Discount_Per,")
            builder.AppendLine("Year,")
            builder.AppendLine("Annex_no,")
            builder.AppendLine("pervalue,")
            builder.AppendLine("dataposted_fin,")
            builder.AppendLine("Location_Code,")
            builder.AppendLine("To_Location,")
            builder.AppendLine("From_Location,")
            builder.AppendLine("Insurance,")
            builder.AppendLine("Frieght_Tax,")
            builder.AppendLine("Sales_Tax_Amount,")
            builder.AppendLine("Surcharge_Sales_Tax_Amount,")
            builder.AppendLine("Frieght_Amount,")
            builder.AppendLine("Packing_Amount,")
            builder.AppendLine("SalesTax_FormValue,")
            builder.AppendLine("total_amount,")
            builder.AppendLine("TurnoverTax_per,")
            builder.AppendLine("Turnover_amt,")
            builder.AppendLine("LoadingChargeTaxAmount,")
            builder.AppendLine("Discount_Amount,")
            builder.AppendLine("TCSTaxAmount,")
            builder.AppendLine("ECESS_Amount,")
            builder.AppendLine("SRCESS_Amount,")
            builder.AppendLine("CVDCESS_Amount,")
            builder.AppendLine("TotalInvoiceAmtRoundOff_diff,")
            builder.AppendLine("SDTax_Amount,")
            builder.AppendLine("ServiceTax_Amount,")
            builder.AppendLine("Prev_Yr_ExportSales,")
            builder.AppendLine("Permissible_Limit_SmpExport,")
            builder.AppendLine("SECESS_Amount,")
            builder.AppendLine("CVDSECESS_Amount,")
            builder.AppendLine("SRSECESS_Amount,")
            builder.AppendLine("Tot_Add_Excise_Amt,")
            builder.AppendLine("Ecess_TotalDuty_Amount,")
            builder.AppendLine("SEcess_TotalDuty_Amount,")
            builder.AppendLine("ADDVAT_Amount,")
            builder.AppendLine("Exchange_Rate,")
            builder.AppendLine("SalesTax_Per,")
            builder.AppendLine("Surcharge_SalesTax_Per,")
            builder.AppendLine("LoadingChargeTax_Per,")
            builder.AppendLine("TCSTax_Per,")
            builder.AppendLine("ECESS_Per,")
            builder.AppendLine("SRCESS_Per,")
            builder.AppendLine("CVDCESS_Per,")
            builder.AppendLine("Excise_Percentage,")
            builder.AppendLine("Permissible_Limit,")
            builder.AppendLine("SDTax_Per,")
            builder.AppendLine("ServiceTax_Per,")
            builder.AppendLine("SECESS_Per,")
            builder.AppendLine("CVDSECESS_Per,")
            builder.AppendLine("SRSECESS_Per,")
            builder.AppendLine("Tot_Add_Excise_PER,")
            builder.AppendLine("bond17OpeningBal,")
            builder.AppendLine("Ecess_TotalDuty_Per,")
            builder.AppendLine("SEcess_TotalDuty_Per,")
            builder.AppendLine("ADDVAT_Per,")
            builder.AppendLine("total_quantity,")
            builder.AppendLine("Ent_UserId,")
            builder.AppendLine("Upd_Userid,")
            builder.AppendLine("Vehicle_No,")
            builder.AppendLine("From_Station,")
            builder.AppendLine("To_Station,")
            builder.AppendLine("Cust_Ref,")
            builder.AppendLine("Amendment_No,")
            builder.AppendLine("Print_DateTime,")
            builder.AppendLine("Form3,")
            builder.AppendLine("Carriage_Name,")
            builder.AppendLine("Ref_Doc_No,")
            builder.AppendLine("Cust_Name,")
            builder.AppendLine("Currency_Code,")
            builder.AppendLine("Nature_of_Contract,")
            builder.AppendLine("OriginStatus,")
            builder.AppendLine("Ctry_Destination_Goods,")
            builder.AppendLine("Delivery_Terms,")
            builder.AppendLine("Payment_Terms,")
            builder.AppendLine("Pre_Carriage_By,")
            builder.AppendLine("Receipt_Precarriage_at,")
            builder.AppendLine("Vessel_Flight_number,")
            builder.AppendLine("Port_Of_Loading,")
            builder.AppendLine("Port_Of_Discharge,")
            builder.AppendLine("Final_destination,")
            builder.AppendLine("Mode_Of_Shipment,")
            builder.AppendLine("Dispatch_mode,")
            builder.AppendLine("Buyer_description_Of_Goods,")
            builder.AppendLine("Invoice_description_of_EPC,")
            builder.AppendLine("Buyer_Id,")
            builder.AppendLine("remarks,")
            builder.AppendLine("Excise_Type,")
            builder.AppendLine("SRVDINO,")
            builder.AppendLine("SRVLocation,")
            builder.AppendLine("ConsigneeContactPerson,")
            builder.AppendLine("ConsigneeAddress1,")
            builder.AppendLine("ConsigneeAddress2,")
            builder.AppendLine("ConsigneeAddress3,")
            builder.AppendLine("ConsigneeECCNo,")
            builder.AppendLine("ConsigneeLST,")
            builder.AppendLine("SchTime,")
            builder.AppendLine("invoice_time,")
            builder.AppendLine("varGeneralRemarks,")
            builder.AppendLine("CheckSheetNo,")
            builder.AppendLine("Lorry_No,")
            builder.AppendLine("OTL_No,")
            builder.AppendLine("RefChallan,")
            builder.AppendLine("price_bases,")
            builder.AppendLine("LorryNo_date,")
            builder.AppendLine("ConsInvString,")
            builder.AppendLine("invoicepicking_status,")
            builder.AppendLine("BasicExciseAndCessValue,")
            builder.AppendLine("TMP_DOC_No,")
            builder.AppendLine("'PAYMENT_TERMS_DESC'= ISNULL((SELECT CRTRM_DESC FROM GEN_CREDITTRMMASTER WHERE UNIT_CODE='" + gstrUNITID + "'  AND CRTRM_STATUS=1 AND GEN_CREDITTRMMASTER.CRTRM_TERMID=SALESCHALLAN_DTL.PAYMENT_TERMS ),''),")
            builder.AppendLine("UNIT_CODE")
            builder.AppendLine(" FROM SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "'  AND DOC_NO=" + Val(strInvoiceNo).ToString)

            Sqlcmd.CommandText = builder.ToString
            SQLRd = Sqlcmd.ExecuteReader
            If SQLRd.HasRows Then
                SQLRd.Read()

                Me.txtSaleTaxType.Text = SQLRd("SalesTax_Type").ToString
                Me.txtECSSTaxType.Text = SQLRd("ECESS_Type").ToString
                Me.txtSECSSTaxType.Text = SQLRd("SECESS_Type").ToString
                Me.dtpDateDesc.Value = Convert.ToDateTime(SQLRd("Invoice_Date")).ToString("dd/MMM/yyyy")
                Me.txtChallanNo.Text = strInvoiceNo
                Me.txtDiscountAmt.Text = SQLRd("Discount_Per").ToString
                Me.txtLocationCode.Text = SQLRd("Location_Code").ToString
                Me.ctlInsurance.Text = SQLRd("Insurance").ToString
                Me.lblSalesTaxValue.Text = SQLRd("Sales_Tax_Amount").ToString
                Me.txtFreight.Text = SQLRd("Frieght_Amount").ToString
                Me.LblNetInvoiceValue.Text = SQLRd("total_amount").ToString
                Me.txtDiscountAmt.Text = SQLRd("Discount_Amount").ToString
                If SQLRd("Discount_Type") = True And Val(Me.txtDiscountAmt.Text) > 0 Then
                    OptDiscountPercentage.Checked = True
                    OptDiscountValue.Checked = False
                ElseIf SQLRd("Discount_Type") = False And Val(Me.txtDiscountAmt.Text) > 0 Then
                    OptDiscountPercentage.Checked = False
                    OptDiscountValue.Checked = True
                Else
                    OptDiscountPercentage.Checked = False
                    OptDiscountValue.Checked = False
                End If

                For intX = 0 To Me.CmbTransType.Items.Count - 1
                    If Mid(CmbTransType.Items(intX).ToString, 1, 1) = SQLRd("Transport_Type").ToString Then
                        CmbTransType.SelectedIndex = intX
                    End If
                Next
                Me.txtRemarks.Text = SQLRd("REMARKS").ToString
                Me.lblBasicExciseAndCess.Text = SQLRd("BasicExciseAndCessValue").ToString
                Me.lblEcessValue.Text = SQLRd("ECESS_Amount").ToString
                Me.lblHCessValue.Text = SQLRd("SECESS_AMOUNT").ToString
                Me.lblSaltax_Per.Text = SQLRd("SalesTax_Per").ToString
                Me.lblECSStax_Per.Text = SQLRd("ECESS_Per").ToString
                Me.lblSECSStax_Per.Text = SQLRd("SECESS_Per").ToString
                Me.txtVehNo.Text = SQLRd("Vehicle_No").ToString
                Me.txtRefNo.Text = SQLRd("Cust_Ref").ToString
                Me.txtAmendNo.Text = SQLRd("Amendment_No").ToString
                Me.lblCustCodeDes.Text = SQLRd("Cust_Name").ToString
                Me.txtCreditTerms.Text = SQLRd("Payment_Terms").ToString
                Me.lblCreditTermDesc.Text = SQLRd("PAYMENT_TERMS_DESC").ToString
                Me.txtCustCode.Text = SQLRd("Account_Code").ToString
                Me.txtCarrServices.Text = SQLRd("Carriage_Name").ToString
                Me.txtAddVAT.Text = SQLRd("ADDVAT_Type").ToString
                Me.lblAddVAT.Text = SQLRd("AddVat_Per").ToString
                Me.lblAddVATValue.Text = SQLRd("ADDVAT_Amount").ToString
                blnBillFlag = SQLRd("BILL_FLAG").ToString
                lblRoundOff.Text = Val(SQLRd("TotalInvoiceAmtRoundOff_diff").ToString).ToString

                If SQLRd("Cancel_Flag").ToString.ToUpper = "TRUE" Then
                    lblCancelledInvoice.Visible = True
                Else
                    lblCancelledInvoice.Visible = False
                End If

                
            End If
            If SQLRd.IsClosed = False Then SQLRd.Close()

            Dim strItemCode As String
            Sqlcmd.CommandText = "SELECT TotalExciseAmount,From_Box,To_Box,Item_Code,Rate,Sales_Tax,Excise_Tax,Basic_Amount,Accessible_amount,Discount_amt,Sales_Quantity,BinQuantity,Cust_Item_Code, isnull(ADD_EXCISE_AMOUNT, 0) ADD_EXCISE_AMOUNT FROM SALES_DTL WHERE UNIT_CODE='" + gstrUNITID + "'  AND DOC_NO=" + Val(strInvoiceNo).ToString
            SQLRd = Sqlcmd.ExecuteReader
            If SQLRd.HasRows Then
                SQLRd.Read()
                'addRowAtEnterKeyPress(0)
                SpChEntry.MaxRows = 0
                SpChEntry.MaxRows = SpChEntry.MaxRows + 1
                ChangeCellTypeStaticText()
                SpChEntry.Row = SpChEntry.MaxRows

                Me.lblExciseValue.Text = SQLRd("TotalExciseAmount").ToString
                Me.lblAEDValue.Text = SQLRd("ADD_EXCISE_AMOUNT").ToString
                SpChEntry.Col = EnumInv.FROMBOX
                SpChEntry.Text = SQLRd("From_Box").ToString

                SpChEntry.Col = EnumInv.TOBOX
                SpChEntry.Text = SQLRd("To_Box").ToString

                SpChEntry.Col = EnumInv.CUMULATIVEBOXES
                SpChEntry.Text = SQLRd("To_Box").ToString

                SpChEntry.Col = EnumInv.CUMULATIVEBOXES : SpChEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : SpChEntry.TypeFloatDecimalPlaces = 0 : SpChEntry.TypeFloatMin = CDbl("0.00") : SpChEntry.TypeFloatMax = CDbl("99999999999999.99") : SpChEntry.Lock = True
                SpChEntry.Col = EnumInv.FROMBOX : SpChEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : SpChEntry.TypeFloatDecimalPlaces = 0 : SpChEntry.TypeFloatMin = CDbl("0.00") : SpChEntry.TypeFloatMax = CDbl("999999.99") : SpChEntry.Lock = True
                SpChEntry.Col = EnumInv.TOBOX : SpChEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : SpChEntry.TypeFloatDecimalPlaces = 0 : SpChEntry.TypeFloatMin = CDbl("0.00") : SpChEntry.TypeFloatMax = CDbl("999999.99") : SpChEntry.Lock = True


                SpChEntry.Col = EnumInv.ENUMITEMCODE
                SpChEntry.Text = SQLRd("Item_Code").ToString
                strItemCode = SQLRd("Item_Code").ToString

                SpChEntry.Col = EnumInv.RATE_PERUNIT
                SpChEntry.Text = SQLRd("Rate").ToString

                Me.lblSaltax_Per.Text = SQLRd("Sales_Tax").ToString
                Me.lblExciseValue.Text = SQLRd("Excise_Tax").ToString
                Me.lblBasicValue.Text = SQLRd("Basic_Amount").ToString
                Me.dblBasicValue = SQLRd("Basic_Amount").ToString
                Me.lblAssValue.Text = SQLRd("Accessible_amount").ToString
                Me.txtDiscountAmt.Text = SQLRd("Discount_amt").ToString

                SpChEntry.Col = EnumInv.ENUMQUANTITY
                SpChEntry.Text = SQLRd("Sales_Quantity").ToString


                SpChEntry.Col = EnumInv.BINQTY
                SpChEntry.Text = SQLRd("BinQuantity").ToString

                SpChEntry.Col = EnumInv.CUSTPARTNO
                SpChEntry.Text = SQLRd("Cust_Item_Code").ToString

                CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
                CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = False


                'If blnBillFlag = True Then ' IF INVOICE IS LOCKED
                '    'CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                '    'CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                'End If

            End If

            If SQLRd.IsClosed = False Then SQLRd.Close()
            Sqlcmd.CommandText = "DELETE FROM TMP_TRADING_INV_GRINS WHERE UNIT_CODE='" + gstrUNITID + "' AND IPADDRESS='" + gstrIpaddressWinSck + "'"
            Sqlcmd.ExecuteNonQuery()
            Sqlcmd.CommandText = "INSERT INTO TMP_TRADING_INV_GRINS  (GRINNO,SLNO,GRIN_PAGE_NO,ITEM_CODE,SALESQTY,GRINQTY,REMQTY,KNOCKOFFQTY,PERPIECEEXCISE,IPADDRESS,UNIT_CODE,PerPieceAED) SELECT GRIN_NO,SLNO,GRIN_PAGE_NO,ITEM_CODE,SALESQTY,GRINQTY,REMQTY,KNOCKOFFQTY,PERPIECEEXCISE,'" + gstrIpaddressWinSck + "',UNIT_CODE, PerPieceAED FROM SALES_TRADING_GRIN_DTL WHERE UNIT_CODE='" + gstrUNITID + "'  AND DOC_NO=" + Val(strInvoiceNo).ToString
            Sqlcmd.ExecuteNonQuery()

            Sqlcmd.CommandText = "UPDATE A SET GRINDATE=B.GRN_DATE FROM TMP_TRADING_INV_GRINS A ,GRN_HDR B WHERE A.UNIT_CODE =B.UNIT_CODE AND A.GRINNO=B.DOC_NO  AND A.UNIT_CODE ='" + gstrUNITID + "' AND A.IPADDRESS ='" + gstrIpaddressWinSck + "'"
            Sqlcmd.ExecuteNonQuery()
            GetItemDescription()


        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            If IsNothing(Sqlcmd) = False Then
                If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
                Sqlcmd.Connection.Dispose()
                Sqlcmd.Dispose()
            End If
        End Try
    End Sub
    Public Function LockTradingInvoice(ByVal strInvoiceNo As String) As Boolean
        'ITEMBAL_MST            -- QUANTITY IS REDUCED FROM THIS TABLE WHILE LOKING THE INVOICE WHEN UPDATE STOCK FLAG IS TRUE IN SALESCONF.
        'CUST_ORD_DTL           -- QUANTITY IS KNOCKED OFF FROM THIS TABLE WHILE LOKING THE INVOICE WHEN UPDATE OP FLAG IS TRUE IN SALESCONF.
        'GRN_DTL                -- QUANTITY IS KNOCKED OFF FROM THIS TABLE WHILE LOKING THE INVOICE WHEN INVOICETYPE='NORMAL INVOICE' AND SUBTYPE='TRADING GOODS'
        'DAILYMKTSCHEDULE       -- QUANTITY IS KNOCKED OFF FROM THIS TABLE WHILE CREATING TRADING INVOICE.

        Dim Sqlcmd As New SqlCommand
        Dim strSql As String
        Dim ObjVal As Object = Nothing
        Dim SqlTrans As SqlTransaction
        Dim IsTrans As Boolean = False
        Dim SQLRd As SqlDataReader
        Dim builder As New StringBuilder

        Dim strDespatchQty As String = String.Empty
        Dim strCustDrgno As String = String.Empty
        Dim strInternalCode As String = String.Empty
        Try
            LockTradingInvoice = False
            Sqlcmd.CommandTimeout = 0

            Sqlcmd.Connection = SqlConnectionclass.GetConnection
            SqlTrans = Sqlcmd.Connection.BeginTransaction(System.Data.IsolationLevel.Serializable)
            Sqlcmd.Transaction = SqlTrans
            Sqlcmd.CommandType = CommandType.Text

            builder.Remove(0, builder.ToString.Length)
            builder.AppendLine("SELECT ITEM_CODE,CUST_ITEM_CODE,SALES_QUANTITY FROM SALES_DTL WHERE UNIT_CODE='" + gstrUNITID + "'  AND DOC_NO=" + Val(strInvoiceNo).ToString)
            Sqlcmd.CommandText = builder.ToString
            SQLRd = Sqlcmd.ExecuteReader
            If SQLRd.HasRows Then
                SQLRd.Read()
                strInternalCode = SQLRd("ITEM_CODE").ToString
                strCustDrgno = SQLRd("CUST_ITEM_CODE").ToString
                strDespatchQty = SQLRd("SALES_QUANTITY").ToString
            End If
            If SQLRd.IsClosed = False Then SQLRd.Close()

            builder.Remove(0, builder.ToString.Length)
            builder.AppendLine("UPDATE ITEMBAL_MST SET CUR_BAL=CUR_BAL-" + Val(strDespatchQty).ToString + " WHERE ITEM_CODE='" + strInternalCode + "' AND UNIT_CODE='" + gstrUNITID + "' ")
            Sqlcmd.CommandText = builder.ToString
            Sqlcmd.Parameters.Clear()
            Sqlcmd.ExecuteNonQuery()


            ' ===== NEW
            Sqlcmd.CommandType = CommandType.StoredProcedure
            Sqlcmd.CommandText = "TRADING_UPDATE_DISPATCH_IN_GRIN"
            Sqlcmd.Parameters.Clear()
            Sqlcmd.Parameters.Add("@UNITCODE", SqlDbType.VarChar).Value = gstrUNITID
            Sqlcmd.Parameters.Add("@INVOICE_NO", SqlDbType.VarChar).Value = Val(strInvoiceNo).ToString
            Sqlcmd.ExecuteNonQuery()
            ' ===== NEW


            Sqlcmd.CommandType = CommandType.Text
            builder.Remove(0, builder.ToString.Length)
            builder.AppendLine("UPDATE CUST_ORD_DTL SET DESPATCH_QTY = DESPATCH_QTY + " + Val(strDespatchQty).ToString)
            builder.AppendLine("WHERE UNIT_CODE='" + gstrUNITID + "'  AND ACCOUNT_CODE ='" + Me.txtCustCode.Text.Trim + "' AND CUST_DRGNO = '" + strCustDrgno + "'  AND CUST_REF = '" + Me.txtRefNo.Text.Trim + "'  AND AMENDMENT_NO = '" + Me.txtAmendNo.Text.Trim + "' AND ACTIVE_FLAG ='A'")
            Sqlcmd.CommandText = builder.ToString
            Sqlcmd.Parameters.Clear()
            Sqlcmd.ExecuteNonQuery()


            builder.Remove(0, builder.ToString.Length)
            builder.AppendLine("UPDATE SALESCHALLAN_DTL SET BILL_FLAG=1 WHERE UNIT_CODE='" + gstrUNITID + "'  AND DOC_NO=" + Val(strInvoiceNo).ToString)
            Sqlcmd.CommandText = builder.ToString
            Sqlcmd.Parameters.Clear()
            Sqlcmd.ExecuteNonQuery()

            SqlTrans.Commit()
            IsTrans = False
            LockTradingInvoice = True
        Catch ex As Exception
            If IsTrans = True Then SqlTrans.Rollback()
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            If IsTrans = True Then SqlTrans.Rollback()
            If IsNothing(Sqlcmd.Connection.State) = False Then
                If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
            End If
            If IsNothing(Sqlcmd) = False Then
                Sqlcmd.Connection.Dispose()
                Sqlcmd.Dispose()
            End If
            If IsNothing(SqlTrans) = False Then SqlTrans.Dispose()
        End Try
    End Function
    Public Function CancelTradingInvoice(ByVal strInvoiceNo As String) As Boolean
        Dim Sqlcmd As New SqlCommand
        Dim ObjVal As Object = Nothing
        Dim SqlTrans As SqlTransaction
        Dim IsTrans As Boolean = False
        Dim SQLRd As SqlDataReader
        Dim builder As New StringBuilder

        Dim strDespatchQty As String = String.Empty
        Dim strCustDrgno As String = String.Empty
        Dim strInternalCode As String = String.Empty
        Try
            CancelTradingInvoice = False
            Sqlcmd.CommandTimeout = 0

            Sqlcmd.Connection = SqlConnectionclass.GetConnection
            SqlTrans = Sqlcmd.Connection.BeginTransaction(System.Data.IsolationLevel.Serializable)
            Sqlcmd.Transaction = SqlTrans
            Sqlcmd.CommandType = CommandType.Text


            builder.Remove(0, builder.ToString.Length)
            builder.AppendLine("SELECT ITEM_CODE,CUST_ITEM_CODE,SALES_QUANTITY FROM SALES_DTL WHERE UNIT_CODE='" + gstrUNITID + "'  AND DOC_NO=" + Val(strInvoiceNo).ToString)
            Sqlcmd.CommandText = builder.ToString
            SQLRd = Sqlcmd.ExecuteReader
            If SQLRd.HasRows Then
                SQLRd.Read()
                strInternalCode = SQLRd("ITEM_CODE").ToString
                strCustDrgno = SQLRd("CUST_ITEM_CODE").ToString
                strDespatchQty = SQLRd("SALES_QUANTITY").ToString
            End If
            If SQLRd.IsClosed = False Then SQLRd.Close()


            builder.Remove(0, builder.ToString.Length)
            builder.AppendLine("UPDATE CUST_ORD_DTL SET DESPATCH_QTY = DESPATCH_QTY - " + Val(strDespatchQty).ToString)
            builder.AppendLine("WHERE UNIT_CODE='" + gstrUNITID + "'  AND ACCOUNT_CODE ='" + Me.txtCustCode.Text.Trim + "' AND CUST_DRGNO = '" + strCustDrgno + "'  AND CUST_REF = '" + Me.txtRefNo.Text.Trim + "'  AND AMENDMENT_NO = '" + Me.txtAmendNo.Text.Trim + "' AND ACTIVE_FLAG ='A'")
            Sqlcmd.CommandText = builder.ToString
            Sqlcmd.Parameters.Clear()
            Sqlcmd.ExecuteNonQuery()

            Dim strRetMessage As String
            If UpdateForSchedules("-", Sqlcmd, strInternalCode, strCustDrgno, strDespatchQty, strRetMessage) = False Then
                IsTrans = False
                SqlTrans.Rollback()
                MsgBox(strRetMessage, MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            End If

            Sqlcmd.CommandType = CommandType.Text
            Sqlcmd.Parameters.Clear()

            builder.Remove(0, builder.ToString.Length)
            builder.AppendLine("UPDATE ITEMBAL_MST SET CUR_BAL=CUR_BAL+" + Val(strDespatchQty).ToString + " WHERE ITEM_CODE='" + strInternalCode + "' AND UNIT_CODE='" + gstrUNITID + "' ")
            Sqlcmd.CommandText = builder.ToString
            Sqlcmd.Parameters.Clear()
            Sqlcmd.ExecuteNonQuery()

            builder.Remove(0, builder.ToString.Length)
            builder.AppendLine("UPDATE A SET DESPATCH_QTY_TRADING=ISNULL(DESPATCH_QTY_TRADING,0)-B.KNOCKOFFQTY")
            builder.AppendLine("FROM GRN_DTL A,")
            builder.AppendLine("(	")
            builder.AppendLine("SELECT GRIN_NO,KNOCKOFFQTY,GRIN_DOC_TYPE,ITEM_CODE,UNIT_CODE ")
            builder.AppendLine("FROM SALES_TRADING_GRIN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND DOC_NO=" + Val(strInvoiceNo).ToString)
            builder.AppendLine(") B")
            builder.AppendLine("WHERE A.DOC_NO=B.GRIN_NO")
            builder.AppendLine("AND A.DOC_TYPE=B.GRIN_DOC_TYPE")
            builder.AppendLine("AND A.ITEM_CODE=B.ITEM_CODE  ")
            builder.AppendLine("AND A.UNIT_CODE=B.UNIT_CODE  ")
            builder.AppendLine("AND A.UNIT_CODE='" + gstrUNITID + "'")
            Sqlcmd.CommandText = builder.ToString
            Sqlcmd.Parameters.Clear()
            Sqlcmd.ExecuteNonQuery()

            '=== NEW
            builder.Remove(0, builder.ToString.Length)
            builder.AppendLine("UPDATE SALES_TRADING_GRIN_DTL SET GRIN_KNOCKOFF_SLNO=-1 WHERE UNIT_CODE='" + gstrUNITID + "' AND DOC_NO=" + Val(strInvoiceNo).ToString)
            Sqlcmd.CommandText = builder.ToString
            Sqlcmd.Parameters.Clear()
            Sqlcmd.ExecuteNonQuery()
            '=== NEW

            builder.Remove(0, builder.ToString.Length)
            builder.AppendLine("UPDATE SALESCHALLAN_DTL SET CANCEL_FLAG=1 WHERE UNIT_CODE='" + gstrUNITID + "'  AND DOC_NO=" + Val(strInvoiceNo).ToString)
            Sqlcmd.CommandText = builder.ToString
            Sqlcmd.Parameters.Clear()
            Sqlcmd.ExecuteNonQuery()

            SqlTrans.Commit()
            IsTrans = False
            CancelTradingInvoice = True
        Catch ex As Exception
            If IsTrans = True Then SqlTrans.Rollback()
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            If IsTrans = True Then SqlTrans.Rollback()
            If IsNothing(Sqlcmd.Connection.State) = False Then
                If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
            End If
            If IsNothing(Sqlcmd) = False Then
                Sqlcmd.Connection.Dispose()
                Sqlcmd.Dispose()
            End If
            If IsNothing(SqlTrans) = False Then SqlTrans.Dispose()
        End Try
    End Function
    Private Function DeleteData() As Boolean
        Dim Sqlcmd As New SqlCommand
        Dim strSql As String
        Dim ObjVal As Object = Nothing
        Dim SqlTrans As SqlTransaction
        Dim IsTrans As Boolean = False
        Dim builder As New StringBuilder
        Dim strInvType As String = String.Empty
        Dim strInvSubType As String = String.Empty
        Dim strInvoiceNo As String = "0"
        Dim strDespatchQty As String
        Dim strCustDrgno As String
        Dim strInternalCode As String
        Try
            DeleteData = False
            If blnBillFlag = True Then
                MsgBox("Deletion Not Allowed For Locked Invoice(s).", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            End If
            Sqlcmd.CommandTimeout = 0
            Sqlcmd.Connection = SqlConnectionclass.GetConnection
            Sqlcmd.CommandType = CommandType.Text

            Sqlcmd.CommandText = "SELECT TOP 1 1 FROM SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND DOC_NO>" + Val(Me.txtChallanNo.Text).ToString + " AND ISNULL(BILL_FLAG,0)=0"
            ObjVal = Sqlcmd.ExecuteScalar()
            If IsNothing(ObjVal) = True Then ObjVal = "0"
            If ObjVal = "1" Then
                MsgBox("Please Delete Invoice(s) Made After - " + Me.txtChallanNo.Text, MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            End If
            ObjVal = Nothing

            Sqlcmd.CommandText = "SELECT TOP 1 1 FROM SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND DOC_NO=" + Me.txtChallanNo.Text + " AND ISNULL(BILL_FLAG,0)=1"
            ObjVal = Sqlcmd.ExecuteScalar()
            If IsNothing(ObjVal) = True Then ObjVal = "0"
            If ObjVal = "1" Then
                MsgBox("This Invoice Has Been Locked.Can't Delete.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            End If
            ObjVal = Nothing

            SqlTrans = Sqlcmd.Connection.BeginTransaction(System.Data.IsolationLevel.Serializable)
            Sqlcmd.Transaction = SqlTrans
            Sqlcmd.CommandType = CommandType.Text

            SpChEntry.Col = EnumInv.ENUMQUANTITY
            strDespatchQty = SpChEntry.Text

            SpChEntry.Col = EnumInv.CUSTPARTNO
            strCustDrgno = SpChEntry.Text

            SpChEntry.Col = EnumInv.ENUMITEMCODE
            strInternalCode = SpChEntry.Text


            strInvoiceNo = Me.txtChallanNo.Text
            Sqlcmd.CommandText = "DELETE FROM SALES_TRADING_GRIN_DTL WHERE DOC_NO=" + strInvoiceNo + " AND UNIT_CODE='" + gstrUNITID + "' "
            Sqlcmd.ExecuteNonQuery()

            Sqlcmd.CommandText = "DELETE FROM SALES_DTL				 WHERE DOC_NO=" + strInvoiceNo + "  AND UNIT_CODE='" + gstrUNITID + "' "
            Sqlcmd.ExecuteNonQuery()

            Sqlcmd.CommandText = "DELETE FROM SALESCHALLAN_DTL		 WHERE DOC_NO=" + strInvoiceNo + "  AND UNIT_CODE='" + gstrUNITID + "' "
            Sqlcmd.ExecuteNonQuery()

            'builder.Remove(0, builder.ToString.Length)
            'builder.AppendLine("UPDATE CUST_ORD_DTL SET DESPATCH_QTY = DESPATCH_QTY - " + Val(strDespatchQty).ToString)
            'builder.AppendLine("WHERE UNIT_CODE='" + gstrUNITID + "'  AND ACCOUNT_CODE ='" + Me.txtCustCode.Text.Trim + "' AND CUST_DRGNO = '" + strCustDrgno + "'  AND CUST_REF = '" + Me.txtRefNo.Text.Trim + "'  AND AMENDMENT_NO = '" + Me.txtAmendNo.Text.Trim + "' AND ACTIVE_FLAG ='A'")
            'Sqlcmd.CommandText = builder.ToString
            'Sqlcmd.Parameters.Clear()
            'Sqlcmd.ExecuteNonQuery()

            Dim strRetMessage As String
            If UpdateForSchedules("-", Sqlcmd, strInternalCode, strCustDrgno, strDespatchQty, strRetMessage) = False Then
                IsTrans = False
                SqlTrans.Rollback()
                MsgBox(strRetMessage, MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            End If

            SqlTrans.Commit()
            IsTrans = False
            DeleteData = True
        Catch ex As Exception
            If IsTrans = True Then SqlTrans.Rollback()
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            If IsTrans = True Then SqlTrans.Rollback()
            If IsNothing(Sqlcmd.Connection) = False Then
                If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
                Sqlcmd.Connection.Dispose()
                Sqlcmd.Dispose()
            End If
            If IsNothing(SqlTrans) = False Then SqlTrans.Dispose()
        End Try
    End Function

    Private Function GetAvailableStockAndDespatchQty(ByRef SQLCmd As SqlCommand, ByVal IsStockOrDespatch As Boolean, ByVal strItemCode As String) As Double
        'IsStockOrDespatch=True for Stock , False For Despatch
        Dim ObjVal As Object
        Dim builder As New StringBuilder
        SQLCmd.CommandType = CommandType.Text
        Try
            If IsStockOrDespatch = True Then
                builder.Remove(0, builder.ToString.Length)
                builder.AppendLine("SELECT 'AVAILABLEQTY'=CUR_BAL-")
                builder.AppendLine("ISNULL((")
                builder.AppendLine("SELECT SUM(B.SALES_QUANTITY) FROM SALESCHALLAN_DTL A,SALES_DTL B")
                builder.AppendLine("WHERE A.UNIT_CODE=B.UNIT_CODE")
                builder.AppendLine("AND   A.DOC_NO=B.DOC_NO")
                builder.AppendLine("AND   A.UNIT_CODE='" + gstrUNITID + "'  AND B.ITEM_CODE='" + strItemCode + "'")
                builder.AppendLine("AND   ISNULL(A.BILL_FLAG,0)=0 AND ISNULL(A.CANCEL_FLAG,0)=0")
                builder.AppendLine("),0) ")
                builder.AppendLine("FROM ITEMBAL_MST ")
                builder.AppendLine("WHERE UNIT_CODE='" + gstrUNITID + "'  AND ITEM_CODE='" + strItemCode + "' AND LOCATION_CODE='01T1'")

                SQLCmd.CommandText = builder.ToString
                ObjVal = SQLCmd.ExecuteScalar
                If IsNothing(ObjVal) = True Then ObjVal = "0"
                GetAvailableStockAndDespatchQty = Val(ObjVal.ToString)
            Else
                SQLCmd.CommandText = "SELECT ISNULL(DESPATCH_QTY,0) AS RESTQTY FROM CUST_ORD_DTL WHERE UNIT_CODE='" + gstrUNITID + "'  AND ACCOUNT_CODE='" + Me.txtCustCode.Text.Trim + "' AND CUST_REF='" + Me.txtRefNo.Text.Trim + "' AND AMENDMENT_NO='" + Me.txtAmendNo.Text.Trim + "' AND ITEM_CODE='" + strItemCode + "'"
                ObjVal = SQLCmd.ExecuteScalar
                If IsNothing(ObjVal) = True Then ObjVal = "0"
                GetAvailableStockAndDespatchQty = Val(ObjVal.ToString)
            End If
        Catch EX As Exception
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
        Finally

        End Try
    End Function
    Private Sub GetItemDescription()
        Dim strSql As String
        Dim Sqlcmd As New SqlCommand
        Dim SQLCon As SqlConnection
        Dim ObjVal As Object

        SQLCon = SqlConnectionclass.GetConnection()
        Sqlcmd.Connection = SQLCon
        Sqlcmd.CommandType = CommandType.Text
        Try


            Sqlcmd.CommandText = "SELECT UNT_UNITNAME FROM GEN_UNITMASTER WHERE UNT_CODEID='" + Me.txtLocationCode.Text + "'"
            ObjVal = Sqlcmd.ExecuteScalar()
            If IsNothing(ObjVal) = True Then ObjVal = String.Empty
            Me.lblLocCodeDes.Text = ObjVal.ToString


            Me.SpChEntry.Row = Me.SpChEntry.MaxRows
            Me.SpChEntry.Col = EnumInv.ENUMITEMCODE
            Sqlcmd.CommandText = "SELECT DESCRIPTION FROM ITEM_MST WHERE ITEM_CODE='" + Me.SpChEntry.Text.Trim + "' AND UNIT_CODE='" + gstrUNITID + "' "
            ObjVal = Sqlcmd.ExecuteScalar()
            If IsNothing(ObjVal) = True Then ObjVal = String.Empty
            Me.lblInternalPartDesc.Text = ObjVal.ToString

            Sqlcmd.CommandText = "SELECT DRG_DESC FROM CUSTITEM_MST WHERE UNIT_CODE='" + gstrUNITID + "'  AND ACCOUNT_CODE='" + Me.txtCustCode.Text + "' AND ITEM_CODE='" + SpChEntry.Text + "'"
            ObjVal = Sqlcmd.ExecuteScalar()
            If IsNothing(ObjVal) = True Then ObjVal = String.Empty
            Me.lblCustPartDesc.Text = ObjVal.ToString
            Me.lblCurrentStock.Text = GetAvailableStockAndDespatchQty(Sqlcmd, True, Me.SpChEntry.Text.Trim)
            Me.lblDespetchQty.Text = GetAvailableStockAndDespatchQty(Sqlcmd, False, Me.SpChEntry.Text.Trim)
        Catch EX As Exception
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
        Finally
            If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
            If SQLCon.State = ConnectionState.Open Then SQLCon.Close()
            Sqlcmd.Connection.Dispose()
            Sqlcmd.Dispose()
            SQLCon.Dispose()
        End Try
    End Sub
    Private Function UpdateForSchedules(ByVal pstrUpdType As String, ByRef SQLCmd As SqlCommand, ByVal Item_Code As String, ByVal CustDrawingNo As String, ByVal Quantity As String, ByRef ReturnMessage As String) As Boolean
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim intCtr As Integer
        Dim strMSG As String
        Dim strYYYYmm As String
        Dim curQty As Decimal
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim varItemQty As Object

        UpdateForSchedules = True
        strYYYYmm = Me.dtpDateDesc.Value.ToString("yyyy")  'Year(ConvertToDate(lblDateDes.Text)) & VB.Right("0" & Month(ConvertToDate(lblDateDes.Text)), 2)
        With SpChEntry

            varItemCode = Item_Code
            varDrgNo = CustDrawingNo
            varItemQty = Quantity

            SQLCmd.CommandType = CommandType.StoredProcedure
            SQLCmd.CommandText = "MKT_SCHEDULE_KNOCKOFF_NORTH"
            SQLCmd.Parameters.Add("@UNITCODE", SqlDbType.VarChar).Value = gstrUNITID
            SQLCmd.Parameters.Add("@CUSTOMER_CODE", SqlDbType.VarChar).Value = Trim(txtCustCode.Text)
            SQLCmd.Parameters.Add("@ITEM_CODE", SqlDbType.VarChar).Value = varItemCode
            SQLCmd.Parameters.Add("@CUSTDRG_NO", SqlDbType.VarChar).Value = varDrgNo
            SQLCmd.Parameters.Add("@FLAG", SqlDbType.VarChar).Value = pstrUpdType
            SQLCmd.Parameters.Add("@SCH_TYPE", SqlDbType.VarChar).Value = "D"
            SQLCmd.Parameters.Add("@YYYYMM", SqlDbType.VarChar).Value = strYYYYmm
            SQLCmd.Parameters.Add("@REQ_QTY", SqlDbType.Decimal).Value = Val(varItemQty)
            SQLCmd.Parameters.Add("@DATE", SqlDbType.VarChar).Value = getDateForDB(Me.dtpDateDesc.Value.ToString("dd/MM/yyyy"))
            SQLCmd.Parameters.Add("@MSG", SqlDbType.VarChar).Value = String.Empty
            SQLCmd.Parameters.Add("@ERR", SqlDbType.VarChar).Value = String.Empty
            SQLCmd.Parameters(9).Direction = ParameterDirection.Output
            SQLCmd.Parameters(10).Direction = ParameterDirection.Output
            SQLCmd.ExecuteNonQuery()
            If Len(SQLCmd.Parameters(9).Value) > 0 Then
                ReturnMessage = SQLCmd.Parameters(9).Value
                UpdateForSchedules = False
                Exit Function
            End If
            If Len(SQLCmd.Parameters(10).Value) > 0 Then
                ReturnMessage = SQLCmd.Parameters(10).Value
                UpdateForSchedules = False
                Exit Function
            End If
        End With
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function SaveData() As Boolean
        Dim Sqlcmd As New SqlCommand
        Dim ObjVal As Object = Nothing
        Dim SqlTrans As SqlTransaction
        Dim IsTrans As Boolean = False
        Dim DataRd As SqlDataReader
        Dim builder As New StringBuilder
        Dim strInvType As String = String.Empty
        Dim strInvSubType As String = String.Empty
        Dim strInvoiceNo As String = "0"
        Dim strDespatchQty As String
        Dim strCustDrgno As String
        Dim strInternalCode As String
        Dim strToLocation As String

        Dim blnISInsExcisable As Boolean
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
        Dim intUNLOCKED_INVOICES As Integer
        Try
            SaveData = False
            If blnBillFlag = True Then
                MsgBox("Modifications Not Allowed In Locked Invoice(s).", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            End If

            If Me.txtCustCode.Text.Trim = String.Empty Then
                MsgBox("Please Select The Customer.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            ElseIf Me.txtRefNo.Text.Trim = String.Empty Then
                MsgBox("Please Select The Refrence Number.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            ElseIf Me.txtCreditTerms.Text.Trim = String.Empty Then
                MsgBox("Please Select Credit Terms.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            ElseIf Me.txtECSSTaxType.Text.Trim = String.Empty Then
                MsgBox("Please Select E Cess Type.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            ElseIf Me.txtSECSSTaxType.Text.Trim = String.Empty Then
                MsgBox("Please Select H Cess Type.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            ElseIf Me.txtSaleTaxType.Text.Trim = String.Empty Then
                MsgBox("Please Select Sale Tax Type.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
                'ElseIf Me.txtAddVAT.Text.Trim = String.Empty Then
                '   MsgBox("Please Select Add. VAT Type.", MsgBoxStyle.Information, ResolveResString(100))
                '  Exit Function
            ElseIf Me.SpChEntry.MaxRows = 0 Then
                MsgBox("Please Select Item To Be Despatched.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            ElseIf Val(LblNetInvoiceValue.Text) = 0 Then
                MsgBox("Please Check The Invoice It's Value Can't Be Zero.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            End If

            SpChEntry.Row = 1
            SpChEntry.Col = EnumInv.ENUMQUANTITY
            If Val(SpChEntry.Text) = 0 Then
                MsgBox("Item Quantity In Invoice Can't Be Zero.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            End If

            Sqlcmd.CommandTimeout = 0
            Sqlcmd.Connection = SqlConnectionclass.GetConnection
            Sqlcmd.CommandType = CommandType.Text

            Sqlcmd.CommandText = "SELECT 'UNLOCKED_INVOICES'=(SELECT COUNT(*) FROM SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND BILL_FLAG=0),NOOFTRADINGINVOICEWITHOUTLOCKING,InsExc_Excise,CustSupp_Inc,EOU_Flag, Basic_Roundoff, Basic_Roundoff_decimal, SalesTax_Roundoff, SalesTax_Roundoff_decimal, Excise_Roundoff, Excise_Roundoff_decimal, "
            Sqlcmd.CommandText = Sqlcmd.CommandText + " SST_Roundoff, SST_Roundoff_decimal, InsInc_SalesTax, TCSTax_Roundoff, TCSTax_Roundoff_decimal, TotalToolCostRoundoff, TotalToolCostRoundoff_Decimal, ECESS_Roundoff, ECESSRoundoff_Decimal, ECESSOnSaleTax_Roundoff, ECESSOnSaleTaxRoundOff_Decimal, "
            Sqlcmd.CommandText = Sqlcmd.CommandText + " TurnOverTax_RoundOff, TurnOverTaxRoundOff_Decimal, TotalInvoiceAmount_RoundOff,TotalInvoiceAmountRoundOff_Decimal, SDTRoundOff, SDTRoundOff_Decimal,SameUnitLoading,ServiceTax_Roundoff,ServiceTaxRoundoff_Decimal=isnull(ServiceTaxRoundoff_Decimal,0),Packing_Roundoff,PackingRoundoff_Decimal=isnull(PackingRoundoff_Decimal,0) FROM Sales_Parameter WHERE UNIT_CODE='" + gstrUNITID + "'"

            strToLocation = ReturnCustomerLocation()
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
                intNOOFTRADINGINVOICEWITHOUTLOCKING = Val(DataRd("NOOFTRADINGINVOICEWITHOUTLOCKING").ToString)
                intUNLOCKED_INVOICES = Val(DataRd("UNLOCKED_INVOICES").ToString)
            Else
                MsgBox("No Data Define In Sales_Parameter Table", MsgBoxStyle.Critical, ResolveResString(100))
                Exit Function
            End If
            If DataRd.IsClosed = False Then DataRd.Close()


            If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If intUNLOCKED_INVOICES >= intNOOFTRADINGINVOICEWITHOUTLOCKING Then
                    MsgBox("Only " + intNOOFTRADINGINVOICEWITHOUTLOCKING.ToString + " Trading Invoice(s) Can Be Made Without Locking Previous Invoice(s).", MsgBoxStyle.Critical, ResolveResString(100))
                    Exit Function
                End If
            End If

            Sqlcmd.CommandText = "SELECT INVOICE_TYPE,SUB_TYPE FROM SALECONF WHERE UNIT_CODE='" + gstrUNITID + "' AND  DESCRIPTION ='" & Trim(CmbInvType.Text) & "'AND SUB_TYPE_DESCRIPTION ='" & Trim(CmbInvSubType.Text) & "' AND (FIN_START_DATE <= GETDATE() AND FIN_END_DATE >= GETDATE())"
            DataRd = Sqlcmd.ExecuteReader()
            If DataRd.HasRows Then
                DataRd.Read()
                strInvType = DataRd("INVOICE_TYPE")
                strInvSubType = DataRd("SUB_TYPE")
            End If
            If DataRd.IsClosed = False Then DataRd.Close()

            SqlTrans = Sqlcmd.Connection.BeginTransaction(System.Data.IsolationLevel.Serializable)
            Sqlcmd.Transaction = SqlTrans
            Sqlcmd.CommandType = CommandType.Text
            IsTrans = True

            '===========================================================================
            If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then

                Dim strAccountCode_Edit As String = String.Empty
                Dim strCust_Ref_Edit As String = String.Empty
                Dim strSalesQuantity_Edit As String = String.Empty
                Dim strAmmendment_Edit As String = String.Empty
                Dim strInternalItem_Edit As String = String.Empty
                Dim strCustItem_Edit As String = String.Empty

                strInvoiceNo = Me.txtChallanNo.Text
                Sqlcmd.CommandText = "SELECT TOP 1 1 FROM SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND DOC_NO>" + Val(Me.txtChallanNo.Text).ToString + " AND ISNULL(BILL_FLAG,0)=0"
                ObjVal = Sqlcmd.ExecuteScalar()
                If IsNothing(ObjVal) = True Then ObjVal = "0"
                If ObjVal = "1" Then
                    SqlTrans.Rollback()
                    IsTrans = False
                    MsgBox("Befor Editing, Please Delete Invoice(s) Made After - " + Me.txtChallanNo.Text, MsgBoxStyle.Information, ResolveResString(100))
                    Exit Function
                End If
                ObjVal = Nothing

                Sqlcmd.CommandText = " SELECT A.ACCOUNT_CODE,A.CUST_REF,B.SALES_QUANTITY,A.AMENDMENT_NO,B.ITEM_CODE,B.CUST_ITEM_CODE FROM SALESCHALLAN_DTL A,SALES_DTL B"
                Sqlcmd.CommandText = Sqlcmd.CommandText + " WHERE	A.DOC_NO=B.DOC_NO"
                Sqlcmd.CommandText = Sqlcmd.CommandText + " AND		A.UNIT_CODE=B.UNIT_CODE"
                Sqlcmd.CommandText = Sqlcmd.CommandText + " AND		A.UNIT_CODE='" + gstrUNITID + "'  AND A.DOC_NO=" + Val(strInvoiceNo).ToString
                DataRd = Sqlcmd.ExecuteReader()
                If DataRd.HasRows Then
                    DataRd.Read()
                    strAccountCode_Edit = DataRd("ACCOUNT_CODE")
                    strCust_Ref_Edit = DataRd("CUST_REF")
                    strSalesQuantity_Edit = DataRd("SALES_QUANTITY")
                    strAmmendment_Edit = DataRd("AMENDMENT_NO")
                    strCustItem_Edit = DataRd("CUST_ITEM_CODE")
                    strInternalItem_Edit = DataRd("ITEM_CODE")
                End If
                If DataRd.IsClosed = False Then DataRd.Close()


                Sqlcmd.CommandText = "DELETE FROM SALES_TRADING_GRIN_DTL WHERE DOC_NO=" + strInvoiceNo + " AND UNIT_CODE='" + gstrUNITID + "' "
                Sqlcmd.ExecuteNonQuery()

                Sqlcmd.CommandText = "DELETE FROM SALES_DTL				 WHERE DOC_NO=" + strInvoiceNo + "  AND UNIT_CODE='" + gstrUNITID + "' "
                Sqlcmd.ExecuteNonQuery()

                Sqlcmd.CommandText = "DELETE FROM SALESCHALLAN_DTL		 WHERE DOC_NO=" + strInvoiceNo + "  AND UNIT_CODE='" + gstrUNITID + "' "
                Sqlcmd.ExecuteNonQuery()

                'builder.Remove(0, builder.ToString.Length)
                'builder.AppendLine("UPDATE CUST_ORD_DTL SET DESPATCH_QTY = DESPATCH_QTY - " + Val(strSalesQuantity_Edit).ToString)
                'builder.AppendLine("WHERE UNIT_CODE='" + gstrUNITID + "'  AND ACCOUNT_CODE ='" + strAccountCode_Edit + "' AND CUST_DRGNO = '" + strCustItem_Edit + "'  AND CUST_REF = '" + strCust_Ref_Edit + "'  AND RTRIM(ISNULL(AMENDMENT_NO,'')) = '" + strAmmendment_Edit + "' AND ACTIVE_FLAG ='A'")
                'Sqlcmd.CommandText = builder.ToString
                'Sqlcmd.Parameters.Clear()
                'Sqlcmd.ExecuteNonQuery()

                Dim strReturnMessage As String
                If UpdateForSchedules("-", Sqlcmd, strInternalItem_Edit, strCustItem_Edit, strSalesQuantity_Edit, strReturnMessage) = False Then
                    IsTrans = False
                    SqlTrans.Rollback()
                    MsgBox(strReturnMessage, MsgBoxStyle.Information, ResolveResString(100))
                    Exit Function
                End If
            End If
            '===========================================================================

            SpChEntry.Col = EnumInv.ENUMQUANTITY
            strDespatchQty = SpChEntry.Text

            SpChEntry.Col = EnumInv.CUSTPARTNO
            strCustDrgno = SpChEntry.Text

            SpChEntry.Col = EnumInv.ENUMITEMCODE
            strInternalCode = SpChEntry.Text


            'CHECKING SALES ORDER QUANTITY AND SALES SCHEDULE TO VALIDATE THE ITEM FOR INVOICING
            Sqlcmd.CommandType = CommandType.Text
            Sqlcmd.Parameters.Clear()
            Sqlcmd.CommandText = "SELECT ORDER_QTY-ISNULL(DESPATCH_QTY,0)"
            Sqlcmd.CommandText = Sqlcmd.CommandText + " - ISNULL(("
            Sqlcmd.CommandText = Sqlcmd.CommandText + " SELECT SUM(B.SALES_QUANTITY) UNLOCKED_QTY FROM SALESCHALLAN_DTL A, SALES_DTL B"
            Sqlcmd.CommandText = Sqlcmd.CommandText + " WHERE  A.DOC_NO=B.DOC_NO AND A.UNIT_CODE=B.UNIT_CODE AND ISNULL(A.BILL_FLAG,0)=0 "
            Sqlcmd.CommandText = Sqlcmd.CommandText + " AND	   ISNULL(CANCEL_FLAG,0)=0 AND	   A.CUST_REF=X.CUST_REF"
            Sqlcmd.CommandText = Sqlcmd.CommandText + " AND	   RTRIM(ISNULL(A.AMENDMENT_NO,''))=RTRIM(ISNULL(X.AMENDMENT_NO,'')) AND A.UNIT_CODE='" + gstrUNITID + "'"
            Sqlcmd.CommandText = Sqlcmd.CommandText + " ),0) "
            Sqlcmd.CommandText = Sqlcmd.CommandText + " AS RESTQTY FROM CUST_ORD_DTL X WHERE X.UNIT_CODE='" + gstrUNITID + "'  AND X.ACCOUNT_CODE='" + Me.txtCustCode.Text.Trim + "' AND X.CUST_REF='" + Me.txtRefNo.Text.Trim + "' AND RTRIM(ISNULL(X.AMENDMENT_NO,''))='" + Me.txtAmendNo.Text.Trim + "' AND X.ITEM_CODE='" + strInternalCode + "'"

            DataRd = Sqlcmd.ExecuteReader()
            Dim strRestQty As String
            If DataRd.HasRows Then
                DataRd.Read()
                strRestQty = DataRd("RESTQTY").ToString
                If Val(strDespatchQty) > Val(strRestQty) Then
                    If DataRd.IsClosed = False Then DataRd.Close()
                    SqlTrans.Rollback()
                    IsTrans = False
                    MsgBox("Can't Despatch More Than Sales Order Quantity. Remaining Quantity of Selected Item is " + strRestQty + ".", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Function
                End If
            End If
            If DataRd.IsClosed = False Then DataRd.Close()


            Sqlcmd.CommandText = "SELECT SUM(SCHEDULE_QUANTITY)-SUM(DESPATCH_QTY) RESTSCHQTY  FROM DAILYMKTSCHEDULE WHERE UNIT_CODE='" + gstrUNITID + "'  AND ACCOUNT_CODE='" + Me.txtCustCode.Text.Trim + "' AND ITEM_CODE='" + strInternalCode + "' AND TRANS_DATE<='" + Me.dtpDateDesc.Value.ToString("dd MMM yyyy") + "' AND STATUS=1"
            DataRd = Sqlcmd.ExecuteReader()
            If DataRd.HasRows Then
                DataRd.Read()
                strRestQty = DataRd("RESTSCHQTY").ToString
                If Val(strDespatchQty) > Val(strRestQty) Then
                    If DataRd.IsClosed = False Then DataRd.Close()
                    SqlTrans.Rollback()
                    IsTrans = False
                    MsgBox("Can't Despatch More Than Sales Schedule. Remaining Sales Schedule Quantity of Selected Item is " + strRestQty + ".", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Function
                End If
            End If
            If DataRd.IsClosed = False Then DataRd.Close()



            builder.Remove(0, builder.ToString.Length)

            builder.AppendLine("SELECT 'AVAILABLEQTY'=CUR_BAL-")
            builder.AppendLine("ISNULL((")
            builder.AppendLine("SELECT SUM(B.SALES_QUANTITY) FROM SALESCHALLAN_DTL A,SALES_DTL B")
            builder.AppendLine("WHERE A.UNIT_CODE=B.UNIT_CODE")
            builder.AppendLine("AND   A.DOC_NO=B.DOC_NO")
            builder.AppendLine("AND   A.UNIT_CODE='" + gstrUNITID + "'  AND B.ITEM_CODE='" + strInternalCode + "'")
            builder.AppendLine("AND ISNULL(A.BILL_FLAG,0)=0 AND ISNULL(A.CANCEL_FLAG,0)=0")
            builder.AppendLine("),0) ")
            builder.AppendLine("FROM ITEMBAL_MST ")
            builder.AppendLine("WHERE UNIT_CODE='" + gstrUNITID + "'  AND ITEM_CODE='" + strInternalCode + "' AND LOCATION_CODE='01T1'")

            Sqlcmd.CommandText = builder.ToString
            DataRd = Sqlcmd.ExecuteReader()
            If DataRd.HasRows Then
                DataRd.Read()
                strRestQty = DataRd("AVAILABLEQTY").ToString
                If Val(strDespatchQty) > Val(strRestQty) Then
                    If DataRd.IsClosed = False Then DataRd.Close()
                    SqlTrans.Rollback()
                    IsTrans = False
                    MsgBox("Can't Despatch More Than Available Stock. Available Stock of Selected Item is " + strRestQty + ".", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Function
                End If
            End If
            If DataRd.IsClosed = False Then DataRd.Close()


            'CHECKING SALES ORDER QUANTITY AND SALES SCHEDULE TO VALIDATE THE ITEM FOR INVOICING

            If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                Sqlcmd.CommandText = "SELECT (CURRENT_NO + 1)CURRENT_NO FROM DOCUMENTTYPE_MST WHERE UNIT_CODE='" + gstrUNITID + "' AND  DOC_TYPE = 9999 AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE"
                DataRd = Sqlcmd.ExecuteReader()
                If DataRd.HasRows Then
                    DataRd.Read()
                    strInvoiceNo = DataRd("CURRENT_NO").ToString
                    While Len(strInvoiceNo) < 6
                        strInvoiceNo = "0" + strInvoiceNo
                    End While
                    strInvoiceNo = "99" + strInvoiceNo
                Else
                    If DataRd.IsClosed = False Then DataRd.Close()
                    SqlTrans.Rollback()
                    IsTrans = False
                    MsgBox("Temporary Invoice No. Series Not Define. Invoice Entry Can Not Be Saved.", MsgBoxStyle.Information, ResolveResString(100))
                    txtChallanNo.Text = String.Empty
                    Exit Function
                End If
                If DataRd.IsClosed = False Then DataRd.Close()

                Sqlcmd.CommandText = "UPDATE DOCUMENTTYPE_MST SET CURRENT_NO = CURRENT_NO + 1 WHERE UNIT_CODE='" + gstrUNITID + "' AND DOC_TYPE = 9999 AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE"
                Sqlcmd.ExecuteNonQuery()
            End If



            builder.Remove(0, builder.ToString.Length)
            builder.AppendLine("INSERT INTO SALESCHALLAN_DTL ")
            builder.AppendLine("(")
            builder.AppendLine("Account_Code,")
            builder.AppendLine("CONSIGNEE_CODE,")
            builder.AppendLine("Bill_Flag,")
            builder.AppendLine("FIFO_flag,")
            builder.AppendLine("Print_Flag,")
            builder.AppendLine("Cancel_flag,")
            builder.AppendLine("dataPosted,")
            builder.AppendLine("ftp,")
            builder.AppendLine("ExciseExumpted,")
            builder.AppendLine("ServiceInvoiceformatExport,")
            builder.AppendLine("Discount_Type,")
            builder.AppendLine("RejectionPosting,")
            builder.AppendLine("FOC_Invoice,")
            builder.AppendLine("PrintExciseFormat,")
            builder.AppendLine("FreshCrRecd,")
            builder.AppendLine("Trans_Parameter_Flag,")
            builder.AppendLine("InvoiceAgainstMultipleSO,")
            builder.AppendLine("TextFileGenerated,")
            builder.AppendLine("sameunitloading,")
            builder.AppendLine("postingFlag,")
            builder.AppendLine("MULTIPLESO,")
            builder.AppendLine("Suffix,")
            builder.AppendLine("Transport_Type,")
            builder.AppendLine("Invoice_Type,")
            builder.AppendLine("Sub_Category,")
            builder.AppendLine("SalesTax_Type,")
            builder.AppendLine("SalesTax_FormNo,")
            builder.AppendLine("other_ref,")
            builder.AppendLine("Surcharge_salesTaxType,")
            builder.AppendLine("LoadingChargeTaxType,")
            builder.AppendLine("CustBankID,")
            builder.AppendLine("USLOC,")
            builder.AppendLine("TCSTax_Type,")
            builder.AppendLine("ECESS_Type,")
            builder.AppendLine("SRCESS_Type,")
            builder.AppendLine("CVDCESS_Type,")
            builder.AppendLine("TurnOverTaxType,")
            builder.AppendLine("SDTax_Type,")
            builder.AppendLine("ServiceTax_Type,")
            builder.AppendLine("SECESS_Type,")
            builder.AppendLine("CVDSECESS_Type,")
            builder.AppendLine("SRSECESS_Type,")
            builder.AppendLine("ISCHALLAN,")
            builder.AppendLine("ISCONSOLIDATE,")
            builder.AppendLine("Ecess_TotalDuty_Type,")
            builder.AppendLine("SEcess_TotalDuty_Type,")
            builder.AppendLine("ADDVAT_Type,")
            builder.AppendLine("Invoice_Date,")
            builder.AppendLine("Form3Date,")
            builder.AppendLine("Exchange_Date,")
            builder.AppendLine("Ent_dt,")
            builder.AppendLine("Upd_dt,")
            builder.AppendLine("Doc_No,")
            builder.AppendLine("NRGPNOIncaseOfServiceInvoice,")
            builder.AppendLine("Discount_Per,")
            builder.AppendLine("Year,")
            builder.AppendLine("Annex_no,")
            builder.AppendLine("pervalue,")
            builder.AppendLine("dataposted_fin,")
            builder.AppendLine("Location_Code,")
            builder.AppendLine("To_Location,")
            builder.AppendLine("From_Location,")
            builder.AppendLine("Insurance,")
            builder.AppendLine("Frieght_Tax,")
            builder.AppendLine("Sales_Tax_Amount,")
            builder.AppendLine("Surcharge_Sales_Tax_Amount,")
            builder.AppendLine("Frieght_Amount,")
            builder.AppendLine("Packing_Amount,")
            builder.AppendLine("SalesTax_FormValue,")
            builder.AppendLine("total_amount,")
            builder.AppendLine("TurnoverTax_per,")
            builder.AppendLine("Turnover_amt,")
            builder.AppendLine("LoadingChargeTaxAmount,")
            builder.AppendLine("Discount_Amount,")
            builder.AppendLine("TCSTaxAmount,")
            builder.AppendLine("ECESS_Amount,")
            builder.AppendLine("SRCESS_Amount,")
            builder.AppendLine("CVDCESS_Amount,")
            builder.AppendLine("TotalInvoiceAmtRoundOff_diff,")
            builder.AppendLine("SDTax_Amount,")
            builder.AppendLine("ServiceTax_Amount,")
            builder.AppendLine("Prev_Yr_ExportSales,")
            builder.AppendLine("Permissible_Limit_SmpExport,")
            builder.AppendLine("SECESS_Amount,")
            builder.AppendLine("CVDSECESS_Amount,")
            builder.AppendLine("SRSECESS_Amount,")
            builder.AppendLine("Tot_Add_Excise_Amt,")
            builder.AppendLine("Ecess_TotalDuty_Amount,")
            builder.AppendLine("SEcess_TotalDuty_Amount,")
            builder.AppendLine("ADDVAT_Amount,")
            builder.AppendLine("Exchange_Rate,")
            builder.AppendLine("SalesTax_Per,")
            builder.AppendLine("Surcharge_SalesTax_Per,")
            builder.AppendLine("LoadingChargeTax_Per,")
            builder.AppendLine("TCSTax_Per,")
            builder.AppendLine("ECESS_Per,")
            builder.AppendLine("SRCESS_Per,")
            builder.AppendLine("CVDCESS_Per,")
            builder.AppendLine("Excise_Percentage,")
            builder.AppendLine("Permissible_Limit,")
            builder.AppendLine("SDTax_Per,")
            builder.AppendLine("ServiceTax_Per,")
            builder.AppendLine("SECESS_Per,")
            builder.AppendLine("CVDSECESS_Per,")
            builder.AppendLine("SRSECESS_Per,")
            builder.AppendLine("Tot_Add_Excise_PER,")
            builder.AppendLine("bond17OpeningBal,")
            builder.AppendLine("Ecess_TotalDuty_Per,")
            builder.AppendLine("SEcess_TotalDuty_Per,")
            builder.AppendLine("ADDVAT_Per,")
            builder.AppendLine("BasicExciseAndCessValue,")
            builder.AppendLine("total_quantity,")
            builder.AppendLine("Ent_UserId,")
            builder.AppendLine("Upd_Userid,")
            builder.AppendLine("Vehicle_No,")
            builder.AppendLine("From_Station,")
            builder.AppendLine("To_Station,")
            builder.AppendLine("Cust_Ref,")
            builder.AppendLine("Amendment_No,")
            builder.AppendLine("Print_DateTime,")
            builder.AppendLine("Form3,")
            builder.AppendLine("Carriage_Name,")
            builder.AppendLine("Ref_Doc_No,")
            builder.AppendLine("Cust_Name,")
            builder.AppendLine("Currency_Code,")
            builder.AppendLine("Nature_of_Contract,")
            builder.AppendLine("OriginStatus,")
            builder.AppendLine("Ctry_Destination_Goods,")
            builder.AppendLine("Delivery_Terms,")
            builder.AppendLine("Payment_Terms,")
            builder.AppendLine("Pre_Carriage_By,")
            builder.AppendLine("Receipt_Precarriage_at,")
            builder.AppendLine("Vessel_Flight_number,")
            builder.AppendLine("Port_Of_Loading,")
            builder.AppendLine("Port_Of_Discharge,")
            builder.AppendLine("Final_destination,")
            builder.AppendLine("Mode_Of_Shipment,")
            builder.AppendLine("Dispatch_mode,")
            builder.AppendLine("Buyer_description_Of_Goods,")
            builder.AppendLine("Invoice_description_of_EPC,")
            builder.AppendLine("Buyer_Id,")
            builder.AppendLine("remarks,")
            builder.AppendLine("Excise_Type,")
            builder.AppendLine("SRVDINO,")
            builder.AppendLine("SRVLocation,")
            builder.AppendLine("ConsigneeContactPerson,")
            builder.AppendLine("ConsigneeAddress1,")
            builder.AppendLine("ConsigneeAddress2,")
            builder.AppendLine("ConsigneeAddress3,")
            builder.AppendLine("ConsigneeECCNo,")
            builder.AppendLine("ConsigneeLST,")
            builder.AppendLine("SchTime,")
            builder.AppendLine("invoice_time,")
            builder.AppendLine("varGeneralRemarks,")
            builder.AppendLine("CheckSheetNo,")
            builder.AppendLine("Lorry_No,")
            builder.AppendLine("OTL_No,")
            builder.AppendLine("RefChallan,")
            builder.AppendLine("price_bases,")
            builder.AppendLine("LorryNo_date,")
            builder.AppendLine("ConsInvString,")
            builder.AppendLine("invoicepicking_status,")
            builder.AppendLine("TMP_DOC_NO,")
            builder.AppendLine("UNIT_CODE")




            builder.AppendLine(")")

            builder.AppendLine(" SELECT ") 'INSERT STATMENT

            builder.AppendLine("@Account_Code,")
            builder.AppendLine("@CONSIGNEE_CODE,")
            builder.AppendLine("@Bill_Flag,")
            builder.AppendLine("@FIFO_flag,")
            builder.AppendLine("@Print_Flag,")
            builder.AppendLine("@Cancel_flag,")
            builder.AppendLine("@dataPosted,")
            builder.AppendLine("@ftp,")
            builder.AppendLine("@ExciseExumpted,")
            builder.AppendLine("@ServiceInvoiceformatExport,")
            builder.AppendLine("@Discount_Type,")
            builder.AppendLine("@RejectionPosting,")
            builder.AppendLine("@FOC_Invoice,")
            builder.AppendLine("@PrintExciseFormat,")
            builder.AppendLine("@FreshCrRecd,")
            builder.AppendLine("@Trans_Parameter_Flag,")
            builder.AppendLine("@InvoiceAgainstMultipleSO,")
            builder.AppendLine("@TextFileGenerated,")
            builder.AppendLine("@sameunitloading,")
            builder.AppendLine("@postingFlag,")
            builder.AppendLine("@MULTIPLESO,")
            builder.AppendLine("@Suffix,")
            builder.AppendLine("@Transport_Type,")
            builder.AppendLine("@Invoice_Type,")
            builder.AppendLine("@Sub_Category,")
            builder.AppendLine("@SalesTax_Type,")
            builder.AppendLine("@SalesTax_FormNo,")
            builder.AppendLine("@other_ref,")
            builder.AppendLine("@Surcharge_salesTaxType,")
            builder.AppendLine("@LoadingChargeTaxType,")
            builder.AppendLine("@CustBankID,")
            builder.AppendLine("@USLOC,")
            builder.AppendLine("@TCSTax_Type,")
            builder.AppendLine("@ECESS_Type,")
            builder.AppendLine("@SRCESS_Type,")
            builder.AppendLine("@CVDCESS_Type,")
            builder.AppendLine("@TurnOverTaxType,")
            builder.AppendLine("@SDTax_Type,")
            builder.AppendLine("@ServiceTax_Type,")
            builder.AppendLine("@SECESS_Type,")
            builder.AppendLine("@CVDSECESS_Type,")
            builder.AppendLine("@SRSECESS_Type,")
            builder.AppendLine("@ISCHALLAN,")
            builder.AppendLine("@ISCONSOLIDATE,")
            builder.AppendLine("@Ecess_TotalDuty_Type,")
            builder.AppendLine("@SEcess_TotalDuty_Type,")
            builder.AppendLine("@ADDVAT_Type,")
            builder.AppendLine("@Invoice_Date,")
            builder.AppendLine("@Form3Date,")
            builder.AppendLine("@Exchange_Date,")
            builder.AppendLine("@Ent_dt,")
            builder.AppendLine("@Upd_dt,")
            builder.AppendLine("@Doc_No,")
            builder.AppendLine("@NRGPNOIncaseOfServiceInvoice,")
            builder.AppendLine("@Discount_Per,")
            builder.AppendLine("@Year,")
            builder.AppendLine("@Annex_no,")
            builder.AppendLine("@pervalue,")
            builder.AppendLine("@dataposted_fin,")
            builder.AppendLine("@Location_Code,")
            builder.AppendLine("@To_Location,")
            builder.AppendLine("@From_Location,")
            builder.AppendLine("@Insurance,")
            builder.AppendLine("@Frieght_Tax,")
            builder.AppendLine("@Sales_Tax_Amount,")
            builder.AppendLine("@Surcharge_Sales_Tax_Amount,")
            builder.AppendLine("@Frieght_Amount,")
            builder.AppendLine("@Packing_Amount,")
            builder.AppendLine("@SalesTax_FormValue,")
            builder.AppendLine("@total_amount,")
            builder.AppendLine("@TurnoverTax_per,")
            builder.AppendLine("@Turnover_amt,")
            builder.AppendLine("@LoadingChargeTaxAmount,")
            builder.AppendLine("@Discount_Amount,")
            builder.AppendLine("@TCSTaxAmount,")
            builder.AppendLine("@ECESS_Amount,")
            builder.AppendLine("@SRCESS_Amount,")
            builder.AppendLine("@CVDCESS_Amount,")
            builder.AppendLine("@TotalInvoiceAmtRoundOff_diff,")
            builder.AppendLine("@SDTax_Amount,")
            builder.AppendLine("@ServiceTax_Amount,")
            builder.AppendLine("@Prev_Yr_ExportSales,")
            builder.AppendLine("@Permissible_Limit_SmpExport,")
            builder.AppendLine("@SECESS_Amount,")
            builder.AppendLine("@CVDSECESS_Amount,")
            builder.AppendLine("@SRSECESS_Amount,")
            builder.AppendLine("@Tot_Add_Excise_Amt,")
            builder.AppendLine("@Ecess_TotalDuty_Amount,")
            builder.AppendLine("@SEcess_TotalDuty_Amount,")
            builder.AppendLine("@ADDVAT_Amount,")
            builder.AppendLine("@Exchange_Rate,")
            builder.AppendLine("@SalesTax_Per,")
            builder.AppendLine("@Surcharge_SalesTax_Per,")
            builder.AppendLine("@LoadingChargeTax_Per,")
            builder.AppendLine("@TCSTax_Per,")
            builder.AppendLine("@ECESS_Per,")
            builder.AppendLine("@SRCESS_Per,")
            builder.AppendLine("@CVDCESS_Per,")
            builder.AppendLine("@Excise_Percentage,")
            builder.AppendLine("@Permissible_Limit,")
            builder.AppendLine("@SDTax_Per,")
            builder.AppendLine("@ServiceTax_Per,")
            builder.AppendLine("@SECESS_Per,")
            builder.AppendLine("@CVDSECESS_Per,")
            builder.AppendLine("@SRSECESS_Per,")
            builder.AppendLine("@Tot_Add_Excise_PER,")
            builder.AppendLine("@bond17OpeningBal,")
            builder.AppendLine("@Ecess_TotalDuty_Per,")
            builder.AppendLine("@SEcess_TotalDuty_Per,")
            builder.AppendLine("@ADDVAT_Per,")
            builder.AppendLine("@BasicExciseAndCessValue,")
            builder.AppendLine("@total_quantity,")
            builder.AppendLine("@Ent_UserId,")
            builder.AppendLine("@Upd_Userid,")
            builder.AppendLine("@Vehicle_No,")
            builder.AppendLine("@From_Station,")
            builder.AppendLine("@To_Station,")
            builder.AppendLine("@Cust_Ref,")
            builder.AppendLine("@Amendment_No,")
            builder.AppendLine("@Print_DateTime,")
            builder.AppendLine("@Form3,")
            builder.AppendLine("@Carriage_Name,")
            builder.AppendLine("@Ref_Doc_No,")
            builder.AppendLine("@Cust_Name,")
            builder.AppendLine("@Currency_Code,")
            builder.AppendLine("@Nature_of_Contract,")
            builder.AppendLine("@OriginStatus,")
            builder.AppendLine("@Ctry_Destination_Goods,")
            builder.AppendLine("@Delivery_Terms,")
            builder.AppendLine("@Payment_Terms,")
            builder.AppendLine("@Pre_Carriage_By,")
            builder.AppendLine("@Receipt_Precarriage_at,")
            builder.AppendLine("@Vessel_Flight_number,")
            builder.AppendLine("@Port_Of_Loading,")
            builder.AppendLine("@Port_Of_Discharge,")
            builder.AppendLine("@Final_destination,")
            builder.AppendLine("@Mode_Of_Shipment,")
            builder.AppendLine("@Dispatch_mode,")
            builder.AppendLine("@Buyer_description_Of_Goods,")
            builder.AppendLine("@Invoice_description_of_EPC,")
            builder.AppendLine("@Buyer_Id,")
            builder.AppendLine("@REMARKS,")
            builder.AppendLine("@Excise_Type,")
            builder.AppendLine("@SRVDINO,")
            builder.AppendLine("@SRVLocation,")
            builder.AppendLine("@ConsigneeContactPerson,")
            builder.AppendLine("@ConsigneeAddress1,")
            builder.AppendLine("@ConsigneeAddress2,")
            builder.AppendLine("@ConsigneeAddress3,")
            builder.AppendLine("@ConsigneeECCNo,")
            builder.AppendLine("@ConsigneeLST,")
            builder.AppendLine("@SchTime,")
            builder.AppendLine("@invoice_time,")
            builder.AppendLine("@varGeneralRemarks,")
            builder.AppendLine("@CheckSheetNo,")
            builder.AppendLine("@Lorry_No,")
            builder.AppendLine("@OTL_No,")
            builder.AppendLine("@RefChallan,")
            builder.AppendLine("@price_bases,")
            builder.AppendLine("@LorryNo_date,")
            builder.AppendLine("@ConsInvString,")
            builder.AppendLine("@invoicepicking_status,")
            builder.AppendLine("@TMP_DOC_NO,")
            builder.AppendLine("@UNIT_CODE")

            Sqlcmd.CommandText = builder.ToString
            Sqlcmd.Parameters.Add("@Account_Code", SqlDbType.VarChar).Value = Me.txtCustCode.Text.Trim
            Sqlcmd.Parameters.Add("@CONSIGNEE_CODE", SqlDbType.VarChar).Value = String.Empty

            Sqlcmd.Parameters.Add("@Bill_Flag", SqlDbType.Bit).Value = False
            Sqlcmd.Parameters.Add("@FIFO_flag", SqlDbType.Bit).Value = False
            Sqlcmd.Parameters.Add("@Print_Flag", SqlDbType.Bit).Value = False
            Sqlcmd.Parameters.Add("@Cancel_flag", SqlDbType.Bit).Value = False
            Sqlcmd.Parameters.Add("@dataPosted", SqlDbType.Bit).Value = False
            Sqlcmd.Parameters.Add("@ftp", SqlDbType.Bit).Value = False
            Sqlcmd.Parameters.Add("@ExciseExumpted", SqlDbType.Bit).Value = False
            Sqlcmd.Parameters.Add("@ServiceInvoiceformatExport", SqlDbType.Bit).Value = False
            Sqlcmd.Parameters.Add("@Discount_Type", SqlDbType.Bit).Value = IIf(OptDiscountPercentage.Checked = True, True, False)
            Sqlcmd.Parameters.Add("@RejectionPosting", SqlDbType.Bit).Value = False
            Sqlcmd.Parameters.Add("@FOC_Invoice", SqlDbType.Bit).Value = False
            Sqlcmd.Parameters.Add("@PrintExciseFormat", SqlDbType.Bit).Value = False
            Sqlcmd.Parameters.Add("@FreshCrRecd", SqlDbType.Bit).Value = False
            Sqlcmd.Parameters.Add("@Trans_Parameter_Flag", SqlDbType.Bit).Value = False
            Sqlcmd.Parameters.Add("@InvoiceAgainstMultipleSO", SqlDbType.Bit).Value = False
            Sqlcmd.Parameters.Add("@TextFileGenerated", SqlDbType.Bit).Value = False
            Sqlcmd.Parameters.Add("@sameunitloading", SqlDbType.Bit).Value = False
            Sqlcmd.Parameters.Add("@postingFlag", SqlDbType.Bit).Value = False
            Sqlcmd.Parameters.Add("@MULTIPLESO", SqlDbType.Bit).Value = False


            Sqlcmd.Parameters.Add("@Suffix", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Transport_Type", SqlDbType.VarChar).Value = Mid(Trim(CmbTransType.Text), 1, 1)
            Sqlcmd.Parameters.Add("@Invoice_Type", SqlDbType.VarChar).Value = strInvType
            Sqlcmd.Parameters.Add("@Sub_Category", SqlDbType.VarChar).Value = strInvSubType
            Sqlcmd.Parameters.Add("@SalesTax_Type", SqlDbType.VarChar).Value = Me.txtSaleTaxType.Text
            Sqlcmd.Parameters.Add("@SalesTax_FormNo", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@other_ref", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Surcharge_salesTaxType", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@LoadingChargeTaxType", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@CustBankID", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@USLOC", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@TCSTax_Type", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@ECESS_Type", SqlDbType.VarChar).Value = Me.txtECSSTaxType.Text
            Sqlcmd.Parameters.Add("@SRCESS_Type", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@CVDCESS_Type", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@TurnOverTaxType", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@SDTax_Type", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@ServiceTax_Type", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@SECESS_Type", SqlDbType.VarChar).Value = txtSECSSTaxType.Text
            Sqlcmd.Parameters.Add("@CVDSECESS_Type", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@SRSECESS_Type", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@ISCHALLAN", SqlDbType.VarChar).Value = "N"
            Sqlcmd.Parameters.Add("@ISCONSOLIDATE", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Ecess_TotalDuty_Type", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@SEcess_TotalDuty_Type", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@ADDVAT_Type", SqlDbType.VarChar).Value = Me.txtAddVAT.Text
            Sqlcmd.Parameters.Add("@REMARKS", SqlDbType.VarChar).Value = Me.txtRemarks.Text

            Sqlcmd.Parameters.Add("@Invoice_Date", SqlDbType.VarChar).Value = Me.dtpDateDesc.Value.ToString("dd MMM yyyy")
            Sqlcmd.Parameters.Add("@Form3Date", SqlDbType.VarChar).Value = DBNull.Value
            Sqlcmd.Parameters.Add("@Exchange_Date", SqlDbType.VarChar).Value = DBNull.Value
            Sqlcmd.Parameters.Add("@Ent_dt", SqlDbType.VarChar).Value = GetServerDateTime.ToString("dd MMM yyyy HH:MM")
            Sqlcmd.Parameters.Add("@Upd_dt", SqlDbType.VarChar).Value = GetServerDateTime.ToString("dd MMM yyyy HH:MM")

            Sqlcmd.Parameters.Add("@Doc_No", SqlDbType.Decimal).Value = Val(strInvoiceNo)
            Sqlcmd.Parameters.Add("@NRGPNOIncaseOfServiceInvoice", SqlDbType.Decimal).Value = 0

            Sqlcmd.Parameters.Add("@Discount_Per", SqlDbType.Float).Value = IIf(Me.OptDiscountPercentage.Checked = True, Val(txtDiscountAmt.Text), 0)


            Sqlcmd.Parameters.Add("@Year", SqlDbType.Int).Value = Val(dtpDateDesc.Value.ToString("yyyy"))
            Sqlcmd.Parameters.Add("@Annex_no", SqlDbType.Int).Value = 0
            Sqlcmd.Parameters.Add("@pervalue", SqlDbType.Int).Value = 0
            Sqlcmd.Parameters.Add("@dataposted_fin", SqlDbType.Int).Value = 0

            Sqlcmd.Parameters.Add("@Location_Code", SqlDbType.VarChar).Value = Me.txtLocationCode.Text.Trim
            Sqlcmd.Parameters.Add("@To_Location", SqlDbType.VarChar).Value = DBNull.Value
            Sqlcmd.Parameters.Add("@From_Location", SqlDbType.VarChar).Value = "01T1"

            Sqlcmd.Parameters.Add("@Insurance", SqlDbType.Money).Value = Val(Me.ctlInsurance.Text)
            Sqlcmd.Parameters.Add("@Frieght_Tax", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@Sales_Tax_Amount", SqlDbType.Money).Value = Val(Me.lblSalesTaxValue.Text)
            Sqlcmd.Parameters.Add("@Surcharge_Sales_Tax_Amount", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@Frieght_Amount", SqlDbType.Money).Value = Val(txtFreight.Text)
            Sqlcmd.Parameters.Add("@Packing_Amount", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@SalesTax_FormValue", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@total_amount", SqlDbType.Money).Value = Val(LblNetInvoiceValue.Text)
            Sqlcmd.Parameters.Add("@TurnoverTax_per", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@Turnover_amt", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@LoadingChargeTaxAmount", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@Discount_Amount", SqlDbType.Money).Value = IIf(Me.OptDiscountValue.Checked = True, Val(txtDiscountAmt.Text), 0)
            Sqlcmd.Parameters.Add("@TCSTaxAmount", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@ECESS_Amount", SqlDbType.Money).Value = Val(Me.lblEcessValue.Text)
            Sqlcmd.Parameters.Add("@SRCESS_Amount", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@CVDCESS_Amount", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@TotalInvoiceAmtRoundOff_diff", SqlDbType.Money).Value = Val(lblRoundOff.Text)
            Sqlcmd.Parameters.Add("@SDTax_Amount", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@ServiceTax_Amount", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@Prev_Yr_ExportSales", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@Permissible_Limit_SmpExport", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@SECESS_Amount", SqlDbType.Money).Value = Val(lblHCessValue.Text)
            Sqlcmd.Parameters.Add("@CVDSECESS_Amount", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@SRSECESS_Amount", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@Tot_Add_Excise_Amt", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@Ecess_TotalDuty_Amount", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@SEcess_TotalDuty_Amount", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@ADDVAT_Amount", SqlDbType.Money).Value = Val(lblAddVATValue.Text)


            Sqlcmd.Parameters.Add("@Exchange_Rate", SqlDbType.Decimal).Value = 1.0
            Sqlcmd.Parameters.Add("@SalesTax_Per", SqlDbType.Decimal).Value = Val(Me.lblSaltax_Per.Text)
            Sqlcmd.Parameters.Add("@Surcharge_SalesTax_Per", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@LoadingChargeTax_Per", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@TCSTax_Per", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@ECESS_Per", SqlDbType.Decimal).Value = Val(Me.lblECSStax_Per.Text)
            Sqlcmd.Parameters.Add("@SRCESS_Per", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@CVDCESS_Per", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@Excise_Percentage", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@Permissible_Limit", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@SDTax_Per", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@ServiceTax_Per", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@SECESS_Per", SqlDbType.Decimal).Value = Val(Me.lblSECSStax_Per.Text)
            Sqlcmd.Parameters.Add("@CVDSECESS_Per", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@SRSECESS_Per", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@Tot_Add_Excise_PER", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@bond17OpeningBal", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@Ecess_TotalDuty_Per", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@SEcess_TotalDuty_Per", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@ADDVAT_Per", SqlDbType.Decimal).Value = Val(lblAddVAT.Text)


            Sqlcmd.Parameters.Add("@BasicExciseAndCessValue", SqlDbType.Decimal).Value = Val(Me.lblBasicExciseAndCess.Text)
            Sqlcmd.Parameters.Add("@total_quantity", SqlDbType.Real).Value = 1

            Sqlcmd.Parameters.Add("@Ent_UserId", SqlDbType.VarChar).Value = mP_User
            Sqlcmd.Parameters.Add("@Upd_Userid", SqlDbType.VarChar).Value = mP_User
            Sqlcmd.Parameters.Add("@Vehicle_No", SqlDbType.VarChar).Value = txtVehNo.Text.Trim
            Sqlcmd.Parameters.Add("@From_Station", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@To_Station", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Cust_Ref", SqlDbType.VarChar).Value = Me.txtRefNo.Text.Trim
            Sqlcmd.Parameters.Add("@Amendment_No", SqlDbType.VarChar).Value = Me.txtAmendNo.Text.Trim
            Sqlcmd.Parameters.Add("@Print_DateTime", SqlDbType.VarChar).Value = DBNull.Value
            Sqlcmd.Parameters.Add("@Form3", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Carriage_Name", SqlDbType.VarChar).Value = Me.txtCarrServices.Text
            Sqlcmd.Parameters.Add("@Ref_Doc_No", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Cust_Name", SqlDbType.VarChar).Value = Me.lblCustCodeDes.Text
            Sqlcmd.Parameters.Add("@Currency_Code", SqlDbType.VarChar).Value = "INR"
            Sqlcmd.Parameters.Add("@Nature_of_Contract", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@OriginStatus", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Ctry_Destination_Goods", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Delivery_Terms", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Payment_Terms", SqlDbType.VarChar).Value = Trim(txtCreditTerms.Text)
            Sqlcmd.Parameters.Add("@Pre_Carriage_By", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Receipt_Precarriage_at", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Vessel_Flight_number", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Port_Of_Loading", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Port_Of_Discharge", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Final_destination", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Mode_Of_Shipment", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Dispatch_mode", SqlDbType.VarChar).Value = DBNull.Value
            Sqlcmd.Parameters.Add("@Buyer_description_Of_Goods", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Invoice_description_of_EPC", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Buyer_Id", SqlDbType.VarChar).Value = String.Empty

            Sqlcmd.Parameters.Add("@Excise_Type", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@SRVDINO", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@SRVLocation", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@ConsigneeContactPerson", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@ConsigneeAddress1", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@ConsigneeAddress2", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@ConsigneeAddress3", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@ConsigneeECCNo", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@ConsigneeLST", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@SchTime", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@invoice_time", SqlDbType.VarChar).Value = GetServerDateTime.ToString("hh:mm tt")
            Sqlcmd.Parameters.Add("@varGeneralRemarks", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@CheckSheetNo", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Lorry_No", SqlDbType.VarChar).Value = DBNull.Value
            Sqlcmd.Parameters.Add("@OTL_No", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@RefChallan", SqlDbType.VarChar).Value = DBNull.Value
            Sqlcmd.Parameters.Add("@price_bases", SqlDbType.VarChar).Value = DBNull.Value
            Sqlcmd.Parameters.Add("@LorryNo_date", SqlDbType.VarChar).Value = DBNull.Value
            Sqlcmd.Parameters.Add("@ConsInvString", SqlDbType.VarChar).Value = DBNull.Value
            Sqlcmd.Parameters.Add("@invoicepicking_status", SqlDbType.VarChar).Value = DBNull.Value

            Sqlcmd.Parameters.Add("@TMP_DOC_NO", SqlDbType.Decimal).Value = Val(strInvoiceNo)
            Sqlcmd.Parameters.Add("@UNIT_CODE", SqlDbType.VarChar).Value = gstrUNITID
            Sqlcmd.ExecuteNonQuery()

            builder.Remove(0, builder.ToString.Length)

            builder.AppendLine("INSERT INTO SALES_DTL ")
            builder.AppendLine("(")
            builder.AppendLine("SupplementaryInvoiceFlag,")
            builder.AppendLine("Suffix,")
            builder.AppendLine("Excise_type,")
            builder.AppendLine("SalesTax_type,")
            builder.AppendLine("CVD_type,")
            builder.AppendLine("SAD_type,")
            builder.AppendLine("GL_code,")
            builder.AppendLine("SL_code,")
            builder.AppendLine("Discount_type,")
            builder.AppendLine("USLOC,")
            builder.AppendLine("Ent_dt,")
            builder.AppendLine("Upd_dt,")
            builder.AppendLine("Doc_No,")
            builder.AppendLine("TotalExciseAmount,")
            builder.AppendLine("From_Box,")
            builder.AppendLine("To_Box,")
            builder.AppendLine("Year,")
            builder.AppendLine("pervalue,")
            builder.AppendLine("Item_Code,")
            builder.AppendLine("Location_Code,")
            builder.AppendLine("To_Location,")
            builder.AppendLine("From_Location,")
            builder.AppendLine("Measure_Code,")
            builder.AppendLine("Rate,")
            builder.AppendLine("Sales_Tax,")
            builder.AppendLine("Excise_Tax,")
            builder.AppendLine("Packing,")
            builder.AppendLine("Others,")
            builder.AppendLine("Cust_Mtrl,")
            builder.AppendLine("Tool_Cost,")
            builder.AppendLine("Basic_Amount,")
            builder.AppendLine("Accessible_amount,")
            builder.AppendLine("CVD_Amount,")
            builder.AppendLine("SVD_amount,")
            builder.AppendLine("Discount_amt,")
            builder.AppendLine("Discount_perc,")
            builder.AppendLine("ItemPacking_Amount,")
            builder.AppendLine("ADD_EXCISE_AMOUNT,")
            builder.AppendLine("Sales_Quantity,")
            builder.AppendLine("Excise_per,")
            builder.AppendLine("CVD_per,")
            builder.AppendLine("SVD_per,")
            builder.AppendLine("CustMtrl_Amount,")
            builder.AppendLine("ToolCost_amount,")
            builder.AppendLine("BinQuantity,")
            builder.AppendLine("pkg_amount,")
            builder.AppendLine("csiexcise_amount,")
            builder.AppendLine("ADD_EXCISE_PER,")
            builder.AppendLine("Ent_UserId,")
            builder.AppendLine("Upd_UserId,")
            builder.AppendLine("Cust_Item_Code,")
            builder.AppendLine("Cust_Item_Desc,")
            builder.AppendLine("Cust_ref,")
            builder.AppendLine("Amendment_No,")
            builder.AppendLine("SRVDINO,")
            builder.AppendLine("SRVLocation,")
            builder.AppendLine("SchTime,")
            builder.AppendLine("Packing_Type,")
            builder.AppendLine("Item_remark,")
            builder.AppendLine("ADD_EXCISE_TYPE,")
            builder.AppendLine("RATEOFDUTY,")
            builder.AppendLine("EXCISEDUTYPERUNIT,")
            builder.AppendLine("UNIT_CODE,")
            builder.AppendLine("RATEOFDUTY_AED,")
            builder.AppendLine("AEDPERUNIT")
            builder.AppendLine(")")

            builder.AppendLine(" SELECT ") ' SELECT STATMENT
            builder.AppendLine("@SupplementaryInvoiceFlag,")
            builder.AppendLine("@Suffix,")
            builder.AppendLine("@Excise_type,")
            builder.AppendLine("@SalesTax_type,")
            builder.AppendLine("@CVD_type,")
            builder.AppendLine("@SAD_type,")
            builder.AppendLine("@GL_code,")
            builder.AppendLine("@SL_code,")
            builder.AppendLine("@Discount_type,")
            builder.AppendLine("@USLOC,")
            builder.AppendLine("@Ent_dt,")
            builder.AppendLine("@Upd_dt,")
            builder.AppendLine("@Doc_No,")
            builder.AppendLine("@TotalExciseAmount,")
            builder.AppendLine("@From_Box,")
            builder.AppendLine("@To_Box,")
            builder.AppendLine("@Year,")
            builder.AppendLine("@pervalue,")
            builder.AppendLine("@Item_Code,")
            builder.AppendLine("@Location_Code,")
            builder.AppendLine("@To_Location,")
            builder.AppendLine("@From_Location,")
            builder.AppendLine("@Measure_Code,")
            builder.AppendLine("@Rate,")
            builder.AppendLine("@Sales_Tax,")
            builder.AppendLine("@Excise_Tax,")
            builder.AppendLine("@Packing,")
            builder.AppendLine("@Others,")
            builder.AppendLine("@Cust_Mtrl,")
            builder.AppendLine("@Tool_Cost,")
            builder.AppendLine("@Basic_Amount,")
            builder.AppendLine("@Accessible_amount,")
            builder.AppendLine("@CVD_Amount,")
            builder.AppendLine("@SVD_amount,")
            builder.AppendLine("@Discount_amt,")
            builder.AppendLine("@Discount_perc,")
            builder.AppendLine("@ItemPacking_Amount,")
            builder.AppendLine("@ADD_EXCISE_AMOUNT,")
            builder.AppendLine("@Sales_Quantity,")
            builder.AppendLine("@Excise_per,")
            builder.AppendLine("@CVD_per,")
            builder.AppendLine("@SVD_per,")
            builder.AppendLine("@CustMtrl_Amount,")
            builder.AppendLine("@ToolCost_amount,")
            builder.AppendLine("@BinQuantity,")
            builder.AppendLine("@pkg_amount,")
            builder.AppendLine("@csiexcise_amount,")
            builder.AppendLine("@ADD_EXCISE_PER,")
            builder.AppendLine("@Ent_UserId,")
            builder.AppendLine("@Upd_UserId,")
            builder.AppendLine("@Cust_Item_Code,")
            builder.AppendLine("@Cust_Item_Desc,")
            builder.AppendLine("@Cust_ref,")
            builder.AppendLine("@Amendment_No,")
            builder.AppendLine("@SRVDINO,")
            builder.AppendLine("@SRVLocation,")
            builder.AppendLine("@SchTime,")
            builder.AppendLine("@Packing_Type,")
            builder.AppendLine("@Item_remark,")
            builder.AppendLine("@ADD_EXCISE_TYPE,")
            builder.AppendLine("@RATEOFDUTY,")
            builder.AppendLine("@EXCISEDUTYPERUNIT,")
            builder.AppendLine("@UNIT_CODE,")
            builder.AppendLine("@RATEOFDUTY_AED,")
            builder.AppendLine("@AEDPERUNIT")

            Sqlcmd.CommandText = builder.ToString
            Sqlcmd.Parameters.Clear()

            SpChEntry.Row = 1

            Sqlcmd.Parameters.Add("@SupplementaryInvoiceFlag", SqlDbType.Bit).Value = False
            Sqlcmd.Parameters.Add("@Suffix", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Excise_type", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@SalesTax_type", SqlDbType.VarChar).Value = Me.txtSaleTaxType.Text
            Sqlcmd.Parameters.Add("@CVD_type", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@SAD_type", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@GL_code", SqlDbType.VarChar).Value = DBNull.Value
            Sqlcmd.Parameters.Add("@SL_code", SqlDbType.VarChar).Value = DBNull.Value
            Sqlcmd.Parameters.Add("@Discount_type", SqlDbType.VarChar).Value = DBNull.Value
            Sqlcmd.Parameters.Add("@USLOC", SqlDbType.VarChar).Value = String.Empty

            Sqlcmd.Parameters.Add("@Ent_dt", SqlDbType.VarChar).Value = GetServerDateTime.ToString("dd MMM yyyy HH:MM")
            Sqlcmd.Parameters.Add("@Upd_dt", SqlDbType.VarChar).Value = GetServerDateTime.ToString("dd MMM yyyy HH:MM")

            Sqlcmd.Parameters.Add("@Doc_No", SqlDbType.Decimal).Value = Val(strInvoiceNo)
            Sqlcmd.Parameters.Add("@TotalExciseAmount", SqlDbType.Float).Value = Val(Me.lblExciseValue.Text)


            SpChEntry.Col = EnumInv.FROMBOX
            Sqlcmd.Parameters.Add("@From_Box", SqlDbType.Int).Value = Val(SpChEntry.Text)

            SpChEntry.Col = EnumInv.TOBOX
            Sqlcmd.Parameters.Add("@To_Box", SqlDbType.Int).Value = Val(SpChEntry.Text)

            Sqlcmd.Parameters.Add("@Year", SqlDbType.Int).Value = Me.dtpDateDesc.Value.ToString("yyyy")
            Sqlcmd.Parameters.Add("@pervalue", SqlDbType.Int).Value = 0

            SpChEntry.Col = EnumInv.ENUMITEMCODE
            Sqlcmd.Parameters.Add("@Item_Code", SqlDbType.VarChar).Value = SpChEntry.Text


            Sqlcmd.Parameters.Add("@Location_Code", SqlDbType.VarChar).Value = Me.txtLocationCode.Text

            Sqlcmd.Parameters.Add("@To_Location", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@From_Location", SqlDbType.VarChar).Value = DBNull.Value

            Sqlcmd.Parameters.Add("@Measure_Code", SqlDbType.VarChar).Value = String.Empty

            SpChEntry.Col = EnumInv.RATE_PERUNIT
            Sqlcmd.Parameters.Add("@Rate", SqlDbType.Money).Value = Val(SpChEntry.Text)


            Sqlcmd.Parameters.Add("@Sales_Tax", SqlDbType.Money).Value = Val(Me.lblSaltax_Per.Text)
            Sqlcmd.Parameters.Add("@Excise_Tax", SqlDbType.Money).Value = Val(Me.lblExciseValue.Text)
            Sqlcmd.Parameters.Add("@Packing", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@Others", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@Cust_Mtrl", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@Tool_Cost", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@Basic_Amount", SqlDbType.Money).Value = Val(Me.lblBasicValue.Text)
            Sqlcmd.Parameters.Add("@Accessible_amount", SqlDbType.Money).Value = Val(Me.lblAssValue.Text)
            Sqlcmd.Parameters.Add("@CVD_Amount", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@SVD_amount", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@Discount_amt", SqlDbType.Money).Value = Val(Me.txtDiscountAmt.Text)
            Sqlcmd.Parameters.Add("@Discount_perc", SqlDbType.Money).Value = IIf(Me.OptDiscountPercentage.Checked = True, 1, 0)
            Sqlcmd.Parameters.Add("@ItemPacking_Amount", SqlDbType.Money).Value = 0
            Sqlcmd.Parameters.Add("@ADD_EXCISE_AMOUNT", SqlDbType.Money).Value = Val(Me.lblAEDValue.Text)

            SpChEntry.Col = EnumInv.ENUMQUANTITY
            Sqlcmd.Parameters.Add("@Sales_Quantity", SqlDbType.Decimal).Value = Val(SpChEntry.Text)
            Sqlcmd.Parameters.Add("@Excise_per", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@CVD_per", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@SVD_per", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@CustMtrl_Amount", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@ToolCost_amount", SqlDbType.Decimal).Value = 0

            SpChEntry.Col = EnumInv.BINQTY
            Sqlcmd.Parameters.Add("@BinQuantity", SqlDbType.Decimal).Value = Val(SpChEntry.Text)

            Sqlcmd.Parameters.Add("@pkg_amount", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@csiexcise_amount", SqlDbType.Decimal).Value = 0
            Sqlcmd.Parameters.Add("@ADD_EXCISE_PER", SqlDbType.Decimal).Value = 0

            Sqlcmd.Parameters.Add("@Ent_UserId", SqlDbType.VarChar).Value = mP_User
            Sqlcmd.Parameters.Add("@Upd_UserId", SqlDbType.VarChar).Value = mP_User

            SpChEntry.Col = EnumInv.CUSTPARTNO
            Sqlcmd.Parameters.Add("@Cust_Item_Code", SqlDbType.VarChar).Value = SpChEntry.Text.Trim

            Sqlcmd.Parameters.Add("@Cust_Item_Desc", SqlDbType.VarChar).Value = Me.lblCustPartDesc.Text
            Sqlcmd.Parameters.Add("@Cust_ref", SqlDbType.VarChar).Value = Me.txtRefNo.Text.Trim
            Sqlcmd.Parameters.Add("@Amendment_No", SqlDbType.VarChar).Value = Me.txtAmendNo.Text.Trim
            Sqlcmd.Parameters.Add("@SRVDINO", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@SRVLocation", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@SchTime", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Packing_Type", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@Item_remark", SqlDbType.VarChar).Value = String.Empty
            Sqlcmd.Parameters.Add("@ADD_EXCISE_TYPE", SqlDbType.VarChar).Value = String.Empty

            SpChEntry.Col = EnumInv.ENUMQUANTITY
            Dim dblExcisePerUnit As Decimal
            Dim dblExciseRateofDuty As Decimal

            Dim dblAEDPerUnit As Decimal
            Dim dblAEDRateofDuty As Decimal

            If Val(lblExciseValue.Text) = 0 Then
                dblExcisePerUnit = 0
            Else
                dblExcisePerUnit = Val(lblExciseValue.Text) / Val(SpChEntry.Text)
            End If

            If Val(lblBasicValue.Text) = 0 Then
                dblExciseRateofDuty = 0
            Else
                'dblExciseRateofDuty = dblExcisePerUnit / Val(Me.lblBasicValue.Text) * 100
                dblExciseRateofDuty = (dblExcisePerUnit * Val(SpChEntry.Text)) / Val(Me.lblBasicValue.Text)
            End If



            If Val(lblAEDValue.Text) = 0 Then
                dblAEDPerUnit = 0
            Else
                dblAEDPerUnit = Val(lblAEDValue.Text) / Val(SpChEntry.Text)
            End If

            If Val(lblBasicValue.Text) = 0 Then
                dblAEDRateofDuty = 0
            Else
                dblAEDRateofDuty = (dblAEDPerUnit * Val(SpChEntry.Text)) / Val(Me.lblBasicValue.Text)
            End If

            Sqlcmd.Parameters.Add("@EXCISEDUTYPERUNIT", SqlDbType.Decimal).Value = dblExcisePerUnit
            Sqlcmd.Parameters.Add("@RATEOFDUTY", SqlDbType.Decimal).Value = dblExciseRateofDuty


            Sqlcmd.Parameters.Add("@UNIT_CODE", SqlDbType.VarChar).Value = gstrUNITID

            Sqlcmd.Parameters.Add("@AEDPERUNIT", SqlDbType.Decimal).Value = dblAEDPerUnit
            Sqlcmd.Parameters.Add("@RATEOFDUTY_AED", SqlDbType.Decimal).Value = dblAEDRateofDuty

            Sqlcmd.ExecuteNonQuery()


            builder.Remove(0, builder.ToString.Length)
            SpChEntry.Col = EnumInv.ENUMITEMCODE
            builder.AppendLine("INSERT INTO SALES_TRADING_GRIN_DTL")
            builder.AppendLine("(DOC_NO,LOCATION_CODE,GRIN_NO,GRIN_DOC_TYPE,ITEM_CODE,SLNO,SALESQTY,GRINQTY,REMQTY,KNOCKOFFQTY,PERPIECEEXCISE,UNIT_CODE,ENT_USERID,ENT_DT,UPD_USERID,UPD_DT,GRIN_PAGE_NO, PERPIECEAED)")
            builder.AppendLine("SELECT '" + strInvoiceNo + "','" + Me.txtLocationCode.Text + "',GRINNO,10,ITEM_CODE,SLNO,SALESQTY,GRINQTY,REMQTY,KNOCKOFFQTY,PERPIECEEXCISE,UNIT_CODE,'" + mP_User + "',GETDATE(),'" + mP_User + "',GETDATE(),GRIN_PAGE_NO, isnull(PERPIECEAED, 0)")
            builder.AppendLine("FROM TMP_TRADING_INV_GRINS WHERE UNIT_CODE='" + gstrUNITID + "' AND IPADDRESS='" + gstrIpaddressWinSck + "'")

            Sqlcmd.CommandText = builder.ToString
            Sqlcmd.Parameters.Clear()
            Sqlcmd.ExecuteNonQuery()

            '---------------------------STOCK UPDATION----------------------------------



            Dim strRetMessage As String
            SpChEntry.Col = EnumInv.ENUMITEMCODE
            If UpdateForSchedules("+", Sqlcmd, SpChEntry.Text, strCustDrgno, strDespatchQty, strRetMessage) = False Then
                IsTrans = False
                SqlTrans.Rollback()
                MsgBox(strRetMessage, MsgBoxStyle.Exclamation, ResolveResString(100))
                Exit Function
            End If

            Sqlcmd.CommandType = CommandType.Text
            Sqlcmd.Parameters.Clear()
            ' THIS WILL BE EXECUTED WHILE LOCKING THE INVOICE.
            'builder.Remove(0, builder.ToString.Length)
            'builder.AppendLine("UPDATE ITEMBAL_MST SET CUR_BAL=CUR_BAL-" + Val(strDespatchQty).ToString + " WHERE ITEM_CODE='" + strInternalCode + "' AND UNIT_CODE='" + gstrUNITID + "' ")
            'Sqlcmd.CommandText = builder.ToString
            'Sqlcmd.Parameters.Clear()
            'Sqlcmd.ExecuteNonQuery()

            'builder.Remove(0, builder.ToString.Length)
            'builder.AppendLine("UPDATE A SET DESPATCH_QTY_TRADING=ISNULL(DESPATCH_QTY_TRADING,0)+B.KNOCKOFFQTY")
            'builder.AppendLine("FROM GRN_DTL A,")
            'builder.AppendLine("(	")
            'builder.AppendLine("SELECT GRIN_NO,KNOCKOFFQTY,GRIN_DOC_TYPE,ITEM_CODE,UNIT_CODE ")
            'builder.AppendLine("FROM SALES_TRADING_GRIN_DTL WHERE UNIT_CODE='" + gstrUNITID + "'  AND DOC_NO=" + Val(strInvoiceNo).ToString)
            'builder.AppendLine(") B")
            'builder.AppendLine("WHERE A.DOC_NO=B.GRIN_NO")
            'builder.AppendLine("AND A.DOC_TYPE=B.GRIN_DOC_TYPE")
            'builder.AppendLine("AND A.ITEM_CODE=B.ITEM_CODE  ")
            'builder.AppendLine("AND A.UNIT_CODE=B.UNIT_CODE  ")
            'builder.AppendLine("AND A.UNIT_CODE='" + gstrUNITID + "'")
            'Sqlcmd.CommandText = builder.ToString
            'Sqlcmd.Parameters.Clear()
            'Sqlcmd.ExecuteNonQuery()
            'THIS WILL BE EXECUTED WHILE LOCKING THE INVOICE.

            '---------------------------STOCK UPDATION----------------------------------

            SqlTrans.Commit()
            Me.txtChallanNo.Text = strInvoiceNo
            IsTrans = False
            SaveData = True
        Catch ex As Exception
            If IsTrans = True Then SqlTrans.Rollback()
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            If IsTrans = True Then SqlTrans.Rollback()
            If IsNothing(Sqlcmd.Connection) = False Then
                If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
                Sqlcmd.Connection.Dispose()
            End If
            If IsNothing(Sqlcmd) = False Then
                Sqlcmd.Dispose()
            End If
            If IsNothing(SqlTrans) = False Then SqlTrans.Dispose()
        End Try
    End Function
    Private Function CalculatePackingValue(ByVal pintRowNo As Short, ByVal blnRoundoff As Boolean, ByRef grid As AxFPSpreadADO.AxfpSpread) As Double
        Dim strPkg_Type As String
        Dim ldblPkg_Per As Double
        Dim ldblRate As Double
        Dim lintQty As Double
        Dim rsTaxRate As ClsResultSetDB
        Dim intPackingRoundoff_Decimal As Short
        On Error GoTo ErrHandler
        With grid

            .Row = pintRowNo

            .Col = EnumInv.RATE_PERUNIT
            ldblRate = 0 'Val(.Text) / Val(ctlPerValue.Text)

            .Col = EnumInv.PACKING
            strPkg_Type = Trim(.Text)

            .Col = EnumInv.ENUMQUANTITY
            lintQty = Val(.Text)
            intPackingRoundoff_Decimal = Val(Find_Value("select PackingRoundoff_Decimal from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'"))
            rsTaxRate = New ClsResultSetDB
            rsTaxRate.GetResult("Select Txrt_Rate_no,TxRt_Percentage from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  Tx_TaxeID = 'PKT' and Txrt_Rate_no = '" & Trim(strPkg_Type) & "'")
            If rsTaxRate.GetNoRows > 0 Then

                ldblPkg_Per = rsTaxRate.GetValue("TxRt_Percentage")
            Else
                ldblPkg_Per = 0
            End If
            rsTaxRate.ResultSetClose()
            If blnRoundoff = True Then
                CalculatePackingValue = System.Math.Round((ldblRate * lintQty) * ldblPkg_Per / 100, 0)
            Else
                CalculatePackingValue = System.Math.Round((ldblRate * lintQty) * ldblPkg_Per / 100, intPackingRoundoff_Decimal)
            End If
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        CalculatePackingValue = 0
    End Function
    Public Function ReturnCustomerLocation() As String
        On Error GoTo Errorhandler
        Dim rsObject As New ClsResultSetDB
        Call rsObject.GetResult("Select Cust_Location=isnull(Cust_Location,'') from Customer_mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Customer_Code = '" & Trim(Me.txtCustCode.Text) & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsObject.RowCount > 0 Then

            ReturnCustomerLocation = rsObject.GetValue("Cust_Location")
        Else
            ReturnCustomerLocation = ""
        End If

        rsObject.ResultSetClose()
        Exit Function
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub CalculateTaxes()
        If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then

            If blnISBasicRoundOff = False Then
                dblBasicValue = Math.Round(dblBasicValue, intExciseRoundOffDecimal).ToString("############.00")
            End If

            lblExciseValue.Text = Math.Round(Val(lblExciseValue.Text)).ToString("############.00")
            lblAEDValue.Text = Math.Round(Val(lblAEDValue.Text)).ToString("############.00")

            lblEcessValue.Text = (Val(lblExciseValue.Text) * (Val(lblECSStax_Per.Text) / 100.0)).ToString("############.00")
            lblHCessValue.Text = (Val(lblExciseValue.Text) * (Val(lblSECSStax_Per.Text) / 100.0)).ToString("############.00")

            If blnECSSTax = False Then
                lblEcessValue.Text = Math.Round(Val(lblEcessValue.Text), intECSRoundOffDecimal).ToString("############.00")
                lblHCessValue.Text = Math.Round(Val(lblHCessValue.Text), intECSRoundOffDecimal).ToString("############.00")
            End If


            If blnISSalesTaxRoundOff = False Then
                lblSalesTaxValue.Text = Math.Round(Val(lblSalesTaxValue.Text), intSaleTaxRoundOffDecimal).ToString("############.00")
                lblAddVATValue.Text = Math.Round(Val(lblAddVATValue.Text), intSaleTaxRoundOffDecimal).ToString("############.00")
            End If

            Me.lblBasicExciseAndCess.Text = (dblBasicValue + Val(lblExciseValue.Text) + Val(lblAEDValue.Text) + Val(lblEcessValue.Text) + Val(lblHCessValue.Text)).ToString("############.00")

            If OptDiscountValue.Checked = True Then
                Me.lblAssValue.Text = Val(Me.lblBasicExciseAndCess.Text) - Val(txtDiscountAmt.Text)
            ElseIf OptDiscountPercentage.Checked = True Then
                Me.lblAssValue.Text = Val(Me.lblBasicExciseAndCess.Text) - (Val(Me.lblBasicExciseAndCess.Text) * Val(txtDiscountAmt.Text) / 100)
            Else
                Me.lblAssValue.Text = Me.lblBasicExciseAndCess.Text
            End If

            lblSalesTaxValue.Text = ((Val(Me.lblAssValue.Text)) * Val(lblSaltax_Per.Text) / 100.0).ToString("############.00")
            lblAddVATValue.Text = ((Val(Me.lblSalesTaxValue.Text)) * Val(lblAddVAT.Text) / 100.0).ToString("############.00")


            LblNetInvoiceValue.Text = (Val(Me.lblAssValue.Text) + Val(lblSalesTaxValue.Text) + Val(lblAddVATValue.Text) + Val(ctlInsurance.Text) + Val(txtFreight.Text)).ToString("###########.00")
            LblNetInvoiceValue.Tag = LblNetInvoiceValue.Text
            LblNetInvoiceValue.Text = Math.Round(Val(LblNetInvoiceValue.Text)).ToString("###########.00")
            lblRoundOff.Text = (Val(Me.LblNetInvoiceValue.Tag.ToString) - Val(Me.LblNetInvoiceValue.Text)).ToString("##########0.00")

        End If
    End Sub

    Private Sub IncludeDefaultTaxes()
        Dim strSql As String
        Dim Sqlcmd As New SqlCommand
        Dim Dr As SqlDataReader
        Dim SQLCon As SqlConnection
        Dim intX As Integer = 0
        SQLCon = SqlConnectionclass.GetConnection()
        Sqlcmd.Connection = SQLCon
        Sqlcmd.CommandType = CommandType.Text

        Try
            strSql = "SELECT SLNO,TXRT_RATE_NO,TXRT_PERCENTAGE "
            strSql = strSql + " FROM"
            strSql = strSql + " ("
            strSql = strSql + " SELECT 1 SLNO,TXRT_RATE_NO,TXRT_PERCENTAGE FROM GEN_TAXRATE "
            strSql = strSql + " WHERE UNIT_CODE='" + gstrUNITID + "' AND TX_TAXEID='ECT' AND TXRT_RATE_NO='ECT2' AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) )"
            strSql = strSql + " UNION ALL"
            strSql = strSql + " SELECT 2 SLNO,TXRT_RATE_NO,TXRT_PERCENTAGE FROM GEN_TAXRATE "
            strSql = strSql + " WHERE UNIT_CODE='" + gstrUNITID + "' AND (TX_TAXEID='ECSST') AND TXRT_RATE_NO='ECSST1' AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) "
            strSql = strSql + " OR (CAST(GETDATE() AS DATE) <= DEACTIVE_DATE))"
            strSql = strSql + " ) ABCD order by 1 "
            Sqlcmd.CommandText = strSql
            Dr = Sqlcmd.ExecuteReader
            If Dr.HasRows = True Then
                While Dr.Read
                    If intX = 0 Then
                        txtECSSTaxType.Text = Dr("TXRT_RATE_NO")
                        lblECSStax_Per.Text = Dr("TXRT_PERCENTAGE")
                    Else
                        txtSECSSTaxType.Text = Dr("TXRT_RATE_NO")
                        lblSECSStax_Per.Text = Dr("TXRT_PERCENTAGE")
                    End If
                    intX = intX + 1
                End While
            End If
            If Dr.IsClosed = False Then Dr.Close()
        Catch EX As Exception
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
        Finally
            If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
            If SQLCon.State = ConnectionState.Open Then SQLCon.Close()
            Sqlcmd.Connection.Dispose()
            Sqlcmd.Dispose()
            SQLCon.Dispose()
        End Try
    End Sub

    Private Function CheckExistanceOfFieldData(ByRef pstrFieldText As String, ByRef pstrColumnName As String, ByRef pstrTableName As String, Optional ByRef pstrCondition As String = "") As Boolean
        On Error GoTo ErrHandler
        CheckExistanceOfFieldData = False
        Dim strTableSql As String 'Declared To Make Select Query
        Dim rsExistData As ClsResultSetDB
        If Len(Trim(pstrCondition)) > 0 Then
            strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and " & pstrCondition
        Else
            strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "'"
        End If
        rsExistData = New ClsResultSetDB
        rsExistData.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsExistData.GetNoRows > 0 Then
            CheckExistanceOfFieldData = True
        Else
            CheckExistanceOfFieldData = False
        End If
        rsExistData.ResultSetClose()
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
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
    Public Function ReturnNoOfDecimals(ByRef pstrItemCode As String) As Short
        '***************************************************************************************
        'Name       :   ReturnNoOfDecimals
        'Type       :   Function
        'Author     :   Nisha Rai
        'Arguments  :   Item Code
        'Return     :   No of Decimals as intiger
        'Purpose    :   Fetches measure code of given item code and then according to decimal
        '               allowed fkag it returns No decimals allowed
        '***************************************************************************************
        Dim rsMeasurementUnit As ClsResultSetDB
        Dim rsNoOfDecimal As ClsResultSetDB
        Dim strMeasurementUnit As String
        Dim intNoofDecimals As Short
        On Error GoTo ErrHandler
        rsMeasurementUnit = New ClsResultSetDB
        rsMeasurementUnit.GetResult("Select Cons_Measure_Code,Pur_Measure_Code from Item_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  item_code = '" & pstrItemCode & "'")
        If rsMeasurementUnit.GetNoRows > 0 Then
            rsMeasurementUnit.MoveFirst()

            If UCase(CmbInvType.Text) = "REJECTION" Then
                strMeasurementUnit = rsMeasurementUnit.GetValue("Pur_Measure_Code")
            Else
                strMeasurementUnit = rsMeasurementUnit.GetValue("Cons_Measure_Code")
            End If
            rsNoOfDecimal = New ClsResultSetDB
            rsNoOfDecimal.GetResult("select Decimal_Allowed_Flag,NoOFDecimal from Measure_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Measure_Code = '" & strMeasurementUnit & "'")
            If rsNoOfDecimal.GetNoRows > 0 Then
                rsNoOfDecimal.MoveFirst()

                If rsNoOfDecimal.GetValue("Decimal_Allowed_Flag") = True Then

                    intNoofDecimals = Val(rsNoOfDecimal.GetValue("NoOFDecimal"))
                    If intNoofDecimals = 0 Then
                        intNoofDecimals = 2
                    End If
                    ReturnNoOfDecimals = intNoofDecimals
                Else
                    ReturnNoOfDecimals = 0
                End If
            End If
            rsNoOfDecimal.ResultSetClose()
        End If
        rsMeasurementUnit.ResultSetClose()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Sub ChangeCellTypeStaticText()
        On Error GoTo ErrHandler
        Dim intRow As Short
        Dim intcol As Short
        Dim varItemCode As Object
        Dim blnQtyChkAccToMeasureCode As Boolean
        Dim rsSalesParameter As ClsResultSetDB
        Dim strMin As String
        Dim strMax As String
        Dim intDecimal As Short
        Dim intloopcounter1 As Short
        Dim blnTrfInvoiceWithSO As Boolean
        If mblnBatchTrack = True And Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
            If Me.SpChEntry.MaxRows > 0 Then
                ReDim mBatchData(Me.SpChEntry.MaxRows)
            End If
            For intloopcounter1 = 1 To Me.SpChEntry.MaxRows
                ReDim mBatchData(intloopcounter1).Batch_No(0)
                ReDim mBatchData(intloopcounter1).Batch_Date(0)
                ReDim mBatchData(intloopcounter1).Batch_Quantity(0)
            Next
        End If
        Dim strQry As String
        Dim rsSOReq As ClsResultSetDB
        Dim dblMaxLimit As Double
        Dim str_NewCurrencyCode As String
        Dim int_NoOfDecimal As Short
        Dim str_Min As String
        Dim str_Max As String
        Dim intLoopCounter As Short
        Dim rs_currencycode As ClsResultSetDB
        Dim strInvSubType As String
        Dim rsSOReq1 As ClsResultSetDB
        Dim strqry1 As String
        With Me.SpChEntry
            Select Case Me.CmdGrpChEnt.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    blnTrfInvoiceWithSO = False

                    If (UCase(Trim(CmbInvType.Text)) = "NORMAL INVOICE") Or ((UCase(Trim(CmbInvType.Text)) = "TRANSFER INVOICE") And blnTrfInvoiceWithSO) Or (UCase(Trim(CmbInvType.Text)) = "EXPORT INVOICE") Or (UCase(Trim(CmbInvType.Text)) = "SERVICE INVOICE") Then

                        If UCase(Trim(CmbInvSubType.Text)) <> "SCRAP" Then
                            For intRow = 1 To .MaxRows
                                .Row = intRow
                                For intcol = 1 To .MaxCols
                                    .Col = intcol
                                    If intcol = EnumInv.ENUMQUANTITY Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                        .Lock = True
                                    ElseIf intcol = EnumInv.BINQTY Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = EnumInv.SelectGrin Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                                        .TypeButtonPicture = My.Resources.ico111.ToBitmap
                                        .Text = "..."
                                    Else
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                    End If
                                Next intcol
                            Next intRow
                        Else
                            For intRow = 1 To .MaxRows
                                .Row = intRow
                                For intcol = 1 To .MaxCols
                                    .Col = intcol
                                    If intcol = EnumInv.ENUMQUANTITY Or intcol = EnumInv.FROMBOX Or intcol = EnumInv.TOBOX Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = EnumInv.RATE_PERUNIT Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = EnumInv.BATCHCOL And mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                                    ElseIf intcol = EnumInv.BINQTY Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = EnumInv.SelectGrin Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                                        .TypeButtonPicture = My.Resources.ico111.ToBitmap
                                        .Text = "..."
                                    Else
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                    End If
                                Next intcol
                            Next intRow
                        End If
                    Else
                        For intRow = 1 To .MaxRows
                            .Row = intRow
                            For intcol = 1 To .MaxCols
                                .Col = intcol
                                If intcol = EnumInv.ENUMQUANTITY Or intcol = EnumInv.FROMBOX Or intcol = EnumInv.TOBOX Then
                                    If mblnRejTracking = True Then
                                        If intcol = EnumInv.ENUMQUANTITY And CmbInvType.Text = "REJECTION" Then
                                            SpChEntry.Col = EnumInv.ENUMQUANTITY
                                            SpChEntry.Row = intRow
                                            dblMaxLimit = .TypeFloatMax
                                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = dblMaxLimit
                                        Else
                                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                        End If
                                    Else
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    End If
                                ElseIf intcol = EnumInv.RATE_PERUNIT Then
                                    rs_currencycode = New ClsResultSetDB
                                    rs_currencycode.GetResult("Select currency_code from cust_ord_hdr WHERE UNIT_CODE='" + gstrUNITID + "' AND  account_code = " & "'" & Me.txtCustCode.Text & "'" & "and cust_ref = " & "'" & Me.txtRefNo.Text & "'")
                                    str_NewCurrencyCode = rs_currencycode.GetValue("Currency_code")
                                    rs_currencycode.ResultSetClose()
                                    int_NoOfDecimal = ToGetDecimalPlaces(Trim(str_NewCurrencyCode))
                                    If int_NoOfDecimal < 2 Then
                                        int_NoOfDecimal = 2
                                    End If
                                    str_Min = "0." : str_Max = "99999999999999."
                                    .Col = EnumInv.RATE_PERUNIT : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = int_NoOfDecimal : .TypeFloatMin = CDbl(str_Min) : .TypeFloatMax = CDbl(str_Max)

                                ElseIf intcol = EnumInv.BATCHCOL And mblnBatchTrack = True And UCase(CmbInvType.Text) <> "REJECTION" Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = "Batch Details"
                                ElseIf intcol = EnumInv.BINQTY Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                Else
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                End If
                            Next intcol
                        Next intRow
                    End If
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                    If Trim(strInvType) = "" Then
                        rsSalesParameter = New ClsResultSetDB
                        rsSalesParameter.GetResult("Select invoice_type from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no='" & Trim(txtChallanNo.Text) & "'")
                        If rsSalesParameter.GetNoRows > 0 Then
                            strInvType = rsSalesParameter.GetValue("invoice_type")
                        Else
                            strInvType = ""
                        End If
                        rsSalesParameter.ResultSetClose()
                    End If
                    rsSOReq1 = New ClsResultSetDB
                    strInvSubType = Find_Value("Select Sub_category from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no='" & Trim(txtChallanNo.Text) & "'")
                    strqry1 = "Select isnull(SORequired,0) as SORequired from saleConf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type ='" & strInvType & "' and Sub_Type='" & Trim(strInvSubType) & "' and  (fin_start_date <= getdate() and fin_end_date >= getdate())"
                    rsSOReq1.GetResult(strqry1)
                    If rsSOReq1.GetNoRows > 0 Then
                        blnTrfInvoiceWithSO = rsSOReq1.GetValue("SORequired")
                    Else
                        blnTrfInvoiceWithSO = False
                    End If
                    rsSOReq1.ResultSetClose()
                    If (UCase(strInvType) = "INV") Or ((UCase(strInvType) = "TRF") And blnTrfInvoiceWithSO) Or (UCase(strInvType) = "EXP") Or (UCase(strInvType) = "SRC") Then
                        If (UCase(strInvType) = "SRC") And mblnServiceInvoiceWithoutSO Then
                            For intRow = 1 To .MaxRows
                                .Row = intRow
                                For intcol = 1 To .MaxCols
                                    .Col = intcol
                                    If intcol = EnumInv.ENUMQUANTITY Or intcol = EnumInv.FROMBOX Or intcol = EnumInv.TOBOX Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = EnumInv.BATCHCOL And mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                                    ElseIf intcol = EnumInv.BINQTY Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    Else
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                    End If
                                    .Col = EnumInv.RATE_PERUNIT : .Lock = False
                                    .CtlEditMode = True : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    .Col = EnumInv.ENUMQUANTITY : .Lock = True

                                Next intcol
                            Next intRow
                        ElseIf (UCase(strInvSubType) <> "L") Then
                            For intRow = 1 To .MaxRows
                                .Row = intRow
                                For intcol = 1 To .MaxCols
                                    .Col = intcol
                                    If intcol = EnumInv.ENUMQUANTITY Or intcol = EnumInv.FROMBOX Or intcol = EnumInv.TOBOX Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat

                                    ElseIf intcol = EnumInv.BATCHCOL And mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                                    ElseIf intcol = EnumInv.BINQTY Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    Else
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                    End If
                                Next intcol
                            Next intRow
                        Else
                            For intRow = 1 To .MaxRows
                                .Row = intRow
                                For intcol = 1 To .MaxCols
                                    .Col = intcol
                                    If intcol = EnumInv.ENUMQUANTITY Or intcol = EnumInv.FROMBOX Or intcol = EnumInv.TOBOX Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = EnumInv.RATE_PERUNIT Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = EnumInv.BATCHCOL And mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                                    ElseIf intcol = EnumInv.BINQTY Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    Else
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                    End If
                                Next intcol
                            Next intRow
                        End If
                    Else
                        For intRow = 1 To .MaxRows
                            .Row = intRow
                            For intcol = 1 To .MaxCols
                                .Col = intcol
                                If intcol = EnumInv.ENUMQUANTITY Or intcol = EnumInv.FROMBOX Or intcol = EnumInv.TOBOX Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = EnumInv.RATE_PERUNIT Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = EnumInv.BATCHCOL And mblnBatchTrack = True And UCase(CmbInvType.Text) <> "REJECTION" Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                                ElseIf intcol = EnumInv.BINQTY Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                Else
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                End If
                            Next intcol
                        Next intRow
                    End If
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    If mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
                        With Me.SpChEntry
                            .Enabled = True
                            SpChEntry.Row = 1 : SpChEntry.Row2 = SpChEntry.MaxRows : SpChEntry.Col = 0 : SpChEntry.Col2 = SpChEntry.MaxCols - 1
                            SpChEntry.BlockMode = True : SpChEntry.Lock = True : SpChEntry.BlockMode = False
                            SpChEntry.Row = 1 : SpChEntry.Row2 = SpChEntry.MaxRows : SpChEntry.Col = EnumInv.BATCHCOL : SpChEntry.Col2 = EnumInv.BATCHCOL
                            SpChEntry.BlockMode = True : SpChEntry.Lock = False : SpChEntry.BlockMode = False
                        End With
                    Else
                        For intRow = 1 To .MaxRows
                            SpChEntry.Row = intRow
                            For intcol = 1 To SpChEntry.MaxCols
                                SpChEntry.Col = intcol
                                If intcol = EnumInv.ENUMQUANTITY Then
                                    SpChEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    SpChEntry.Lock = True
                                ElseIf intcol = EnumInv.BINQTY Then
                                    SpChEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = EnumInv.SelectGrin Then
                                    SpChEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                                    SpChEntry.TypeButtonPicture = My.Resources.ico111.ToBitmap
                                    SpChEntry.Text = "..."
                                Else
                                    SpChEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                End If
                            Next intcol
                        Next intRow
                    End If
            End Select
            rsSalesParameter = New ClsResultSetDB
            rsSalesParameter.GetResult("Select QtyChkAccToMeasureCode from Sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")
            If rsSalesParameter.GetNoRows > 0 Then
                If rsSalesParameter.GetValue("QtyChkAccToMeasureCode") = False Then
                    blnQtyChkAccToMeasureCode = False
                Else
                    blnQtyChkAccToMeasureCode = True
                End If
            End If
            rsSalesParameter.ResultSetClose()
            If blnQtyChkAccToMeasureCode = True Then
                For intRow = 1 To .MaxRows
                    varItemCode = Nothing
                    Call .GetText(EnumInv.ENUMITEMCODE, intRow, varItemCode)
                    If Len(Trim(varItemCode)) > 0 Then
                        intDecimal = ReturnNoOfDecimals(CStr(varItemCode))
                        strMin = "0." : strMax = "99999999999999."
                        For intloopcounter1 = 1 To intDecimal
                            strMin = strMin & "0"
                            strMax = strMax & "9"
                        Next
                        If intDecimal = 0 Then
                            strMin = "0" : strMax = "99999999999999"
                        End If
                        .Row = intRow : .Row2 = intRow : .Col = EnumInv.ENUMQUANTITY : .Col2 = EnumInv.ENUMQUANTITY : .BlockMode = True '.CellType = CellTypeFloat
                        .TypeFloatDecimalPlaces = intDecimal
                        .BlockMode = False
                    End If
                Next
            End If
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Public Function Find_Value(ByRef strField As String) As String
        On Error GoTo ErrHandler
        Dim Rs As New ADODB.Recordset
        Rs = New ADODB.Recordset
        Rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        Rs.Open(strField, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
        If Rs.RecordCount > 0 Then
            If IsDBNull(Rs.Fields(0).Value) = False Then
                Find_Value = Rs.Fields(0).Value
            Else
                Find_Value = ""
            End If
        Else
            Find_Value = ""
        End If
        Rs.Close()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Public Function GetDocumentDetail(ByRef strItem_code As String) As String
        On Error GoTo ErrHandler
        Dim strCompileString As String
        Dim vardoc_no As Object
        Dim varBatch_No As Object
        Dim varBatch_Date As Object
        Dim varBatchReq As Object
        Dim varQty As Object
        Dim intRow As Short
        Dim rsTmp As New ClsResultSetDB
        Dim strSql As String
        strSql = "Select REF_DOC_NO, Batch_No, Quantity from MKT_INVREJ_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_No=" & Trim(txtChallanNo.Text) & " and Item_code='" & strItem_code & "'"
        rsTmp.GetResult(strSql)
        strCompileString = ""
        If rsTmp.RowCount > 0 Then
            Do While Not rsTmp.EOFRecord
                strCompileString = strCompileString & rsTmp.GetValue("Ref_Doc_no") & "§" & rsTmp.GetValue("Batch_No") & "§§" & rsTmp.GetValue("Quantity") & "¶"
                rsTmp.MoveNext()
            Loop
        End If
        rsTmp.ResultSetClose()
        GetDocumentDetail = strCompileString
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Sub AddMaxAllowedQuanity(ByRef intRow As Short)
        On Error GoTo ErrHandler
        Dim dblqty As Double
        Dim dblStock As Double
        Dim varItemCode As Object
        Dim varMaxQuanity As Object
        Dim strsaledtl As String
        varItemCode = Nothing
        SpChEntry.GetText(EnumInv.ENUMITEMCODE, intRow, varItemCode)
        If Find_Value("Select top 1 REJ_TYPE from MKT_INVREJ_DTL  WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_No=" & txtChallanNo.Text) = "1" Then
            strsaledtl = "select MaxAllowedQty = SUM( ((a.Rejected_Quantity + a.excess_po_quantity) - (isnull(a.Despatch_Quantity,0) + isnull(a.Inspected_Quantity,0) + isnull(a.RGP_Quantity,0)))) from grn_Dtl a, grn_hdr b Where "
            strsaledtl = strsaledtl & " a.Doc_type = b.Doc_type and a.unit_code=b.unit_code and a.unit_code='" + gstrUNITID + "' And a.Doc_No = b.Doc_No and "
            strsaledtl = strsaledtl & " a.From_Location = b.From_Location and a.From_Location ='01R1'"
            strsaledtl = strsaledtl & " and a.Rejected_quantity > 0 and b.Vendor_code = '" & txtCustCode.Text
            strsaledtl = strsaledtl & "' and a.Doc_No in (" & Trim(txtRefNo.Text) & ") and a.Item_code = '" & varItemCode & "' AND ISNULL(b.GRN_Cancelled,0) = 0"
            strsaledtl = strsaledtl & " Group by a.Item_code "
            dblqty = Val(Find_Value(strsaledtl))
        ElseIf Find_Value("Select top 1 REJ_TYPE from MKT_INVREJ_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_No=" & txtChallanNo.Text) = "2" Then
            strsaledtl = "Select   MaxAllowedQty = Sum(rejected_Quantity) from LRN_HDR as a " & " Inner Join LRN_DTL as b on a.doc_No=b.doc_no and a.unit_code=b.unit_code and a.unit_code='" + gstrUNITID + "' and a.Doc_Type=b.doc_Type and a.from_Location=b.from_location " & " Where Authorized_Code Is Not Null " & " and a.Doc_No IN (" & Trim(txtRefNo.Text) & ") and ITem_code = '" & varItemCode & "'" & " Group by B.Item_Code "
            dblqty = Val(Find_Value(strsaledtl))
        End If
        dblStock = Val(Find_Value("Select Cur_Bal From ItemBal_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Location_code='01J1' and Item_code='" & varItemCode & "'"))
        If dblStock < dblqty Then
            varMaxQuanity = dblStock
        Else
            varMaxQuanity = dblqty
        End If

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Function SetMaxLengthInSpread(ByRef pintDecimalSize As Short) As Object
        On Error GoTo ErrHandler
        Dim intRow As Short
        Dim strMin As String
        Dim strMax As String
        Dim intLoopCounter As Short
        Dim intDecimal As Short
        If pintDecimalSize < 2 Then
            pintDecimalSize = 2
        End If
        strMin = "0." : strMax = "99999999999999."
        For intLoopCounter = 1 To intDecimal
            strMin = strMin & "0"
            strMax = strMax & "9"
        Next
        With Me.SpChEntry
            For intRow = 1 To .MaxRows
                .Row = intRow
                .Col = EnumInv.ENUMITEMCODE : .TypeMaxEditLen = 16
                .Col = EnumInv.CUSTPARTNO : .TypeMaxEditLen = 30
                .Col = EnumInv.RATE_PERUNIT : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax)
                .Col = EnumInv.CUSTSUPPMAT_PERUNIT : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax)
                .Col = EnumInv.ENUMQUANTITY : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("99999999999999.99")
                .Col = EnumInv.PACKING : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit

                .CtlEditMode = False
                If CmbInvType.Text = "NORMAL INVOICE" Or CmbInvType.Text = "JOBWORK INVOICE" Or CmbInvType.Text = "EXPORT INVOICE" Or (CmbInvType.Text = "SERVICE INVOICE" And Not mblnServiceInvoiceWithoutSO) Then
                    If UCase(Trim(CmbInvSubType.Text)) <> "SCRAP" Then
                        .CtlEditMode = False
                    Else
                        .CtlEditMode = True
                    End If
                Else
                    .CtlEditMode = True
                End If
                .Col = EnumInv.OTHERS_PERUNIT : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("99999999999999.99")
                .Col = EnumInv.CUMULATIVEBOXES : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 0 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("99999999999999.99")
                .Col = EnumInv.FROMBOX : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 0 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("999999.99")
                .Col = EnumInv.TOBOX : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 0 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("999999.99")
                .Col = EnumInv.BINQTY : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("999999.99")

                .CtlEditMode = True : .Enabled = True
                If mblnRejTracking = False Then
                    If mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
                        If mblnServiceInvoiceWithoutSO Then
                            .Col = EnumInv.BATCHCOL : .ColHidden = True
                        Else
                            .Col = EnumInv.BATCHCOL : .ColHidden = False : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = "Batch Details"
                        End If
                    End If
                Else
                    If UCase(CmbInvType.Text) = "REJECTION" Or UCase(strInvType) = "REJ" Then
                        .Col = EnumInv.BATCHCOL : .ColHidden = True
                    Else
                        If mblnBatchTrack = True And UCase(CmbInvType.Text) <> "REJECTION" Then
                            If mblnServiceInvoiceWithoutSO Then
                                .Col = EnumInv.BATCHCOL : .ColHidden = True
                            Else
                                .Col = EnumInv.BATCHCOL : .ColHidden = False : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = "Batch Details" : .set_ColWidth(EnumInv.BATCHCOL, 1200)
                            End If
                        End If
                    End If
                End If
            Next intRow
        End With
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub addRowAtEnterKeyPress(ByRef pintRows As Short)
        On Error GoTo ErrHandler
        Dim intRowHeight As Short
        With Me.SpChEntry
            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            For intRowHeight = 1 To pintRows
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .set_RowHeight(.Row, 300)
            Next intRowHeight
            If .MaxRows > 3 Then .ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsBoth
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub

    Public Function ToGetDecimalPlaces(ByRef pstrCurrency As String) As Short
        Dim rscurrency As ClsResultSetDB
        rscurrency = New ClsResultSetDB
        rscurrency.GetResult("Select Decimal_Place from Currency_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Currency_code ='" & pstrCurrency & "'")

        ToGetDecimalPlaces = Val(rscurrency.GetValue("Decimal_Place"))
        rscurrency.ResultSetClose()
    End Function

    Public Sub displayDeatilsfromCustOrdHdrandDtl()
        On Error GoTo ErrHandler
        Dim strCustOrdHdr As String
        Dim rsCustOrdHdr As ClsResultSetDB
        Dim strCurrency As String
        Dim intDecimalPlace As Short
        'To Get Data from Cusft_Ord_hdr
        Dim rsSOReq As ClsResultSetDB
        Dim blnSoReq As Boolean
        Select Case CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                strCustOrdHdr = "Select max(Order_date), SalesTax_Type,AddVAT_Type"
                strCustOrdHdr = strCustOrdHdr & "Currency_Code,PerValue from Cust_ord_hdr"
                strCustOrdHdr = strCustOrdHdr & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & txtCustCode.Text & "' and Cust_Ref ='"
                strCustOrdHdr = strCustOrdHdr & mstrRefNo & "'and Amendment_No ='" & mstrAmmNo & "' Group By salestax_type,AddVAT_Type,currency_code"
                rsCustOrdHdr = New ClsResultSetDB
                rsCustOrdHdr.GetResult(strCustOrdHdr, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                strCurrency = rsCustOrdHdr.GetValue("Currency_code")
                intDecimalPlace = ToGetDecimalPlaces(strCurrency)
                If intDecimalPlace < 2 Then
                    intDecimalPlace = 2
                End If

                txtSaleTaxType.Text = rsCustOrdHdr.GetValue("SalesTax_Type")

                Call txtSaleTaxType_Validating(txtSaleTaxType, New System.ComponentModel.CancelEventArgs(False))


                rsCustOrdHdr.ResultSetClose()
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                rsSOReq = New ClsResultSetDB
                rsSOReq.GetResult("Select isnull(SORequired,0) as SORequired from saleConf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type ='TRF' and Sub_Type_Description='" & Trim(CmbInvSubType.Text) & "' and  (fin_start_date <= getdate() and fin_end_date >= getdate())")
                If rsSOReq.GetNoRows > 0 Then
                    blnSoReq = rsSOReq.GetValue("SORequired")
                Else
                    blnSoReq = False
                End If
                rsSOReq.ResultSetClose()
                If UCase(CStr((Trim(CmbInvType.Text)) = "NORMAL INVOICE")) Or ((Trim(CmbInvType.Text) = "TRANSFER INVOICE") And blnSoReq) Or UCase(CStr((Trim(CmbInvType.Text)) = "JOBWORK INVOICE")) Or UCase(CStr((Trim(CmbInvType.Text)) = "EXPORT INVOICE")) Or (UCase(CStr((Trim(CmbInvType.Text)) = "SERVICE INVOICE")) And Not mblnServiceInvoiceWithoutSO) Then
                    If CBool(UCase(CStr((Trim(CmbInvSubType.Text)) <> "SCRAP"))) Then
                        If Len(Trim(txtRefNo.Text)) Then
                            strCustOrdHdr = "Select max(Order_date),SalesTax_Type,AddVAT_Type,Currency_code,PerValue,term_payment, surcharge_code from Cust_ord_hdr"
                            strCustOrdHdr = strCustOrdHdr & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & txtCustCode.Text & "' and Cust_Ref ='"
                            strCustOrdHdr = strCustOrdHdr & mstrRefNo & "'and Amendment_No ='" & mstrAmmNo & "'"
                            strCustOrdHdr = strCustOrdHdr & " and active_flag = 'A' Group by salestax_type,currency_code,AddVAT_Type,PerValue,term_payment, surcharge_code"
                            rsCustOrdHdr = New ClsResultSetDB
                            rsCustOrdHdr.GetResult(strCustOrdHdr, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            strSOSaleTaxType = rsCustOrdHdr.GetValue("SalesTax_Type")
                            'txtSaleTaxType.Text = rsCustOrdHdr.GetValue("SalesTax_Type")

                            strCurrency = rsCustOrdHdr.GetValue("Currency_code")
                            txtCreditTerms.Text = rsCustOrdHdr.GetValue("term_payment")

                            Dim RsTermMst As New ClsResultSetDB
                            RsTermMst.GetResult("SELECT CRTRM_TERMID,CRTRM_DESC FROM GEN_CREDITTRMMASTER WHERE UNIT_CODE='" + gstrUNITID + "'  AND CRTRM_TERMID='" + txtCreditTerms.Text + "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            If RsTermMst.GetNoRows > 0 Then
                                Me.lblCreditTermDesc.Text = RsTermMst.GetValue("CRTRM_DESC")
                            End If
                            RsTermMst.ResultSetClose()
                            RsTermMst = Nothing


                            intDecimalPlace = ToGetDecimalPlaces(strCurrency)
                            If intDecimalPlace < 2 Then
                                intDecimalPlace = 2
                            End If

                            rsCustOrdHdr.ResultSetClose()
                        End If
                    End If
                End If
        End Select
        Call DisplayDetailsInSpread(strCurrency)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Function DisplayDetailsInSpread(ByRef pstrCurrency As String) As Boolean
        'Description    -  To display Details From Sales_Dtl Acc To Location Code,Challan No and Drawing No
        On Error GoTo ErrHandler
        Dim intLoopCounter As Short
        Dim intRecordCount As Short
        Dim Intcounter As Short
        Dim inti As Short
        Dim strsaledtl As String
        Dim dblPacking As Double
        Dim varItem_Code As Object
        Dim varCustItemCode As Object
        Dim varItemAlready As Object
        Dim rsSalesDtl As ClsResultSetDB
        Dim rsBatch As ClsResultSetDB
        Dim rsSOReq As ClsResultSetDB
        Dim blnQtyChkAccToMeasureCode As Boolean
        Dim intDecimal As Short
        Dim strMin As String
        Dim strMax As String
        Dim intloopcounter1 As Short
        Dim strCompileString As String
        Dim blnSoReq As Boolean


        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strsaledtl = ""
                strsaledtl = "SELECT Location_Code,Doc_No,Suffix,Item_Code,Sales_Quantity,From_Box,To_Box,Rate,Sales_Tax,Excise_Tax,Packing,Others,Cust_Mtrl,Year,Cust_Item_Code,Cust_Item_Desc,Tool_Cost,Measure_Code,Excise_type,SalesTax_type,CVD_type,SAD_type,GL_code,SL_code,Basic_Amount,Accessible_amount,CVD_Amount,SVD_amount,Excise_per,CVD_per,SVD_per,CustMtrl_Amount,ToolCost_amount,pervalue,TotalExciseAmount,SupplementaryInvoiceFlag,To_Location,Discount_type,Discount_amt,Discount_perc,From_Location,Cust_ref,Amendment_No,SRVDINO,SRVLocation,USLOC,SchTime,BinQuantity,Packing_Type,ItemPacking_Amount,Item_remark,pkg_amount,csiexcise_amount,ADD_EXCISE_TYPE,ADD_EXCISE_PER,ADD_EXCISE_AMOUNT from Sales_Dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Location_Code='" & Trim(txtLocationCode.Text) & "'"
                strsaledtl = strsaledtl & " and Doc_No=" & Val(txtChallanNo.Text) & " and Cust_Item_Code in(" & Trim(mstrItemCode) & ")"
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                rsSOReq = New ClsResultSetDB
                rsSOReq.GetResult("Select isnull(SORequired,0) as SORequired from saleConf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type ='TRF' and Sub_Type_Description='" & Trim(CmbInvSubType.Text) & "' and  (fin_start_date <= getdate() and fin_end_date >= getdate())")
                If rsSOReq.GetNoRows > 0 Then
                    blnSoReq = rsSOReq.GetValue("SORequired")
                Else
                    blnSoReq = False
                End If
                rsSOReq.ResultSetClose()
                If UCase(CStr(Trim(CmbInvType.Text) = "NORMAL INVOICE")) Or (UCase(CStr(Trim(CmbInvType.Text) = "TRANSFER INVOICE")) And blnSoReq) Or UCase(CStr(Trim(CmbInvType.Text) = "JOBWORK INVOICE")) Or UCase(CStr(Trim(CmbInvType.Text) = "EXPORT INVOICE")) Or UCase(CStr(Trim(CmbInvType.Text) = "SERVICE INVOICE")) Then
                    strsaledtl = "Select Item_Code,Cust_DrgNo,Rate,Cust_Mtrl,Packing,Packing_Type,Others,tool_Cost,Excise_Duty from Cust_ord_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
                    strsaledtl = strsaledtl & "Account_Code ='" & txtCustCode.Text & "'and Cust_ref ='"
                    strsaledtl = strsaledtl & txtRefNo.Text & "' and Amendment_No = '" & mstrAmmNo & "'and "
                    strsaledtl = strsaledtl & " Active_flag ='A' and Cust_DrgNo in(" & mstrItemCode & ")"
                Else
                    strsaledtl = ""
                    strsaledtl = "SELECT Item_Code,standard_Rate from Item_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
                    strsaledtl = strsaledtl & " Status = 'A' and Hold_flag <> 1 and Item_Code in (" & mstrItemCode & ")"
                End If
        End Select
        rsSalesDtl = New ClsResultSetDB
        rsSalesDtl.GetResult(strsaledtl, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim intLoopCount As Short
        Dim varCumulative As Object
        Dim intcnt As Short
        Dim dblqty As Double
        Dim dblStock As Double
        Dim strpono As String
        Dim intMaxSerial_No As Short
        Dim strCustDrgNo As Object
        Dim strSqlBins As String
        Dim dblBins As Double
        Dim rsBinQty As ClsResultSetDB
        If rsSalesDtl.GetNoRows > 0 Then
            intRecordCount = rsSalesDtl.GetNoRows
            ReDim mdblPrevQty(intRecordCount - 1) ' To get value of Quantity in Arrey for updation in despatch
            ReDim mdblToolCost(intRecordCount - 1) ' To get value of Quantity i


            ' Call Me.SpChEntry.SetText(EnumInv.ENUMQUANTITY, 1, 0)
            '-----------------------
            SpChEntry.MaxRows = 0
            SpChEntry.MaxRows = SpChEntry.MaxRows + 1
            BlankTaxDetails()
            GetDefaultTaxexFromSO()
            FRMMKTTRN0076A.DeleteTmpTable()
            '-----------------------

            'If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            '    If SpChEntry.MaxRows > 0 Then
            '        varItemAlready = Nothing
            '        Call SpChEntry.GetText(EnumInv.ENUMITEMCODE, 1, varItemAlready)
            '        If Len(Trim(varItemAlready)) = 0 Then
            '            Call addRowAtEnterKeyPress(intRecordCount - 1)
            '        End If
            '    Else
            '        Call addRowAtEnterKeyPress(intRecordCount)
            '    End If
            'Else
            '    Call addRowAtEnterKeyPress(intRecordCount - 1)
            'End If



            rsSalesDtl.MoveFirst()
            If CmbInvType.Text = "NORMAL INVOICE" Or CmbInvType.Text = "JOBWORK INVOICE" Or CmbInvType.Text = "EXPORT INVOICE" Or (CmbInvType.Text = "SERVICE INVOICE" And Not mblnServiceInvoiceWithoutSO) Then
                If UCase(Trim(CmbInvSubType.Text)) <> "SCRAP" Then
                    For intLoopCount = 1 To intRecordCount
                        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                            mdblToolCost(intLoopCount - 1) = Val(rsSalesDtl.GetValue("Tool_Cost"))
                        Else
                            mdblToolCost(intLoopCount - 1) = Val(rsSalesDtl.GetValue("Tool_Cost"))
                        End If
                        rsSalesDtl.MoveNext()
                    Next
                End If
            End If
            rsSalesDtl.MoveFirst()
            intDecimal = ToGetDecimalPlaces(pstrCurrency)
            Call SetMaxLengthInSpread(intDecimal)



            'If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            '    If SpChEntry.MaxRows > 0 Then
            '        varItemAlready = Nothing
            '        Call SpChEntry.GetText(EnumInv.ENUMITEMCODE, 1, varItemAlready)
            '        If Len(Trim(varItemAlready)) > 0 Then
            '            inti = SpChEntry.MaxRows + 1
            '            SpChEntry.MaxRows = SpChEntry.MaxRows + intRecordCount
            '            intRecordCount = SpChEntry.MaxRows
            '        Else
            '            inti = 1
            '        End If
            '    Else
            '        inti = 1
            '        SpChEntry.MaxRows = intRecordCount
            '    End If
            'Else
            '    inti = 1
            'End If

            inti = SpChEntry.MaxRows

            For intLoopCounter = inti To intRecordCount
                With Me.SpChEntry
                    Select Case Me.CmdGrpChEnt.Mode
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                            .Row = 1 : .Row2 = .MaxRows : .Col = 0 : .Col2 = .MaxCols
                            .Enabled = True : .BlockMode = True : .Lock = True : .BlockMode = False
                            Call .SetText(EnumInv.ENUMITEMCODE, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
                            Call .SetText(EnumInv.CUSTPARTNO, intLoopCounter, rsSalesDtl.GetValue("Cust_Item_Code"))
                            Call .SetText(EnumInv.RATE_PERUNIT, intLoopCounter, rsSalesDtl.GetValue("Rate") * Val("CTLPERVALUE"))


                            Call .SetText(EnumInv.ENUMQUANTITY, intLoopCounter, rsSalesDtl.GetValue("Sales_Quantity"))
                            mdblPrevQty(intLoopCounter - 1) = Nothing
                            Call .GetText(EnumInv.ENUMQUANTITY, intLoopCounter, mdblPrevQty(intLoopCounter - 1))
                            If mblnRejTracking = True Then
                                Call AddMaxAllowedQuanity(intLoopCounter)
                            End If
                            Call .SetText(EnumInv.PACKING, intLoopCounter, rsSalesDtl.GetValue("Packing_Type"))



                            Call .SetText(EnumInv.FROMBOX, intLoopCounter, rsSalesDtl.GetValue("From_Box"))
                            Call .SetText(EnumInv.TOBOX, intLoopCounter, rsSalesDtl.GetValue("To_Box"))
                            Call .SetText(EnumInv.BINQTY, intLoopCounter, rsSalesDtl.GetValue("BinQuantity"))

                            If mblnRejTracking = True And mblnBatchTracking = True Then
                                strCompileString = GetDocumentDetail(rsSalesDtl.GetValue("Item_Code"))
                            End If
                            If intLoopCounter = 1 Then
                                Call .SetText(EnumInv.CUMULATIVEBOXES, intLoopCounter, (rsSalesDtl.GetValue("To_Box") - rsSalesDtl.GetValue("From_Box")) + 1)
                            Else
                                varCumulative = Nothing
                                Call .GetText(EnumInv.CUMULATIVEBOXES, intLoopCounter - 1, varCumulative)
                                Call .SetText(EnumInv.CUMULATIVEBOXES, intLoopCounter, varCumulative + ((rsSalesDtl.GetValue("To_Box") - rsSalesDtl.GetValue("From_Box")) + 1))
                            End If
                            If mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
                                varItem_Code = Nothing
                                Call .GetText(EnumInv.ENUMITEMCODE, intLoopCounter, varItem_Code)
                                varCustItemCode = Nothing
                                Call .GetText(EnumInv.CUSTPARTNO, intLoopCounter, varCustItemCode)
                                rsBatch = New ClsResultSetDB
                                Call rsBatch.GetResult("Select Batch_No,Batch_Date,Batch_Qty from ItemBatch_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Item_Code = '" & varItem_Code & "' and Cust_Item_Code = '" & varCustItemCode & "' and  From_Location = '" & mstrLocationCode & "' and Doc_No = " & Trim(Me.txtChallanNo.Text) & " and Doc_type = 9999 ", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                If rsBatch.RowCount > 0 Then
                                    ReDim Preserve mBatchData(intLoopCounter)
                                    ReDim Preserve mBatchData(intLoopCounter).Batch_No(rsBatch.RowCount)
                                    ReDim Preserve mBatchData(intLoopCounter).Batch_Date(rsBatch.RowCount)
                                    ReDim Preserve mBatchData(intLoopCounter).Batch_Quantity(rsBatch.RowCount)
                                    Intcounter = 1
                                    While Not rsBatch.EOFRecord
                                        mBatchData(intLoopCounter).Batch_No(Intcounter) = rsBatch.GetValue("Batch_No")
                                        mBatchData(intLoopCounter).Batch_Date(Intcounter) = ConvertToDate(VB6.Format(rsBatch.GetValue("Batch_Date"), gstrDateFormat))
                                        mBatchData(intLoopCounter).Batch_Quantity(Intcounter) = rsBatch.GetValue("Batch_Qty")
                                        Intcounter = Intcounter + 1
                                        rsBatch.MoveNext()
                                    End While
                                End If
                                rsBatch.ResultSetClose()
                            End If
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                            .Enabled = True
                            .Row = 1 : .Row2 = .MaxRows : .Col = 0 : .Col2 = .MaxCols : .BlockMode = True : .Lock = False : .set_RowHeight(.MaxRows, 12) : .BlockMode = False
                            If (Trim(CmbInvType.Text) = "NORMAL INVOICE") Or (UCase(CStr(Trim(CmbInvType.Text) = "TRANSFER INVOICE")) And blnSoReq) Or (Trim(CmbInvType.Text) = "JOBWORK INVOICE") Or (Trim(CmbInvType.Text) = "EXPORT INVOICE") Or (Trim(CmbInvType.Text) = "SERVICE INVOICE") Then

                                Call .SetText(EnumInv.ENUMITEMCODE, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
                                Call .SetText(EnumInv.CUSTPARTNO, intLoopCounter, rsSalesDtl.GetValue("Cust_DrgNo"))
                                Call .SetText(EnumInv.RATE_PERUNIT, intLoopCounter, (Val(rsSalesDtl.GetValue("Rate"))))

                                Call Me.SpChEntry.SetText(EnumInv.ENUMQUANTITY, intLoopCounter, 0)
                                Call .SetText(EnumInv.CUSTSUPPMAT_PERUNIT, intLoopCounter, (Val(rsSalesDtl.GetValue("Cust_Mtrl")) * Val(0)))

                                Call .SetText(EnumInv.PACKING, intLoopCounter, rsSalesDtl.GetValue("Packing_Type"))
                                Call .SetText(EnumInv.OTHERS_PERUNIT, intLoopCounter, (Val(rsSalesDtl.GetValue("Others")) * Val(0)))




                            Else
                                Call .SetText(EnumInv.ENUMITEMCODE, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
                                Call .SetText(EnumInv.CUSTPARTNO, intLoopCounter, rsSalesDtl.GetValue("Item_code"))
                                Call .SetText(EnumInv.RATE_PERUNIT, intLoopCounter, (rsSalesDtl.GetValue("Standard_Rate") * Val("CTLPERVALUE")))
                            End If
                            rsBinQty = New ClsResultSetDB
                            strCustDrgNo = Nothing
                            Call SpChEntry.GetText(EnumInv.CUSTPARTNO, intLoopCounter, strCustDrgNo)
                            strSqlBins = "Select isnull(BinQuantity,1) as BinQuantity from custitem_mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  cust_drgno= '" & strCustDrgNo & "' and Account_code='" & Trim(Me.txtCustCode.Text) & "' "
                            rsBinQty.GetResult(strSqlBins)
                            If rsBinQty.GetNoRows > 0 Then
                                If rsBinQty.GetValue("BinQuantity") = 0 Then
                                    dblBins = 1
                                Else
                                    dblBins = rsBinQty.GetValue("BinQuantity")
                                End If
                            Else
                                dblBins = 1
                            End If
                            rsBinQty.ResultSetClose()
                            Call SpChEntry.SetText(EnumInv.BINQTY, intLoopCounter, dblBins)
                    End Select
                End With
                rsSalesDtl.MoveNext()
            Next intLoopCounter

        End If

        If SpChEntry.MaxRows > 3 Then
            SpChEntry.ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsBoth
        End If
        rsSalesDtl.ResultSetClose()
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function


    Private Function GetMode() As String
        On Error GoTo ErrHandler
        Select Case CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                GetMode = "ADD"
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                GetMode = "EDIT"
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                GetMode = "VIEW"
        End Select
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Function SelectDataFromTable(ByRef mstrFieldName As String, ByRef mstrTableName As String, ByRef mstrCondition As String) As String
        Dim StrSQLQuery As String
        Dim GetDataFromTable As ClsResultSetDB
        On Error GoTo ErrHandler
        If UCase(mstrTableName) = "SALESCHALLAN_DTL" Then
            StrSQLQuery = "Select TOP 1 " & mstrFieldName & " From " & mstrTableName & mstrCondition
        Else
            StrSQLQuery = "Select " & mstrFieldName & " From " & mstrTableName & mstrCondition
        End If
        GetDataFromTable = New ClsResultSetDB
        If GetDataFromTable.GetResult(StrSQLQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic) Then
            If GetDataFromTable.GetNoRows > 0 Then
                GetDataFromTable.MoveFirst()
                SelectDataFromTable = GetDataFromTable.GetValue(mstrFieldName)
            Else
                SelectDataFromTable = ""
            End If
        Else
            SelectDataFromTable = ""
        End If
        GetDataFromTable.ResultSetClose()
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Public Sub ShowMultipleSOItemHelp()
        Dim frmMKTTRN0021B As New frmMKTTRN0021B
        On Error GoTo ErrHandler
        If Trim(txtCustCode.Text) = "" Then
            MsgBox("Select Customer Code before selecting items.", MsgBoxStyle.Information, ResolveResString(100))
            CmdCustCodeHelp.Focus()
            mblnCheckArray = False
            Exit Sub
        Else
            With frmMKTTRN0021B

                .mdtInvoiceDate = GetServerDate()

                .mstrCustomerCode = Trim(txtCustCode.Text)

                .mstrInvType = mstrInvTypenew

                .mstrInvSubType = mstrInvSubTypenew

                .mstrLocationCode = Trim(txtLocationCode.Text)

                .mstrStockLocation = mstrStockLocation

                .mstrDocNo = Trim(txtChallanNo.Text)

                .mstrMode = GetMode()

                If .FillSOItemHelp = True Then mblnCheckArray = True Else mblnCheckArray = False
            End With
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Function OriginalRefNoOVER(ByVal strRefNumber As String) As Boolean
        On Error GoTo ErrHandler
        '1st Check if Any Blank Amendment no for Ref No. Exists
        If SelectDataFromTable("Active_Flag", "Cust_ORD_HDR", " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code ='" & Trim(txtCustCode.Text) & "' AND Cust_Ref = '" & txtRefNo.Text & "' AND Amendment_No = ''") = "O" Then
            OriginalRefNoOVER = True
        Else
            OriginalRefNoOVER = False
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function

    Private Sub GetInvoiceType()
        Dim Sqlcmd As New SqlCommand
        Dim Rd As SqlDataReader
        Dim strSql As String
        Sqlcmd.Connection = SqlConnectionclass.GetConnection()
        Sqlcmd.CommandType = CommandType.Text
        Try
            strSql = "SELECT DISTINCT DESCRIPTION FROM SALECONF WHERE UNIT_CODE=@UNITCODE AND  INVOICE_TYPE  IN('INV') AND (FIN_START_DATE <= GETDATE() AND FIN_END_DATE >= GETDATE())  "
            Sqlcmd.CommandText = strSql
            Sqlcmd.Parameters.Clear()
            Sqlcmd.Parameters.Add("@UNITCODE", SqlDbType.VarChar).Value = gstrUNITID
            Rd = Sqlcmd.ExecuteReader()
            If Rd.HasRows = True Then
                While Rd.Read
                    CmbInvType.Items.Add(Rd("DESCRIPTION"))
                End While
            End If
            If Rd.IsClosed = False Then Rd.Close()
            CmbInvType.SelectedIndex = 0

            strSql = "SELECT SUB_TYPE_DESCRIPTION FROM SALECONF WHERE UNIT_CODE=@UNITCODE AND  INVOICE_TYPE  IN('INV') AND SUB_TYPE='T' AND (FIN_START_DATE <= GETDATE() AND FIN_END_DATE >= GETDATE())"
            Sqlcmd.CommandText = strSql
            Sqlcmd.Parameters.Clear()
            Sqlcmd.Parameters.Add("@UNITCODE", SqlDbType.VarChar).Value = gstrUNITID
            Rd = Sqlcmd.ExecuteReader()
            If Rd.HasRows = True Then
                While Rd.Read
                    CmbInvSubType.Items.Add(Rd("SUB_TYPE_DESCRIPTION"))
                End While
            End If
            If Rd.IsClosed = False Then Rd.Close()
            CmbInvSubType.SelectedIndex = 0

            CmbTransType.Items.Add("R - Road")
            CmbTransType.Items.Add("L - Rail")
            CmbTransType.Items.Add("S - Sea")
            CmbTransType.Items.Add("A - Air")
            CmbTransType.Items.Add("H - Hand")
            CmbTransType.Items.Add("C - Courier")
            CmbTransType.SelectedIndex = 0

        Catch Ex As Exception
            MsgBox(Ex.Message.ToString, MsgBoxStyle.Information, ResolveResString(100))
        Finally
            If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
            Sqlcmd.Connection.Dispose()
            Sqlcmd.Dispose()
        End Try
    End Sub


    Private Sub SelectDescriptionForField(ByRef pstrFieldName1 As String, ByRef pstrFieldName2 As String, ByRef pstrTableName As String, ByRef pContrName As System.Windows.Forms.Control, ByRef pstrControlText As String)
        On Error GoTo ErrHandler
        Dim strDesSql As String 'Declared to make Select Query
        Dim rsDescription As ClsResultSetDB
        If pstrFieldName2 = "Customer_Code" Then
            strDesSql = "Select " & Trim(pstrFieldName1) & " from " & Trim(pstrTableName) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  " & Trim(pstrFieldName2) & "='" & Trim(pstrControlText) & "'"
        Else
            strDesSql = "Select " & Trim(pstrFieldName1) & " from " & Trim(pstrTableName) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  " & Trim(pstrFieldName2) & "='" & Trim(pstrControlText) & "'"
        End If

        rsDescription = New ClsResultSetDB
        rsDescription.GetResult(strDesSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsDescription.GetNoRows > 0 Then
            pContrName.Text = rsDescription.GetValue(Trim(pstrFieldName1))
        End If
        rsDescription.ResultSetClose()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub BlankFields()
        lblCancelledInvoice.Visible = False
        blnBillFlag = False
        dblBasicValue = 0
        mstrItemCode = String.Empty
        dblGrinQuantityForSale = 0
        txtLocationCode.Text = gstrUNITID
        lblLocCodeDes.Text = String.Empty
        txtChallanNo.Text = String.Empty
        txtCustCode.Text = String.Empty
        lblCustCodeDes.Text = String.Empty
        txtRefNo.Text = String.Empty
        txtCarrServices.Text = String.Empty
        txtAmendNo.Text = String.Empty
        txtVehNo.Text = String.Empty
        lblAddressDes.Text = String.Empty
        txtRemarks.Text = String.Empty
        txtCreditTerms.Text = String.Empty
        lblCreditTermDesc.Text = String.Empty
        ctlInsurance.Text = "0.00"
        txtFreight.Text = "0.00"
        txtSaleTaxType.Text = String.Empty
        lblSaltax_Per.Text = "0.00"
        lblSalesTaxValue.Text = "0.00"
        txtECSSTaxType.Text = String.Empty
        lblECSStax_Per.Text = "0.00"
        lblEcessValue.Text = "0.00"
        txtSECSSTaxType.Text = String.Empty
        lblSECSStax_Per.Text = "0.00"
        lblHCessValue.Text = "0"
        txtDiscountAmt.Text = "0.00"
        lblCustPartDesc.Text = String.Empty
        lblCurrentStock.Text = String.Empty
        lblDespetchQty.Text = String.Empty
        Me.txtCarrServices.Text = String.Empty

        lblBasicValue.Text = "0.00"
        lblAssValue.Text = "0.00"
        lblExciseValue.Text = "0.00"
        lblEcessValue.Text = "0.00"
        lblHCessValue.Text = "0.00"
        lblSalesTaxValue.Text = "0.00"
        LblNetInvoiceValue.Text = "0.00"
        lblBasicExciseAndCess.Text = "0.00"
        lblAEDValue.Text = "0.00"

        txtAddVAT.Text = String.Empty
        lblAddVAT.Text = "0.00"
        lblAddVATValue.Text = "0.00"
        Me.lblRoundOff.Text = "0.00"
        strSOSaleTaxType = String.Empty

        Me.lblInternalPartDesc.Text = String.Empty
        Me.lblCustPartDesc.Text = String.Empty
        Me.dtpDateDesc.Value = GetServerDate.ToString("dd/MMM/yyyy")

        SpChEntry.Row = 1
        SpChEntry.Action = FPSpreadADO.ActionConstants.ActionDeleteRow
        SpChEntry.MaxRows = 0
        FRMMKTTRN0076A.DeleteTmpTable()
        GetItemDescription()

    End Sub
    Private Sub BlankTaxDetails()
        dblBasicValue = 0
        dblGrinQuantityForSale = 0
        ctlInsurance.Text = "0.00"
        txtFreight.Text = "0.00"
        txtSaleTaxType.Text = String.Empty
        lblSaltax_Per.Text = "0.00"
        lblSalesTaxValue.Text = "0.00"
        txtECSSTaxType.Text = String.Empty
        lblECSStax_Per.Text = "0.00"
        lblEcessValue.Text = "0.00"
        txtSECSSTaxType.Text = String.Empty
        lblSECSStax_Per.Text = "0.00"
        lblHCessValue.Text = "0"
        txtDiscountAmt.Text = "0.00"
        lblBasicValue.Text = "0.00"
        lblAssValue.Text = "0.00"
        lblExciseValue.Text = "0.00"
        lblEcessValue.Text = "0.00"
        lblHCessValue.Text = "0.00"
        lblSalesTaxValue.Text = "0.00"
        lblBasicExciseAndCess.Text = "0.00"
        lblAEDValue.Text = "0.00"

        txtAddVAT.Text = String.Empty
        lblAddVAT.Text = "0.00"
        lblAddVATValue.Text = "0.00"
    End Sub
    Private Function StockLocationSalesConf(ByRef pstrInvType As String, ByRef pstrInvSubtype As String, ByRef pstrFeild As String) As String
        Dim rsSalesConf As ClsResultSetDB
        Dim StockLocation As String
        On Error GoTo ErrHandler
        rsSalesConf = New ClsResultSetDB
        Select Case pstrFeild
            Case "DESCRIPTION"
                rsSalesConf.GetResult("Select Stock_Location from SaleConf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Description ='" & Trim(pstrInvType) & "' and Sub_type_Description ='" & Trim(pstrInvSubtype) & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
            Case "TYPE"
                rsSalesConf.GetResult("Select Stock_Location from SaleConf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_type ='" & Trim(pstrInvType) & "' and Sub_type ='" & Trim(pstrInvSubtype) & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
        End Select
        If rsSalesConf.GetNoRows > 0 Then
            StockLocation = rsSalesConf.GetValue("Stock_Location")
        End If
        rsSalesConf.ResultSetClose()
        StockLocationSalesConf = StockLocation
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function

#End Region
    Private Sub TRADINGINVOICE_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = "'" Then e.Handled = True
    End Sub
    Private Sub TRADINGINVOICE_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Visible = False
            Me.KeyPreview = True
            Me.MdiParent = mdifrmMain
            Me.Icon = mdifrmMain.Icon
            Call FitToClient(Me, PnlMain, ctlHeader, CmdGrpChEnt, 250)
            Me.Group1.Enabled = False
            Me.Group2.Enabled = False
            Me.Group3.Enabled = False
            Me.Group4.Enabled = False
            GetInvoiceType()
            BlankFields()
            SubGetRoundoffConfig()
            CmdGrpChEnt_ButtonClick(sender, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL))
        Catch Ex As Exception
            MsgBox(Ex.Message, MsgBoxStyle.Information, ResolveResString(100))
        Finally
            Me.Visible = True
        End Try
    End Sub
    Private Sub CmdCustCodeHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdCustCodeHelp.Click
        Dim strCustMst As String
        Dim rsCustMst As ClsResultSetDB
        On Error GoTo ErrHandler
        Dim strHelpString As String
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If Len(Trim(txtCustCode.Text)) = 0 Then
                    strHelpString = ShowList(1, (txtCustCode.MaxLength), "", "Customer_Code", "Cust_Name", "Customer_Mst", "  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                    If strHelpString = "-1" Then 'If No Record Found
                        Call ConfirmWindow(10225, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    Else
                        txtCustCode.Text = strHelpString
                        Call SelectDescriptionForField("Cust_Name", "Customer_Code", "Customer_Mst", lblCustCodeDes, (txtCustCode.Text))
                    End If
                Else
                    strHelpString = ShowList(1, (txtCustCode.MaxLength), "", "Customer_Code", "Cust_Name", "Customer_Mst", "  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                    If strHelpString = "-1" Then 'If No Record Found
                        Call ConfirmWindow(10225, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    Else
                        txtCustCode.Text = strHelpString
                        Call SelectDescriptionForField("Cust_Name", "Customer_Code", "Customer_Mst", lblCustCodeDes, (txtCustCode.Text))
                    End If
                End If
        End Select
        If Len(Trim(txtCustCode.Text)) > 0 Then
            rsCustMst = New ClsResultSetDB
            strCustMst = "SELECT Bill_Address1 + ', '  +  Bill_Address2 + ', ' + Bill_City + ' - ' + Bill_Pin as  invoiceAddress from Customer_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Customer_code ='" & txtCustCode.Text & "'  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
            rsCustMst.GetResult(strCustMst)
            If rsCustMst.GetNoRows > 0 Then
                lblAddressDes.Text = rsCustMst.GetValue("InvoiceAddress")
            End If
            rsCustMst.ResultSetClose()
            Me.txtRefNo.Focus()
        End If
        rsCustMst = Nothing
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub

    Private Sub CmdRefNoHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdRefNoHelp.Click
        Dim frmMKTTRN0009a As New frmMKTTRN0009a
        Dim frmMKTTRN0020 As New frmMKTTRN0020
        On Error GoTo ErrHandler
        If Len(txtCustCode.Text) = 0 Then
            Call ConfirmWindow(10416, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            txtCustCode.Focus()
            Exit Sub
        End If
        Dim strRefAmm As String
        Dim intPos As Short
        strRefAmm = frmMKTTRN0020.SelectDataFromCustOrd_Dtl(txtCustCode.Text, CmbInvType.Text)
        If Len(strRefAmm) > 0 Then
            intPos = InStr(1, Trim(strRefAmm), ",", CompareMethod.Text)
            mstrRefNo = Mid(Trim(strRefAmm), 2, intPos - 3)
            mstrAmmNo = Mid(strRefAmm, intPos + 2, ((Len(Trim(strRefAmm))) - intPos) - 2)
            txtRefNo.Text = Trim(mstrRefNo)
            txtAmendNo.Text = mstrAmmNo
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub PrintInvoice()
        Dim strSql As String
        Dim Sqlcmd As New SqlCommand
        Dim SQLCon As SqlConnection
        Dim ObjVal As Object
        Dim RdAddSold As ReportDocument
        Dim RepPath As String


        SQLCon = SqlConnectionclass.GetConnection()
        Sqlcmd.Connection = SQLCon
        Sqlcmd.CommandType = CommandType.Text
        Try
            Sqlcmd.CommandText = "SELECT REPORT_FILENAME FROM SALECONF WHERE UNIT_CODE='" + gstrUNITID + "' AND INVOICE_TYPE='INV' AND SUB_TYPE='T' AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE"
            ObjVal = Sqlcmd.ExecuteScalar()
            If IsNothing(ObjVal) = True Then ObjVal = String.Empty
            If ObjVal.ToString = String.Empty Then
                MsgBox("Report File Not Found In Sale Configuration.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If

            Sqlcmd.CommandType = CommandType.StoredProcedure
            Sqlcmd.CommandText = "TRADING_INVOICE_PRINT"
            Sqlcmd.Parameters.Add("@UNITCODE", SqlDbType.VarChar).Value = gstrUNITID
            Sqlcmd.Parameters.Add("@IPADDRESS", SqlDbType.VarChar).Value = gstrIpaddressWinSck
            Sqlcmd.Parameters.Add("@INVOICENO", SqlDbType.Decimal).Value = Val(Me.txtChallanNo.Text).ToString
            Sqlcmd.ExecuteNonQuery()

            Dim Frm As New eMProCrystalReportViewer_Inv
            RdAddSold = Frm.GetReportDocument()
            ' RepPath = "C:\Documents and Settings\amitrana\Desktop\" + ObjVal.ToString.Trim & ".rpt"
            RepPath = My.Application.Info.DirectoryPath & "\Reports\" & ObjVal.ToString.Trim & ".rpt"
            RdAddSold.Load(RepPath)
            RdAddSold.DataDefinition.RecordSelectionFormula = "{TRADING_TMP_INVOICE_PRINT_HDR.UNIT_CODE}='" + gstrUNITID + "' And {TRADING_TMP_INVOICE_PRINT_HDR.IPADDRESS}='" + gstrIpaddressWinSck + "'"
            Dim c As New CrystalDecisions.Shared.PageMargins
            c.rightMargin = 0
            c.leftMargin = 0
            c.topMargin = 0

            RdAddSold.PrintOptions.ApplyPageMargins(c)
            Dim Section As Section              'Defining the section of report
            Dim objFieldObject As PictureObject 'For storing field object of report
            Dim intSectionCount As Integer      'Counter for Sections in report    
            Dim intSectionFieldCount As Integer 'Counter for fields in the section

            Try

                For intSectionCount = 0 To RdAddSold.ReportDefinition.Sections.Count - 1
                    ' Get the Section object by name.
                    Section = RdAddSold.ReportDefinition.Sections.Item(intSectionCount)
                    ' Get the ReportObject by name and cast it as a FieldObject.
                    For intSectionFieldCount = 0 To Section.ReportObjects.Count - 1
                        If Section.ReportObjects(intSectionFieldCount).Kind = ReportObjectKind.PictureObject Then
                            objFieldObject = Section.ReportObjects(intSectionFieldCount)
                            If objFieldObject.Name.ToUpper = "INVIMAGE" Then
                                objFieldObject.Height = 0
                            End If

                        End If

                    Next
                Next intSectionCount

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Frm.SetReportDocument()
            Frm.Show()
        Catch EX As Exception
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
        Finally
            If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
            If SQLCon.State = ConnectionState.Open Then SQLCon.Close()
            Sqlcmd.Connection.Dispose()
            Sqlcmd.Dispose()
            SQLCon.Dispose()
        End Try
    End Sub
    Private Sub CmdGrpChEnt_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles CmdGrpChEnt.ButtonClick
        Try
            Select Case e.Button
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                    BlankFields()
                    FRMMKTTRN0076A.DeleteTmpTable()
                    Me.Group1.Enabled = True
                    Me.Group2.Enabled = True
                    Me.Group3.Enabled = True
                    Me.Group4.Enabled = True
                    Me.CmdChallanNo.Enabled = False
                    Me.Cmditems.Enabled = True
                    Me.CmdCustCodeHelp.Enabled = True
                    Me.CmdRefNoHelp.Enabled = True

                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = True
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                    Me.Group1.Enabled = True
                    Me.Group2.Enabled = True
                    Me.Group3.Enabled = True
                    Me.Group4.Enabled = True
                    Me.Cmditems.Enabled = False
                    Me.CmdCustCodeHelp.Enabled = False
                    Me.CmdRefNoHelp.Enabled = False
                    Me.CmdChallanNo.Enabled = False


                    If blnBillFlag = True Then ' IF INVOICE IS LOCKED
                        Me.Group1.Enabled = False
                        Me.Group2.Enabled = False
                        Me.Group3.Enabled = False
                        Me.SpChEntry.Row = 1
                        Me.SpChEntry.Col = EnumInv.BINQTY
                        Me.SpChEntry.Lock = True
                    Else
                        Me.SpChEntry.Row = 1
                        Me.SpChEntry.Col = EnumInv.BINQTY
                        Me.SpChEntry.Lock = False
                    End If


                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = True
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                    If SaveData() = True Then
                        Me.Group1.Enabled = False
                        Me.Group2.Enabled = False
                        Me.Group3.Enabled = False
                        Me.Group4.Enabled = False
                        Me.CmdChallanNo.Enabled = True
                        CmdGrpChEnt.Revert()
                        CmdGrpChEnt.Top = 578
                        CmdGrpChEnt.Left = 201
                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = False
                        MsgBox("Invoice Saved Successfully.", MsgBoxStyle.Information, ResolveResString(100))
                    End If
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                    BlankFields()
                    Me.Group1.Enabled = False
                    Me.Group2.Enabled = False
                    Me.Group3.Enabled = False
                    Me.Group4.Enabled = False
                    Me.CmdChallanNo.Enabled = True
                    Me.Cmditems.Enabled = False
                    Me.CmdCustCodeHelp.Enabled = False
                    Me.CmdRefNoHelp.Enabled = False
                    CmdGrpChEnt.Revert()

                    CmdGrpChEnt.Top = 578
                    CmdGrpChEnt.Left = 201

                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = False
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
                    If MsgBox("Are You Sure To Delete Select Invoice?", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                        If Me.txtChallanNo.Text.Trim = String.Empty Then
                            MsgBox("Please Select Invoice Number To Delete.", MsgBoxStyle.Information, ResolveResString(100))
                            Exit Sub
                        End If
                        If DeleteData() = True Then
                            CmdGrpChEnt_ButtonClick(Sender, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL))
                        End If
                    End If
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
                    PrintInvoice()
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                    Me.Close()
            End Select
        Catch Ex As Exception
            MsgBox(Ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub
    Private Sub cmdCreditHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCreditHelp.Click
        Try
            Dim strHelp() As String
            Dim strQuery As String
            strQuery = "SELECT CRTRM_TERMID,CRTRM_DESC FROM GEN_CREDITTRMMASTER WHERE UNIT_CODE='" + gstrUNITID + "'  AND CRTRM_STATUS=1 "
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Credit Terms")
            If UBound(strHelp) > 0 Then
                If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
                    MsgBox("Credit Terms Found.", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If
                txtCreditTerms.Text = strHelp(0)
                Me.lblCreditTermDesc.Text = strHelp(1)
            End If
        Catch Ex As Exception
            MsgBox(Ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        Finally
        End Try
    End Sub
    Private Sub Cmditems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmditems.Click
        Dim frmMKTTRN0021A As New frmMKTTRN0021A
        Dim frmMKTTRN0021B As New frmMKTTRN0021B
        Dim frmMKTTRN0021 As New frmMKTTRN0021
        On Error GoTo ErrHandler
        Dim rssalechallan As ClsResultSetDB
        Dim salechallan As String
        Dim strItemNotIn As String
        Dim varItemCode As Object
        Dim rsSaleConf As ClsResultSetDB
        Dim strStockLocation As String
        Dim rsCurrencyType As ClsResultSetDB
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim blnTrfInvoiceWithSO As Boolean
        With Me.SpChEntry
            If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                .MaxRows = 1
                .Row = 1 : .Row2 = .MaxRows : .Col = EnumInv.ENUMITEMCODE : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
            End If
        End With

        Dim strQry As String
        Dim rsSOReq As ClsResultSetDB
        Dim strtemp() As String
        frmMKTTRN0021.IsTradingInVoice = True
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                rssalechallan = New ClsResultSetDB
                salechallan = ""
                salechallan = "SELECT Invoice_type,SUB_CATEGORY FROM saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_No = "
                salechallan = salechallan & Val(txtChallanNo.Text)
                rssalechallan.GetResult(salechallan, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rssalechallan.GetNoRows > 0 Then
                    rssalechallan.MoveFirst()
                    strInvType = rssalechallan.GetValue("Invoice_type")
                    strInvSubType = rssalechallan.GetValue("sub_category")
                End If
                rssalechallan.ResultSetClose()
                strStockLocation = ""
                strStockLocation = StockLocationSalesConf(strInvType, strInvSubType, "TYPE")
                mstrLocationCode = Trim(strStockLocation)
                If (UCase(strInvType) = "SRC") And mblnServiceInvoiceWithoutSO Then
                    mstrItemCode = frmMKTTRN0021.SelectDatafromsaleDtl(Trim(txtChallanNo.Text))
                    If Len(Trim(mstrItemCode)) = 0 Then
                        SpChEntry.MaxRows = 0
                        frmMKTTRN0021 = Nothing
                    End If
                Else
                    If Len(Trim(strStockLocation)) > 0 Then

                        If (UCase(strInvType) = "INV") Or (UCase(strInvType) = "EXP") Or (UCase(strInvType) = "SRC") Then
                            mstrItemCode = frmMKTTRN0021.SelectDatafromsaleDtl(Trim(txtChallanNo.Text))
                            If Len(Trim(mstrItemCode)) = 0 Then
                                SpChEntry.MaxRows = 0
                                frmMKTTRN0021 = Nothing
                            End If
                        Else
                            mstrItemCode = frmMKTTRN0021.SelectDatafromsaleDtl(Trim(txtChallanNo.Text))
                            If Len(Trim(mstrItemCode)) = 0 Then
                                SpChEntry.MaxRows = 0
                                frmMKTTRN0021 = Nothing
                            End If
                        End If
                    Else
                        MsgBox("Please Define Stock Location in Sales Conf")
                        Exit Sub
                    End If
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
                    If MsgBox(" Do You Want To Follow FIFO Wise Batch Tracking ", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                        mblnbatchfifomode = True
                    Else
                        mblnbatchfifomode = False
                    End If
                End If

                rsSOReq = New ClsResultSetDB
                strQry = "Select isnull(SORequired,0) as SORequired from saleConf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type ='TRF' and Sub_Type_Description='" & Trim(CmbInvSubType.Text) & "' and  (fin_start_date <= getdate() and fin_end_date >= getdate())"
                rsSOReq.GetResult(strQry)
                If rsSOReq.GetNoRows > 0 Then
                    blnTrfInvoiceWithSO = rsSOReq.GetValue("SORequired")
                Else
                    blnTrfInvoiceWithSO = False
                End If
                rsSOReq.ResultSetClose()
                If UCase(CStr((Trim(CmbInvType.Text)) = "NORMAL INVOICE")) Or (UCase(CStr(Trim(CmbInvType.Text) = "TRANSFER INVOICE")) And blnTrfInvoiceWithSO) Or UCase(CStr((Trim(CmbInvType.Text)) = "JOBWORK INVOICE")) Or UCase(CStr((Trim(CmbInvType.Text)) = "EXPORT INVOICE")) Or (UCase(CStr((Trim(CmbInvType.Text)) = "SERVICE INVOICE")) And Not mblnServiceInvoiceWithoutSO) Then
                    If (UCase(CmbInvSubType.Text) <> "SCRAP" And mblnMultipleSOAllowed = False) Then
                        If Len(Trim(txtRefNo.Text)) = 0 Then
                            Call ConfirmWindow(10240, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            If txtRefNo.Enabled Then txtRefNo.Focus()
                            Exit Sub
                        ElseIf Len(Trim(txtAmendNo.Text)) = 0 Then
                            'User Can Enter Ref Code And Amendment From KeyBoard 1.Check If Ref No with Blank Amend is Over Or NOT
                            '   2.If Over Then see y No Amendments are added
                            If OriginalRefNoOVER(Trim(txtRefNo.Text)) Then
                                'Orig Ref No is OVER , So Amendment Number should be added
                                MsgBox("Enter Amendment No.", MsgBoxStyle.Information, ResolveResString(100))
                                If txtAmendNo.Enabled Then txtAmendNo.Focus() : Exit Sub
                            Else
                                'Original Ref No is Still Active
                                mstrRefNo = Trim(txtRefNo.Text)
                                mstrAmmNo = "" 'As Amend No is Blank
                            End If
                        Else
                            'Reference Number is Already Verfied in Validate Event of txtRefNo Amend No. is Also Verified in Its Validate Event
                            'Then Pass It to Form variables
                            mstrRefNo = Trim(txtRefNo.Text)
                            mstrAmmNo = Trim(txtAmendNo.Text)
                        End If
                    End If
                End If

                If SpChEntry.MaxRows > 0 Then
                    intMaxLoop = SpChEntry.MaxRows
                    strItemNotIn = ""
                    For intLoopCounter = 1 To intMaxLoop
                        With SpChEntry
                            varItemCode = Nothing
                            Call .GetText(EnumInv.ENUMITEMCODE, intLoopCounter, varItemCode)
                            If Len(Trim(strItemNotIn)) > 0 Then
                                strItemNotIn = Trim(strItemNotIn) & ",'" & Trim(varItemCode) & "'"
                            Else
                                strItemNotIn = "'" & Trim(varItemCode) & "'"
                            End If
                        End With
                    Next
                End If
                If False = True And 2 = 1 Then
                    ' to be removed
                Else
                    If UCase(CStr(Trim(CmbInvType.Text))) = "NORMAL INVOICE" Or UCase(CStr(Trim(CmbInvType.Text))) = "EXPORT INVOICE" Or UCase(CStr(Trim(CmbInvType.Text))) = "SERVICE INVOICE" Then
                        If UCase(Trim(CmbInvType.Text)) = "SERVICE INVOICE" And mblnServiceInvoiceWithoutSO Then
                            'If Len(Trim(strItemNotIn)) > 0 Then
                            '    mstrItemCode = frmMKTTRN0021.SelectDatafromItem_Mst(Trim(CmbInvType.Text), Trim(CmbInvSubType.Text), strStockLocation, , strItemNotIn, SpChEntry.MaxRows)
                            'Else
                            '    mstrItemCode = frmMKTTRN0021.SelectDatafromItem_Mst(Trim(CmbInvType.Text), Trim(CmbInvSubType.Text), strStockLocation)
                            'End If
                        Else
                            strStockLocation = ""
                            strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                            mstrLocationCode = strStockLocation
                            If Len(Trim(strStockLocation)) > 0 Then

                                If Len(Trim(strItemNotIn)) > 0 Then
                                    mstrItemCode = frmMKTTRN0021.SelectDataFromCustOrd_Dtl(Trim(txtCustCode.Text), Trim(txtRefNo.Text), mstrAmmNo, Trim(CmbInvSubType.Text), Trim(CmbInvType.Text), strStockLocation, strItemNotIn, SpChEntry.MaxRows)
                                Else
                                    mstrItemCode = frmMKTTRN0021.SelectDataFromCustOrd_Dtl(Trim(txtCustCode.Text), Trim(txtRefNo.Text), mstrAmmNo, Trim(CmbInvSubType.Text), Trim(CmbInvType.Text), strStockLocation)
                                End If
                                BlankTaxDetails()

                            Else
                                MsgBox("Please Define Stock Location in Sales Conf", MsgBoxStyle.Information, ResolveResString(100))
                                Exit Sub
                            End If
                            If Len(Trim(mstrItemCode)) = 0 Then SpChEntry.MaxRows = 0
                        End If

                    Else
                        rsSaleConf = New ClsResultSetDB
                        rsSaleConf.GetResult("select Stock_Location From saleconf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Description ='" & Trim(CmbInvType.Text) & "' and sub_type_description ='" & Trim(CmbInvSubType.Text) & "' AND Location_Code ='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                        If ((Len(Trim(rsSaleConf.GetValue("Stock_Location"))) = 0) Or (Trim(CStr(rsSaleConf.GetValue("Stock_Location") = "Unknown")))) Then
                            MsgBox("Plese Select Stock Location in SalesConf first", MsgBoxStyle.Information, ResolveResString(100))
                            If Cmditems.Enabled Then Cmditems.Focus()
                            Exit Sub
                        End If
                        mstrLocationCode = rsSaleConf.GetValue("Stock_Location")
                        rsSaleConf.ResultSetClose()
                        If Len(Trim(mstrItemCode)) = 0 And Len(Trim(strItemNotIn)) = 0 Then
                            SpChEntry.MaxRows = 0
                        Else
                            If Len(Trim(mstrItemCode)) = 0 Then
                            End If
                        End If
                    End If
                End If
        End Select
        Dim intDecimalPlace As Short
        Dim strCurrency As String
        If Len(mstrItemCode) > 0 Then
            mstrItemCode = Mid(mstrItemCode, 1, Len(mstrItemCode) - 1)
            Select Case Me.CmdGrpChEnt.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                    rsCurrencyType = New ClsResultSetDB
                    rsCurrencyType.GetResult("Select Currency_code from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_No = " & Val(txtChallanNo.Text))
                    If rsCurrencyType.GetNoRows > 0 Then
                        rsCurrencyType.MoveFirst()
                        strCurrency = rsCurrencyType.GetValue("Currency_code")
                    End If
                    rsCurrencyType.ResultSetClose()
                    intDecimalPlace = ToGetDecimalPlaces(strCurrency)
                    If intDecimalPlace < 2 Then
                        intDecimalPlace = 2
                    End If
                    DisplayDetailsInSpread(strCurrency) 'Procedure Call To Select Data >From Sales_Dtl
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    Call displayDeatilsfromCustOrdHdrandDtl()
                    System.Windows.Forms.Application.DoEvents()
            End Select
            If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                If CDbl(txtChallanNo.Text.Trim.Substring(0, 2)) = 99 Then
                    Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                    Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                End If
            End If

        Else
            frmMKTTRN0021 = Nothing
        End If
        'Set Cell Type In Spread
        Call ChangeCellTypeStaticText()
        Call GetItemDescription()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub

    Private Sub txtSaleTaxType_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtSaleTaxType.Validating
        '        Dim Cancel As Boolean = e.Cancel
        '        Dim strInvoiceType, strsql As String
        '        Dim rsChallanEntry, rsadditionaltax, rsadditionalVattax, rsadditionalsurcharge, rsadditionalVatsurcharge As ClsResultSetDB
        '        On Error GoTo ErrHandler
        '        If Len(txtSaleTaxType.Text) > 0 Then
        '            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
        '                strInvoiceType = UCase(Trim(CmbInvType.Text))
        '            ElseIf (CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT) Or (CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW) Then
        '                rsChallanEntry = New ClsResultSetDB
        '                rsChallanEntry.GetResult("Select a.Description,a.Sub_Type_Description from SaleConf a,SalesChallan_Dtl b where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND Doc_No = " & txtChallanNo.Text & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and a.Location_code = b.Location_code and (fin_start_date <= getdate() and fin_end_date >= getdate())")
        '                strInvoiceType = UCase(rsChallanEntry.GetValue("Description"))
        '                rsChallanEntry.ResultSetClose()
        '            End If
        '            If UCase(Trim(strInvoiceType)) <> "SERVICE INVOICE" Then
        '                If UCase(Trim(strInvoiceType)) <> "JOBWORK INVOICE" Then
        '                    If CheckExistanceOfFieldData((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST' OR Tx_TaxeID='VAT')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
        '                        lblSaltax_Per.Text = CStr(GetTaxRate((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST' OR Tx_TaxeID='VAT')"))

        '                    Else
        '                        Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
        '                        Cancel = True
        '                        txtSaleTaxType.Text = ""
        '                        If txtSaleTaxType.Enabled Then txtSaleTaxType.Focus()
        '                    End If
        '                Else
        '                    If CheckExistanceOfFieldData((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST' Or Tx_TaxeID='SRT' OR Tx_TaxeID='VAT')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
        '                        lblSaltax_Per.Text = CStr(GetTaxRate((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST' Or Tx_TaxeID='SRT' OR Tx_TaxeID='VAT')"))

        '                    Else
        '                        Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
        '                        Cancel = True
        '                        txtSaleTaxType.Text = ""
        '                        If txtSaleTaxType.Enabled Then txtSaleTaxType.Focus()
        '                    End If
        '                End If
        '            Else
        '                If CheckExistanceOfFieldData((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='SRT')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
        '                    lblSaltax_Per.Text = CStr(GetTaxRate((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='SRT')"))

        '                Else
        '                    Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
        '                    Cancel = True
        '                    txtSaleTaxType.Text = ""
        '                    If txtSaleTaxType.Enabled Then txtSaleTaxType.Focus()
        '                End If
        '            End If
        '        End If
        '        If UCase(Trim(GetPlantName)) = "MATM" And UCase(strInvoiceType) = "NORMAL INVOICE" Then
        '            strsql = " select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  (Tx_TaxeID='CST' OR Tx_TaxeID='LST') and txrt_percentage > 2.0 and TxRt_Rate_No='" & txtSaleTaxType.Text & "'  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date)) "
        '            rsadditionaltax = New ClsResultSetDB
        '            rsadditionaltax.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        '            If rsadditionaltax.GetNoRows > 0 Then
        '                rsadditionalsurcharge = New ClsResultSetDB
        '                strsql = " select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  Tx_TaxeID='SsT' and txrt_percentage=5.0"
        '                rsadditionalsurcharge.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        '                If rsadditionalsurcharge.GetNoRows > 0 Then
        '                    'txtSurchargeTaxType.Text = rsadditionalsurcharge.GetValue("TxRt_Rate_No")
        '                    'lblSurcharge_Per.Text = rsadditionalsurcharge.GetValue("TxRt_Percentage")
        '                End If
        '                rsadditionalsurcharge.ResultSetClose()
        '                rsadditionalsurcharge = Nothing
        '            End If
        '            rsadditionaltax.ResultSetClose()
        '            rsadditionaltax = Nothing
        '            strsql = " select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  (Tx_TaxeID='VAT') and TxRt_Rate_No='" & txtSaleTaxType.Text & "'  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
        '            rsadditionalVattax = New ClsResultSetDB
        '            rsadditionalVattax.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        '            If rsadditionalVattax.GetNoRows > 0 Then
        '                rsadditionalVatsurcharge = New ClsResultSetDB
        '                strsql = " select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  Tx_TaxeID='SsT' and txrt_percentage=5.0  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
        '                rsadditionalVatsurcharge.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        '                If rsadditionalVatsurcharge.GetNoRows > 0 Then
        '                    'txtSurchargeTaxType.Text = rsadditionalVatsurcharge.GetValue("TxRt_Rate_No")
        '                    'lblSurcharge_Per.Text = rsadditionalVatsurcharge.GetValue("TxRt_Percentage")
        '                End If
        '                rsadditionalVatsurcharge.ResultSetClose()
        '                rsadditionalVatsurcharge = Nothing
        '            End If
        '            rsadditionalVattax.ResultSetClose()
        '            rsadditionalVattax = Nothing
        '        End If
        '        GoTo EventExitSub
        'ErrHandler:  'The Error Handling Code Starts here
        '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        'EventExitSub:
        '        e.Cancel = Cancel
    End Sub

    Private Sub SpChEntry_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles SpChEntry.ButtonClicked
        Try
            If e.col = EnumInv.SelectGrin Then
                Dim FRM As New FRMMKTTRN0076A
                Dim dblItemRate As Double
                Dim dblBinQty As Double


                SpChEntry.Row = e.row
                SpChEntry.Col = EnumInv.RATE_PERUNIT
                dblItemRate = SpChEntry.Text

                SpChEntry.Col = EnumInv.ENUMITEMCODE
                FRM.strInternalPartNo = SpChEntry.Text

                SpChEntry.Col = EnumInv.CUSTPARTNO
                FRM.strCustPartNo = SpChEntry.Text

                FRM.strInternalPartDesc = Me.lblInternalPartDesc.Text
                FRM.strCustomerPartDesc = Me.lblCustPartDesc.Text

                FRM.strCurrentStockQty = Me.lblCurrentStock.Text

                If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    FRM.ParentFormOperationMode = "ADD"
                Else
                    FRM.ParentFormOperationMode = "EDIT"
                    FRM.strInvoiceNo = Me.txtChallanNo.Text
                End If
                FRM.blnBillFlag = blnBillFlag
                FRM.ShowDialog()

                If strGrinAllocationOKCancel = True Then
                    SpChEntry.Col = EnumInv.ENUMQUANTITY
                    SpChEntry.Text = dblGrinQuantityForSale.ToString
                    dblBasicValue = (dblGrinQuantityForSale * dblItemRate)
                    Me.lblBasicValue.Text = dblBasicValue.ToString("#############.00")
                    Me.lblAssValue.Text = dblBasicValue.ToString("#############.00")
                    IncludeDefaultTaxes()
                    CalculateTaxes()
                    SpChEntry_EditChange(Me, New AxFPSpreadADO._DSpreadEvents_EditChangeEvent(EnumInv.BINQTY, 1))
                    SpChEntry.Col = EnumInv.CUMULATIVEBOXES : SpChEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : SpChEntry.TypeFloatDecimalPlaces = 0 : SpChEntry.TypeFloatMin = CDbl("0.00") : SpChEntry.TypeFloatMax = CDbl("99999999999999.99") : SpChEntry.Lock = True
                    SpChEntry.Col = EnumInv.FROMBOX : SpChEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : SpChEntry.TypeFloatDecimalPlaces = 0 : SpChEntry.TypeFloatMin = CDbl("0.00") : SpChEntry.TypeFloatMax = CDbl("999999.99") : SpChEntry.Lock = True
                    SpChEntry.Col = EnumInv.TOBOX : SpChEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : SpChEntry.TypeFloatDecimalPlaces = 0 : SpChEntry.TypeFloatMin = CDbl("0.00") : SpChEntry.TypeFloatMax = CDbl("999999.99") : SpChEntry.Lock = True

                End If
            End If
        Catch Ex As Exception
            MsgBox(Ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try

    End Sub
    Private Sub SpChEntry_Change(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ChangeEvent) Handles SpChEntry.Change

    End Sub

    Private Sub txtECSSTaxType_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtECSSTaxType.Validating
        '        Dim Cancel As Boolean = e.Cancel
        '        On Error GoTo ErrHandler
        '        If Len(txtECSSTaxType.Text) > 0 Then
        '            If CheckExistanceOfFieldData((txtECSSTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='ECS')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
        '                lblECSStax_Per.Text = CStr(GetTaxRate((txtECSSTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='ECS')"))

        '                If OptDiscountValue.Enabled Then OptDiscountValue.Focus()


        '            Else
        '                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
        '                Cancel = True
        '                txtECSSTaxType.Text = ""
        '                If txtECSSTaxType.Enabled Then txtECSSTaxType.Focus()
        '            End If
        '        End If
        '        GoTo EventExitSub
        'ErrHandler:  'The Error Handling Code Starts here
        '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        'EventExitSub:
        '        e.Cancel = Cancel
    End Sub
    Private Sub CmdECSSTaxType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdECSSTaxType.Click
        Try
            Dim strHelp() As String
            Dim strQuery As String
            strQuery = "SELECT TXRT_RATE_NO,TXRT_PERCENTAGE FROM GEN_TAXRATE WHERE UNIT_CODE='" + gstrUNITID + "' AND TX_TAXEID='ECT' AND ((ISNULL(DEACTIVE_FLAG,0) <> 1))"
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "E Cess")
            If IsNothing(strHelp) = True Then
                txtECSSTaxType.Text = ""
                lblECSStax_Per.Text = ""
                CalculateTaxes()
                Exit Sub
            End If

            If UBound(strHelp) > 0 Then
                If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
                    MsgBox("Tax Rate Not Found.", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If
                txtECSSTaxType.Text = strHelp(0)
                lblECSStax_Per.Text = strHelp(1)
                CalculateTaxes()
            End If
        Catch Ex As Exception
            MsgBox(Ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        Finally
        End Try
    End Sub
    Private Sub BtnHCess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnHCess.Click
        Try
            Dim strHelp() As String
            Dim strQuery As String
            strQuery = "SELECT TXRT_RATE_NO,TXRT_PERCENTAGE FROM GEN_TAXRATE WHERE UNIT_CODE='" + gstrUNITID + "' AND TX_TAXEID='ECSST' AND ((ISNULL(DEACTIVE_FLAG,0) <> 1))"
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "H Cess")
            If IsNothing(strHelp) = True Then
                txtSECSSTaxType.Text = ""
                lblSECSStax_Per.Text = ""
                CalculateTaxes()
                Exit Sub
            End If

            If UBound(strHelp) > 0 Then
                If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
                    MsgBox("Tax Rate Not Found.", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If
                txtSECSSTaxType.Text = strHelp(0)
                lblSECSStax_Per.Text = strHelp(1)
                CalculateTaxes()
            End If
        Catch Ex As Exception
            MsgBox(Ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        Finally
        End Try
    End Sub

    Private Sub CmdSaleTaxType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdSaleTaxType.Click
        ' To Display Help From SaleTax Master
        On Error GoTo ErrHandler
        Dim strHelp() As String
        Dim rssalechallan, rsadditionaltax, rsadditionalVattax, rsadditionalsurcharge, rsadditionalVatsurcharge As ClsResultSetDB
        Dim salechallan, strsql As String
        Dim strInvoiceType As Object
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    rssalechallan = New ClsResultSetDB
                    salechallan = ""
                    salechallan = "select b.Description, b.Sub_type_Description from SalesChallan_dtl a,saleconf b where doc_no = " & Trim(txtChallanNo.Text)
                    salechallan = salechallan & " and a.Location_code = b.Location_code and a.unit_code=b.unit_code and a.unit_code='" + gstrUNITID + "' and a.Invoice_type = b.invoice_type and a.sub_category = b.Sub_type"
                    rssalechallan.GetResult(salechallan, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    If rssalechallan.GetNoRows > 0 Then
                        rssalechallan.MoveFirst()
                        strInvoiceType = rssalechallan.GetValue("Description")
                    End If
                    rssalechallan.ResultSetClose()
                Else
                    strInvoiceType = CmbInvType.Text
                End If



                Dim strQuery As String
                strQuery = "SELECT TXRT_RATE_NO,TXRT_PERCENTAGE FROM GEN_TAXRATE WHERE UNIT_CODE='" + gstrUNITID + "' And (TX_TAXEID='CSTT' OR TX_TAXEID='LSTT' OR TX_TAXEID='VATT')  AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) OR (CAST(GETDATE() AS DATE) <= DEACTIVE_DATE))"
                strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Sales Tax")

                If IsNothing(strHelp) = True Then
                    txtSaleTaxType.Text = ""
                    lblSaltax_Per.Text = ""
                    CalculateTaxes()
                    Exit Sub
                End If

                If UBound(strHelp) > 0 Then
                    If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
                        MsgBox("Tax Rate Not Found.", MsgBoxStyle.Information, ResolveResString(100))
                        Exit Sub
                    End If
                    txtSaleTaxType.Text = strHelp(0)
                    lblSaltax_Per.Text = strHelp(1)
                    CalculateTaxes()
                End If

        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub OptDiscountValue_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptDiscountValue.CheckedChanged
        If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            If OptDiscountValue.Checked = True Then
                OptDiscountPercentage.Checked = False
            ElseIf OptDiscountPercentage.Checked = True Then
                OptDiscountValue.Checked = False
            End If
            If Me.OptDiscountPercentage.Checked = False And Me.OptDiscountValue.Checked = False Then
                txtDiscountAmt.Text = "0.00"

                txtDiscountAmt.Enabled = False
            Else
                txtDiscountAmt.Enabled = True
            End If
            CalculateTaxes()
        End If
    End Sub
    Private Sub txtDiscountAmt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDiscountAmt.KeyPress
        AllowNumericValueInTextBox(txtDiscountAmt, e)
    End Sub
    Private Sub txtDiscountAmt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDiscountAmt.TextChanged
        If Me.OptDiscountPercentage.Checked = False And Me.OptDiscountValue.Checked = False Then
            txtDiscountAmt.Enabled = False
        End If
        CalculateTaxes()
    End Sub

    Private Sub OptDiscountPercentage_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptDiscountPercentage.CheckedChanged
        If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            If OptDiscountValue.Checked = True Then
                OptDiscountPercentage.Checked = False
            ElseIf OptDiscountPercentage.Checked = True Then
                OptDiscountValue.Checked = False
            End If
            If Me.OptDiscountPercentage.Checked = False And Me.OptDiscountValue.Checked = False Then
                txtDiscountAmt.Text = "0.00"

                txtDiscountAmt.Enabled = False
            Else
                txtDiscountAmt.Enabled = True
            End If
            CalculateTaxes()
        End If
    End Sub

    Private Sub ctlInsurance_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ctlInsurance.KeyPress
        AllowNumericValueInTextBox(ctlInsurance, e)
    End Sub

    Private Sub ctlInsurance_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ctlInsurance.TextChanged
        CalculateTaxes()
    End Sub

    Private Sub txtFreight_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFreight.KeyPress
        AllowNumericValueInTextBox(txtFreight, e)
    End Sub

    Private Sub txtFreight_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFreight.TextChanged
        CalculateTaxes()
    End Sub
    Private Sub CmdChallanNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdChallanNo.Click
        Try
            Dim strHelp() As String
            Dim strQuery As String
            strQuery = "SELECT DOC_NO INVOICE_NO,CONVERT(CHAR(11),INVOICE_DATE,106) INVOICE_DATE,ACCOUNT_CODE CUSTOMER_CODE,CASE WHEN ISNULL(CANCEL_FLAG,0)=1 THEN 'CANCELLED' ELSE '' END as Status FROM SALESCHALLAN_DTL WHERE INVOICE_TYPE='INV' AND SUB_CATEGORY='T' AND UNIT_CODE='" + gstrUNITID + "'  ORDER BY ENT_DT DESC"
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Invoice List")
            If IsNothing(strHelp) = True Then Exit Sub
            If UBound(strHelp) > 0 Then
                If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
                    MsgBox("No Any Saved Invoice Found.", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If
                GetSavedData(strHelp(0))
            End If
        Catch Ex As Exception
            MsgBox(Ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        Finally
        End Try
    End Sub
    Private Sub cmdAddVAT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddVAT.Click
        Dim strHelp As String
        Dim strSTaxHelp() As String
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strHelp = "Select TxRT_Rate_No,TxRt_Percentage from Gen_taxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  Tx_TaxeID in('ADVAT','ADCST')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                strSTaxHelp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, "Add. VAT/CST Tax Help")
                If IsNothing(strSTaxHelp) = True Then
                    txtAddVAT.Text = ""
                    lblAddVAT.Text = ""
                    CalculateTaxes()
                    Exit Sub
                End If
                If UBound(strSTaxHelp) <= 0 Then Exit Sub
                If strSTaxHelp(0) = "0" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtAddVAT.Text = "" : txtAddVAT.Focus() : Exit Sub
                Else
                    txtAddVAT.Text = strSTaxHelp(0)
                    lblAddVAT.Text = strSTaxHelp(1)
                End If
        End Select
        CalculateTaxes()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub SpChEntry_EditChange(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles SpChEntry.EditChange
        If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            If e.col = EnumInv.BINQTY Then
                Dim strSalqQuantity As String
                Dim strBinQty As String

                SpChEntry.Row = e.row
                SpChEntry.Col = EnumInv.ENUMQUANTITY
                strSalqQuantity = SpChEntry.Text

                SpChEntry.Col = EnumInv.BINQTY
                strBinQty = SpChEntry.Text

                If VAL(strBinQty) > 0 Then
                    SpChEntry.Col = EnumInv.FROMBOX
                    SpChEntry.Text = "1"

                    SpChEntry.Col = EnumInv.TOBOX
                    SpChEntry.Text = Math.Ceiling(Val(strSalqQuantity) / Val(strBinQty))

                    SpChEntry.Col = EnumInv.CUMULATIVEBOXES
                    SpChEntry.Text = Math.Ceiling(Val(strSalqQuantity) / Val(strBinQty))
                Else
                    SpChEntry.Col = EnumInv.FROMBOX
                    SpChEntry.Text = "0"

                    SpChEntry.Col = EnumInv.TOBOX
                    SpChEntry.Text = "0"

                End If
            End If
        End If
    End Sub

    Private Sub CmdGrpChEnt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdGrpChEnt.Load

    End Sub

    Private Sub SpChEntry_Advance(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_AdvanceEvent) Handles SpChEntry.Advance

    End Sub
End Class