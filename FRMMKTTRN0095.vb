'========================================================================================
'Copyright          :   Mothersonsumi Infotech & Design Ltd.
'Module             :   Marketing
'Author             :   Vipin Dubey [2288]
'Creation Date      :   01 June 2017
'Description        :   Form For Performa Invoice Entry GST.
'========================================================================================
Imports System.Data.SqlClient
Imports System.Text
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class FRMMKTTRN0095
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
        SL = 1
        ItemCode
        HSN_SAC_No
        HSN_SAC_Type
        Quantity
        Rate
        Basic_value
        DiscountPer
        DiscountVal
        Assable_Value
        Advance_Amt
        IGST_Tax_type
        IGST_Tax_Per
        IGST_Tax_Value
        CGST_Tax_type
        CGST_Tax_Per
        CGST_Tax_Value
        SGST_Tax_type
        SGST_Tax_Per
        SGST_Tax_Value
        UTGST_Tax_type
        UTGST_Tax_Per
        UTGST_Tax_Value
        CSESS_Tax_type
        CSESS_Tax_Per
        CSESS_Tax_Value
        ItemTotal
        Remarks
        Internal_Item_Desc
        Cust_Drgno
        Cust_DrgNo_Desc
        PrevQty

    End Enum

    '#Region "Functions And Subs"
    '    Private Sub GetDefaultTaxexFromSO()

    '        Dim Sqlcmd As New SqlCommand
    '        Dim strSql As String
    '        Dim ObjVal As Object = Nothing
    '        Dim SQLRd As SqlDataReader
    '        Dim builder As New StringBuilder

    '        Dim strDespatchQty As String = String.Empty
    '        Dim strCustDrgno As String = String.Empty
    '        Dim strInternalCode As String = String.Empty
    '        Try
    '            Sqlcmd.CommandTimeout = 0
    '            Sqlcmd.Connection = SqlConnectionclass.GetConnection
    '            Sqlcmd.CommandType = CommandType.Text
    '            If strSOSaleTaxType.Trim <> String.Empty Then
    '                Sqlcmd.CommandText = "SELECT TXRT_PERCENTAGE FROM GEN_TAXRATE WHERE UNIT_CODE='" + gstrUNITID + "' And LTRIM(RTRIM(TXRT_RATE_NO))='" + strSOSaleTaxType.Trim + "' AND (TX_TAXEID='CSTT' OR TX_TAXEID='LSTT' OR TX_TAXEID='VATT')  AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) OR (CAST(GETDATE() AS DATE) <= DEACTIVE_DATE))"
    '                ObjVal = Sqlcmd.ExecuteScalar()
    '                If IsNothing(ObjVal) = True Then ObjVal = String.Empty
    '                If ObjVal.ToString <> String.Empty Then
    '                    Me.txtSaleTaxType.Text = strSOSaleTaxType
    '                    Me.lblSaltax_Per.Text = ObjVal.ToString
    '                End If
    '            End If
    '        Catch ex As Exception
    '            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '        Finally
    '            If IsNothing(Sqlcmd.Connection.State) = False Then
    '                If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
    '            End If
    '            If IsNothing(Sqlcmd) = False Then
    '                Sqlcmd.Connection.Dispose()
    '                Sqlcmd.Dispose()
    '            End If
    '        End Try
    '    End Sub
    '    Private Sub SubGetRoundoffConfig()
    '        Dim Sqlcmd As New SqlCommand
    '        Dim ObjVal As Object = Nothing
    '        Dim DataRd As SqlDataReader
    '        Try

    '            Sqlcmd.CommandTimeout = 0
    '            Sqlcmd.Connection = SqlConnectionclass.GetConnection
    '            Sqlcmd.CommandType = CommandType.Text

    '            Sqlcmd.CommandText = "SELECT InsExc_Excise,CustSupp_Inc,EOU_Flag, Basic_Roundoff, Basic_Roundoff_decimal, SalesTax_Roundoff, SalesTax_Roundoff_decimal, Excise_Roundoff, Excise_Roundoff_decimal, "
    '            Sqlcmd.CommandText = Sqlcmd.CommandText + " SST_Roundoff, SST_Roundoff_decimal, InsInc_SalesTax, TCSTax_Roundoff, TCSTax_Roundoff_decimal, TotalToolCostRoundoff, TotalToolCostRoundoff_Decimal, ECESS_Roundoff, ECESSRoundoff_Decimal, ECESSOnSaleTax_Roundoff, ECESSOnSaleTaxRoundOff_Decimal, "
    '            Sqlcmd.CommandText = Sqlcmd.CommandText + " TurnOverTax_RoundOff, TurnOverTaxRoundOff_Decimal, TotalInvoiceAmount_RoundOff,TotalInvoiceAmountRoundOff_Decimal, SDTRoundOff, SDTRoundOff_Decimal,SameUnitLoading,ServiceTax_Roundoff,ServiceTaxRoundoff_Decimal=isnull(ServiceTaxRoundoff_Decimal,0),Packing_Roundoff,PackingRoundoff_Decimal=isnull(PackingRoundoff_Decimal,0) FROM Sales_Parameter WHERE UNIT_CODE='" + gstrUNITID + "'"
    '            DataRd = Sqlcmd.ExecuteReader()
    '            If DataRd.HasRows Then
    '                DataRd.Read()
    '                blnISInsExcisable = DataRd("InsExc_Excise")
    '                blnISBasicRoundOff = DataRd("Basic_Roundoff")
    '                blnISExciseRoundOff = DataRd("Excise_Roundoff")
    '                blnISSalesTaxRoundOff = DataRd("SalesTax_Roundoff")
    '                blnISSurChargeTaxRoundOff = DataRd("SST_Roundoff")
    '                blnAddCustMatrl = DataRd("CustSupp_Inc")
    '                blnInsIncSTax = DataRd("InsInc_SalesTax")
    '                blnTotalToolCostRoundOff = DataRd("TotalToolCostRoundoff")
    '                blnTCSTax = DataRd("TCSTax_Roundoff")
    '                intBasicRoundOffDecimal = DataRd("Basic_Roundoff_decimal")
    '                intSaleTaxRoundOffDecimal = DataRd("SalesTax_Roundoff_decimal")
    '                intExciseRoundOffDecimal = DataRd("Excise_Roundoff_decimal")
    '                intSSTRoundOffDecimal = DataRd("SST_Roundoff_decimal")
    '                intTCSRoundOffDecimal = DataRd("TCSTax_Roundoff_decimal")
    '                intToolCostRoundOffDecimal = DataRd("TotalToolCostRoundoff_decimal")
    '                blnECSSTax = DataRd("ECESS_Roundoff")
    '                intECSRoundOffDecimal = DataRd("ECESSRoundoff_Decimal")
    '                blnECSSOnSaleTax = DataRd("ECESSOnSaleTax_Roundoff")
    '                intECSSOnSaleRoundOffDecimal = DataRd("ECESSOnSaleTaxRoundOff_Decimal")
    '                blnTurnOverTax = DataRd("TurnOverTax_RoundOff")
    '                intTurnOverTaxRoundOffDecimal = DataRd("TurnOverTaxRoundOff_Decimal")
    '                blnTotalInvoiceAmount = DataRd("TotalInvoiceAmount_RoundOff")
    '                intTotalInvoiceAmountRoundOffDecimal = DataRd("TotalInvoiceAmountRoundOff_Decimal")
    '                blnIsSDTRoundoff = DataRd("SDTRoundOff")
    '                intSDTNoofDecimal = DataRd("SDTRoundOff_Decimal")
    '                blnSameUnitLoading = DataRd("SameUnitLoading")
    '                blnServiceTax_Roundoff = DataRd("ServiceTax_Roundoff")
    '                intServiceTaxRoundoff_Decimal = DataRd("ServiceTaxRoundoff_Decimal")
    '                blnPackingRoundoff = DataRd("Packing_Roundoff")
    '                intPackingRoundoff_Decimal = DataRd("PackingRoundoff_Decimal")

    '            Else
    '                MsgBox("No Data Define In Sales_Parameter Table", MsgBoxStyle.Critical, ResolveResString(100))
    '                Exit Sub
    '            End If
    '            If DataRd.IsClosed = False Then DataRd.Close()
    '        Catch ex As Exception
    '            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '        Finally
    '            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
    '            If IsNothing(Sqlcmd.Connection) = False Then
    '                If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
    '                Sqlcmd.Connection.Dispose()
    '            End If
    '            If IsNothing(Sqlcmd) = False Then
    '                Sqlcmd.Dispose()
    '            End If
    '        End Try

    '    End Sub
    '    Private Sub GetSavedData(ByVal strInvoiceNo As String)
    '        Dim Sqlcmd As New SqlCommand
    '        Dim ObjVal As Object = Nothing
    '        Dim intX As Int16
    '        Dim builder As New StringBuilder
    '        Dim SQLRd As SqlDataReader
    '        Try

    '            BlankTaxDetails()
    '            Sqlcmd.CommandTimeout = 0
    '            Sqlcmd.Connection = SqlConnectionclass.GetConnection
    '            Sqlcmd.CommandType = CommandType.Text

    '            builder.Remove(0, builder.ToString.Length)
    '            builder.AppendLine(" SELECT  ")
    '            builder.AppendLine("Account_Code,")
    '            builder.AppendLine("CONSIGNEE_CODE,")
    '            builder.AppendLine("Bill_Flag,")
    '            builder.AppendLine("FIFO_flag,")
    '            builder.AppendLine("Print_Flag,")
    '            builder.AppendLine("Cancel_flag,")
    '            builder.AppendLine("dataPosted,")
    '            builder.AppendLine("ftp,")
    '            builder.AppendLine("ExciseExumpted,")
    '            builder.AppendLine("ServiceInvoiceformatExport,")
    '            builder.AppendLine("Discount_Type,")
    '            builder.AppendLine("RejectionPosting,")
    '            builder.AppendLine("FOC_Invoice,")
    '            builder.AppendLine("PrintExciseFormat,")
    '            builder.AppendLine("FreshCrRecd,")
    '            builder.AppendLine("Trans_Parameter_Flag,")
    '            builder.AppendLine("InvoiceAgainstMultipleSO,")
    '            builder.AppendLine("TextFileGenerated,")
    '            builder.AppendLine("sameunitloading,")
    '            builder.AppendLine("postingFlag,")
    '            builder.AppendLine("MULTIPLESO,")
    '            builder.AppendLine("Suffix,")
    '            builder.AppendLine("Transport_Type,")
    '            builder.AppendLine("Invoice_Type,")
    '            builder.AppendLine("Sub_Category,")
    '            builder.AppendLine("SalesTax_Type,")
    '            builder.AppendLine("SalesTax_FormNo,")
    '            builder.AppendLine("other_ref,")
    '            builder.AppendLine("Surcharge_salesTaxType,")
    '            builder.AppendLine("LoadingChargeTaxType,")
    '            builder.AppendLine("CustBankID,")
    '            builder.AppendLine("USLOC,")
    '            builder.AppendLine("TCSTax_Type,")
    '            builder.AppendLine("ECESS_Type,")
    '            builder.AppendLine("SRCESS_Type,")
    '            builder.AppendLine("CVDCESS_Type,")
    '            builder.AppendLine("TurnOverTaxType,")
    '            builder.AppendLine("SDTax_Type,")
    '            builder.AppendLine("ServiceTax_Type,")
    '            builder.AppendLine("SECESS_Type,")
    '            builder.AppendLine("CVDSECESS_Type,")
    '            builder.AppendLine("SRSECESS_Type,")
    '            builder.AppendLine("ISCHALLAN,")
    '            builder.AppendLine("ISCONSOLIDATE,")
    '            builder.AppendLine("Ecess_TotalDuty_Type,")
    '            builder.AppendLine("SEcess_TotalDuty_Type,")
    '            builder.AppendLine("ADDVAT_Type,")
    '            builder.AppendLine("Invoice_Date,")
    '            builder.AppendLine("Form3Date,")
    '            builder.AppendLine("Exchange_Date,")
    '            builder.AppendLine("Ent_dt,")
    '            builder.AppendLine("Upd_dt,")
    '            builder.AppendLine("Doc_No,")
    '            builder.AppendLine("NRGPNOIncaseOfServiceInvoice,")
    '            builder.AppendLine("Discount_Per,")
    '            builder.AppendLine("Year,")
    '            builder.AppendLine("Annex_no,")
    '            builder.AppendLine("pervalue,")
    '            builder.AppendLine("dataposted_fin,")
    '            builder.AppendLine("Location_Code,")
    '            builder.AppendLine("To_Location,")
    '            builder.AppendLine("From_Location,")
    '            builder.AppendLine("Insurance,")
    '            builder.AppendLine("Frieght_Tax,")
    '            builder.AppendLine("Sales_Tax_Amount,")
    '            builder.AppendLine("Surcharge_Sales_Tax_Amount,")
    '            builder.AppendLine("Frieght_Amount,")
    '            builder.AppendLine("Packing_Amount,")
    '            builder.AppendLine("SalesTax_FormValue,")
    '            builder.AppendLine("total_amount,")
    '            builder.AppendLine("TurnoverTax_per,")
    '            builder.AppendLine("Turnover_amt,")
    '            builder.AppendLine("LoadingChargeTaxAmount,")
    '            builder.AppendLine("Discount_Amount,")
    '            builder.AppendLine("TCSTaxAmount,")
    '            builder.AppendLine("ECESS_Amount,")
    '            builder.AppendLine("SRCESS_Amount,")
    '            builder.AppendLine("CVDCESS_Amount,")
    '            builder.AppendLine("TotalInvoiceAmtRoundOff_diff,")
    '            builder.AppendLine("SDTax_Amount,")
    '            builder.AppendLine("ServiceTax_Amount,")
    '            builder.AppendLine("Prev_Yr_ExportSales,")
    '            builder.AppendLine("Permissible_Limit_SmpExport,")
    '            builder.AppendLine("SECESS_Amount,")
    '            builder.AppendLine("CVDSECESS_Amount,")
    '            builder.AppendLine("SRSECESS_Amount,")
    '            builder.AppendLine("Tot_Add_Excise_Amt,")
    '            builder.AppendLine("Ecess_TotalDuty_Amount,")
    '            builder.AppendLine("SEcess_TotalDuty_Amount,")
    '            builder.AppendLine("ADDVAT_Amount,")
    '            builder.AppendLine("Exchange_Rate,")
    '            builder.AppendLine("SalesTax_Per,")
    '            builder.AppendLine("Surcharge_SalesTax_Per,")
    '            builder.AppendLine("LoadingChargeTax_Per,")
    '            builder.AppendLine("TCSTax_Per,")
    '            builder.AppendLine("ECESS_Per,")
    '            builder.AppendLine("SRCESS_Per,")
    '            builder.AppendLine("CVDCESS_Per,")
    '            builder.AppendLine("Excise_Percentage,")
    '            builder.AppendLine("Permissible_Limit,")
    '            builder.AppendLine("SDTax_Per,")
    '            builder.AppendLine("ServiceTax_Per,")
    '            builder.AppendLine("SECESS_Per,")
    '            builder.AppendLine("CVDSECESS_Per,")
    '            builder.AppendLine("SRSECESS_Per,")
    '            builder.AppendLine("Tot_Add_Excise_PER,")
    '            builder.AppendLine("bond17OpeningBal,")
    '            builder.AppendLine("Ecess_TotalDuty_Per,")
    '            builder.AppendLine("SEcess_TotalDuty_Per,")
    '            builder.AppendLine("ADDVAT_Per,")
    '            builder.AppendLine("total_quantity,")
    '            builder.AppendLine("Ent_UserId,")
    '            builder.AppendLine("Upd_Userid,")
    '            builder.AppendLine("Vehicle_No,")
    '            builder.AppendLine("From_Station,")
    '            builder.AppendLine("To_Station,")
    '            builder.AppendLine("Cust_Ref,")
    '            builder.AppendLine("Amendment_No,")
    '            builder.AppendLine("Print_DateTime,")
    '            builder.AppendLine("Form3,")
    '            builder.AppendLine("Carriage_Name,")
    '            builder.AppendLine("Ref_Doc_No,")
    '            builder.AppendLine("Cust_Name,")
    '            builder.AppendLine("Currency_Code,")
    '            builder.AppendLine("Nature_of_Contract,")
    '            builder.AppendLine("OriginStatus,")
    '            builder.AppendLine("Ctry_Destination_Goods,")
    '            builder.AppendLine("Delivery_Terms,")
    '            builder.AppendLine("Payment_Terms,")
    '            builder.AppendLine("Pre_Carriage_By,")
    '            builder.AppendLine("Receipt_Precarriage_at,")
    '            builder.AppendLine("Vessel_Flight_number,")
    '            builder.AppendLine("Port_Of_Loading,")
    '            builder.AppendLine("Port_Of_Discharge,")
    '            builder.AppendLine("Final_destination,")
    '            builder.AppendLine("Mode_Of_Shipment,")
    '            builder.AppendLine("Dispatch_mode,")
    '            builder.AppendLine("Buyer_description_Of_Goods,")
    '            builder.AppendLine("Invoice_description_of_EPC,")
    '            builder.AppendLine("Buyer_Id,")
    '            builder.AppendLine("remarks,")
    '            builder.AppendLine("Excise_Type,")
    '            builder.AppendLine("SRVDINO,")
    '            builder.AppendLine("SRVLocation,")
    '            builder.AppendLine("ConsigneeContactPerson,")
    '            builder.AppendLine("ConsigneeAddress1,")
    '            builder.AppendLine("ConsigneeAddress2,")
    '            builder.AppendLine("ConsigneeAddress3,")
    '            builder.AppendLine("ConsigneeECCNo,")
    '            builder.AppendLine("ConsigneeLST,")
    '            builder.AppendLine("SchTime,")
    '            builder.AppendLine("invoice_time,")
    '            builder.AppendLine("varGeneralRemarks,")
    '            builder.AppendLine("CheckSheetNo,")
    '            builder.AppendLine("Lorry_No,")
    '            builder.AppendLine("OTL_No,")
    '            builder.AppendLine("RefChallan,")
    '            builder.AppendLine("price_bases,")
    '            builder.AppendLine("LorryNo_date,")
    '            builder.AppendLine("ConsInvString,")
    '            builder.AppendLine("invoicepicking_status,")
    '            builder.AppendLine("BasicExciseAndCessValue,")
    '            builder.AppendLine("TMP_DOC_No,")
    '            builder.AppendLine("'PAYMENT_TERMS_DESC'= ISNULL((SELECT CRTRM_DESC FROM GEN_CREDITTRMMASTER WHERE UNIT_CODE='" + gstrUNITID + "'  AND CRTRM_STATUS=1 AND GEN_CREDITTRMMASTER.CRTRM_TERMID=SALESCHALLAN_DTL.PAYMENT_TERMS ),''),")
    '            builder.AppendLine("UNIT_CODE")
    '            builder.AppendLine(" FROM SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "'  AND DOC_NO=" + Val(strInvoiceNo).ToString)

    '            Sqlcmd.CommandText = builder.ToString
    '            SQLRd = Sqlcmd.ExecuteReader
    '            If SQLRd.HasRows Then
    '                SQLRd.Read()

    '                Me.txtSaleTaxType.Text = SQLRd("SalesTax_Type").ToString
    '                Me.txtECSSTaxType.Text = SQLRd("ECESS_Type").ToString
    '                Me.txtSECSSTaxType.Text = SQLRd("SECESS_Type").ToString
    '                Me.dtpDateDesc.Value = Convert.ToDateTime(SQLRd("Invoice_Date")).ToString("dd/MMM/yyyy")
    '                Me.txtChallanNo.Text = strInvoiceNo
    '                Me.txtDiscountAmt.Text = SQLRd("Discount_Per").ToString
    '                Me.txtLocationCode.Text = SQLRd("Location_Code").ToString
    '                Me.ctlInsurance.Text = SQLRd("Insurance").ToString
    '                Me.lblSalesTaxValue.Text = SQLRd("Sales_Tax_Amount").ToString
    '                Me.txtFreight.Text = SQLRd("Frieght_Amount").ToString
    '                Me.LblNetInvoiceValue.Text = SQLRd("total_amount").ToString
    '                Me.txtDiscountAmt.Text = SQLRd("Discount_Amount").ToString
    '                If SQLRd("Discount_Type") = True And Val(Me.txtDiscountAmt.Text) > 0 Then
    '                    OptDiscountPercentage.Checked = True
    '                    OptDiscountValue.Checked = False
    '                ElseIf SQLRd("Discount_Type") = False And Val(Me.txtDiscountAmt.Text) > 0 Then
    '                    OptDiscountPercentage.Checked = False
    '                    OptDiscountValue.Checked = True
    '                Else
    '                    OptDiscountPercentage.Checked = False
    '                    OptDiscountValue.Checked = False
    '                End If

    '                For intX = 0 To Me.CmbTransType.Items.Count - 1
    '                    If Mid(CmbTransType.Items(intX).ToString, 1, 1) = SQLRd("Transport_Type").ToString Then
    '                        CmbTransType.SelectedIndex = intX
    '                    End If
    '                Next
    '                Me.txtRemarks.Text = SQLRd("REMARKS").ToString
    '                Me.lblBasicExciseAndCess.Text = SQLRd("BasicExciseAndCessValue").ToString
    '                Me.lblEcessValue.Text = SQLRd("ECESS_Amount").ToString
    '                Me.lblHCessValue.Text = SQLRd("SECESS_AMOUNT").ToString
    '                Me.lblSaltax_Per.Text = SQLRd("SalesTax_Per").ToString
    '                Me.lblECSStax_Per.Text = SQLRd("ECESS_Per").ToString
    '                Me.lblSECSStax_Per.Text = SQLRd("SECESS_Per").ToString
    '                Me.txtVehNo.Text = SQLRd("Vehicle_No").ToString
    '                Me.txtRefNo.Text = SQLRd("Cust_Ref").ToString
    '                Me.txtAmendNo.Text = SQLRd("Amendment_No").ToString
    '                Me.lblCustCodeDes.Text = SQLRd("Cust_Name").ToString
    '                Me.txtCreditTerms.Text = SQLRd("Payment_Terms").ToString
    '                Me.lblCreditTermDesc.Text = SQLRd("PAYMENT_TERMS_DESC").ToString
    '                Me.txtCustCode.Text = SQLRd("Account_Code").ToString
    '                Me.txtCarrServices.Text = SQLRd("Carriage_Name").ToString
    '                Me.txtAddVAT.Text = SQLRd("ADDVAT_Type").ToString
    '                Me.lblAddVAT.Text = SQLRd("AddVat_Per").ToString
    '                Me.lblAddVATValue.Text = SQLRd("ADDVAT_Amount").ToString
    '                blnBillFlag = SQLRd("BILL_FLAG").ToString
    '                lblRoundOff.Text = Val(SQLRd("TotalInvoiceAmtRoundOff_diff").ToString).ToString

    '                If SQLRd("Cancel_Flag").ToString.ToUpper = "TRUE" Then
    '                    lblCancelledInvoice.Visible = True
    '                Else
    '                    lblCancelledInvoice.Visible = False
    '                End If


    '            End If
    '            If SQLRd.IsClosed = False Then SQLRd.Close()

    '            Dim strItemCode As String
    '            Sqlcmd.CommandText = "SELECT TotalExciseAmount,From_Box,To_Box,Item_Code,Rate,Sales_Tax,Excise_Tax,Basic_Amount,Accessible_amount,Discount_amt,Sales_Quantity,BinQuantity,Cust_Item_Code, isnull(ADD_EXCISE_AMOUNT, 0) ADD_EXCISE_AMOUNT FROM SALES_DTL WHERE UNIT_CODE='" + gstrUNITID + "'  AND DOC_NO=" + Val(strInvoiceNo).ToString
    '            SQLRd = Sqlcmd.ExecuteReader
    '            If SQLRd.HasRows Then
    '                SQLRd.Read()
    '                'addRowAtEnterKeyPress(0)
    '                SpChEntry.MaxRows = 0
    '                SpChEntry.MaxRows = SpChEntry.MaxRows + 1
    '                ChangeCellTypeStaticText()
    '                SpChEntry.Row = SpChEntry.MaxRows

    '                Me.lblExciseValue.Text = SQLRd("TotalExciseAmount").ToString
    '                Me.lblAEDValue.Text = SQLRd("ADD_EXCISE_AMOUNT").ToString
    '                SpChEntry.Col = EnumInv.FROMBOX
    '                SpChEntry.Text = SQLRd("From_Box").ToString

    '                SpChEntry.Col = EnumInv.TOBOX
    '                SpChEntry.Text = SQLRd("To_Box").ToString

    '                SpChEntry.Col = EnumInv.CUMULATIVEBOXES
    '                SpChEntry.Text = SQLRd("To_Box").ToString

    '                SpChEntry.Col = EnumInv.CUMULATIVEBOXES : SpChEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : SpChEntry.TypeFloatDecimalPlaces = 0 : SpChEntry.TypeFloatMin = CDbl("0.00") : SpChEntry.TypeFloatMax = CDbl("99999999999999.99") : SpChEntry.Lock = True
    '                SpChEntry.Col = EnumInv.FROMBOX : SpChEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : SpChEntry.TypeFloatDecimalPlaces = 0 : SpChEntry.TypeFloatMin = CDbl("0.00") : SpChEntry.TypeFloatMax = CDbl("999999.99") : SpChEntry.Lock = True
    '                SpChEntry.Col = EnumInv.TOBOX : SpChEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : SpChEntry.TypeFloatDecimalPlaces = 0 : SpChEntry.TypeFloatMin = CDbl("0.00") : SpChEntry.TypeFloatMax = CDbl("999999.99") : SpChEntry.Lock = True


    '                SpChEntry.Col = EnumInv.ENUMITEMCODE
    '                SpChEntry.Text = SQLRd("Item_Code").ToString
    '                strItemCode = SQLRd("Item_Code").ToString

    '                SpChEntry.Col = EnumInv.RATE_PERUNIT
    '                SpChEntry.Text = SQLRd("Rate").ToString

    '                Me.lblSaltax_Per.Text = SQLRd("Sales_Tax").ToString
    '                Me.lblExciseValue.Text = SQLRd("Excise_Tax").ToString
    '                Me.lblBasicValue.Text = SQLRd("Basic_Amount").ToString
    '                Me.dblBasicValue = SQLRd("Basic_Amount").ToString
    '                Me.lblAssValue.Text = SQLRd("Accessible_amount").ToString
    '                Me.txtDiscountAmt.Text = SQLRd("Discount_amt").ToString

    '                SpChEntry.Col = EnumInv.ENUMQUANTITY
    '                SpChEntry.Text = SQLRd("Sales_Quantity").ToString


    '                SpChEntry.Col = EnumInv.BINQTY
    '                SpChEntry.Text = SQLRd("BinQuantity").ToString

    '                SpChEntry.Col = EnumInv.CUSTPARTNO
    '                SpChEntry.Text = SQLRd("Cust_Item_Code").ToString

    '                CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
    '                CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
    '                CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
    '                CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
    '                CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
    '                CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = False


    '                'If blnBillFlag = True Then ' IF INVOICE IS LOCKED
    '                '    'CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
    '                '    'CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
    '                'End If

    '            End If

    '            If SQLRd.IsClosed = False Then SQLRd.Close()
    '            Sqlcmd.CommandText = "DELETE FROM TMP_TRADING_INV_GRINS WHERE UNIT_CODE='" + gstrUNITID + "' AND IPADDRESS='" + gstrIpaddressWinSck + "'"
    '            Sqlcmd.ExecuteNonQuery()
    '            Sqlcmd.CommandText = "INSERT INTO TMP_TRADING_INV_GRINS  (GRINNO,SLNO,GRIN_PAGE_NO,ITEM_CODE,SALESQTY,GRINQTY,REMQTY,KNOCKOFFQTY,PERPIECEEXCISE,IPADDRESS,UNIT_CODE,PerPieceAED) SELECT GRIN_NO,SLNO,GRIN_PAGE_NO,ITEM_CODE,SALESQTY,GRINQTY,REMQTY,KNOCKOFFQTY,PERPIECEEXCISE,'" + gstrIpaddressWinSck + "',UNIT_CODE, PerPieceAED FROM SALES_TRADING_GRIN_DTL WHERE UNIT_CODE='" + gstrUNITID + "'  AND DOC_NO=" + Val(strInvoiceNo).ToString
    '            Sqlcmd.ExecuteNonQuery()

    '            Sqlcmd.CommandText = "UPDATE A SET GRINDATE=B.GRN_DATE FROM TMP_TRADING_INV_GRINS A ,GRN_HDR B WHERE A.UNIT_CODE =B.UNIT_CODE AND A.GRINNO=B.DOC_NO  AND A.UNIT_CODE ='" + gstrUNITID + "' AND A.IPADDRESS ='" + gstrIpaddressWinSck + "'"
    '            Sqlcmd.ExecuteNonQuery()
    '            GetItemDescription()


    '        Catch ex As Exception
    '            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '        Finally
    '            If IsNothing(Sqlcmd) = False Then
    '                If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
    '                Sqlcmd.Connection.Dispose()
    '                Sqlcmd.Dispose()
    '            End If
    '        End Try
    '    End Sub
    '    Public Function LockTradingInvoice(ByVal strInvoiceNo As String) As Boolean
    '        'ITEMBAL_MST            -- QUANTITY IS REDUCED FROM THIS TABLE WHILE LOKING THE INVOICE WHEN UPDATE STOCK FLAG IS TRUE IN SALESCONF.
    '        'CUST_ORD_DTL           -- QUANTITY IS KNOCKED OFF FROM THIS TABLE WHILE LOKING THE INVOICE WHEN UPDATE OP FLAG IS TRUE IN SALESCONF.
    '        'GRN_DTL                -- QUANTITY IS KNOCKED OFF FROM THIS TABLE WHILE LOKING THE INVOICE WHEN INVOICETYPE='NORMAL INVOICE' AND SUBTYPE='TRADING GOODS'
    '        'DAILYMKTSCHEDULE       -- QUANTITY IS KNOCKED OFF FROM THIS TABLE WHILE CREATING TRADING INVOICE.

    '        Dim Sqlcmd As New SqlCommand
    '        Dim strSql As String
    '        Dim ObjVal As Object = Nothing
    '        Dim SqlTrans As SqlTransaction
    '        Dim IsTrans As Boolean = False
    '        Dim SQLRd As SqlDataReader
    '        Dim builder As New StringBuilder

    '        Dim strDespatchQty As String = String.Empty
    '        Dim strCustDrgno As String = String.Empty
    '        Dim strInternalCode As String = String.Empty
    '        Try
    '            LockTradingInvoice = False
    '            Sqlcmd.CommandTimeout = 0

    '            Sqlcmd.Connection = SqlConnectionclass.GetConnection
    '            SqlTrans = Sqlcmd.Connection.BeginTransaction(System.Data.IsolationLevel.Serializable)
    '            Sqlcmd.Transaction = SqlTrans
    '            Sqlcmd.CommandType = CommandType.Text

    '            builder.Remove(0, builder.ToString.Length)
    '            builder.AppendLine("SELECT ITEM_CODE,CUST_ITEM_CODE,SALES_QUANTITY FROM SALES_DTL WHERE UNIT_CODE='" + gstrUNITID + "'  AND DOC_NO=" + Val(strInvoiceNo).ToString)
    '            Sqlcmd.CommandText = builder.ToString
    '            SQLRd = Sqlcmd.ExecuteReader
    '            If SQLRd.HasRows Then
    '                SQLRd.Read()
    '                strInternalCode = SQLRd("ITEM_CODE").ToString
    '                strCustDrgno = SQLRd("CUST_ITEM_CODE").ToString
    '                strDespatchQty = SQLRd("SALES_QUANTITY").ToString
    '            End If
    '            If SQLRd.IsClosed = False Then SQLRd.Close()

    '            builder.Remove(0, builder.ToString.Length)
    '            builder.AppendLine("UPDATE ITEMBAL_MST SET CUR_BAL=CUR_BAL-" + Val(strDespatchQty).ToString + " WHERE ITEM_CODE='" + strInternalCode + "' AND UNIT_CODE='" + gstrUNITID + "' ")
    '            Sqlcmd.CommandText = builder.ToString
    '            Sqlcmd.Parameters.Clear()
    '            Sqlcmd.ExecuteNonQuery()


    '            ' ===== NEW
    '            Sqlcmd.CommandType = CommandType.StoredProcedure
    '            Sqlcmd.CommandText = "TRADING_UPDATE_DISPATCH_IN_GRIN"
    '            Sqlcmd.Parameters.Clear()
    '            Sqlcmd.Parameters.Add("@UNITCODE", SqlDbType.VarChar).Value = gstrUNITID
    '            Sqlcmd.Parameters.Add("@INVOICE_NO", SqlDbType.VarChar).Value = Val(strInvoiceNo).ToString
    '            Sqlcmd.ExecuteNonQuery()
    '            ' ===== NEW


    '            Sqlcmd.CommandType = CommandType.Text
    '            builder.Remove(0, builder.ToString.Length)
    '            builder.AppendLine("UPDATE CUST_ORD_DTL SET DESPATCH_QTY = DESPATCH_QTY + " + Val(strDespatchQty).ToString)
    '            builder.AppendLine("WHERE UNIT_CODE='" + gstrUNITID + "'  AND ACCOUNT_CODE ='" + Me.txtCustCode.Text.Trim + "' AND CUST_DRGNO = '" + strCustDrgno + "'  AND CUST_REF = '" + Me.txtRefNo.Text.Trim + "'  AND AMENDMENT_NO = '" + Me.txtAmendNo.Text.Trim + "' AND ACTIVE_FLAG ='A'")
    '            Sqlcmd.CommandText = builder.ToString
    '            Sqlcmd.Parameters.Clear()
    '            Sqlcmd.ExecuteNonQuery()


    '            builder.Remove(0, builder.ToString.Length)
    '            builder.AppendLine("UPDATE SALESCHALLAN_DTL SET BILL_FLAG=1 WHERE UNIT_CODE='" + gstrUNITID + "'  AND DOC_NO=" + Val(strInvoiceNo).ToString)
    '            Sqlcmd.CommandText = builder.ToString
    '            Sqlcmd.Parameters.Clear()
    '            Sqlcmd.ExecuteNonQuery()

    '            SqlTrans.Commit()
    '            IsTrans = False
    '            LockTradingInvoice = True
    '        Catch ex As Exception
    '            If IsTrans = True Then SqlTrans.Rollback()
    '            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '        Finally
    '            If IsTrans = True Then SqlTrans.Rollback()
    '            If IsNothing(Sqlcmd.Connection.State) = False Then
    '                If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
    '            End If
    '            If IsNothing(Sqlcmd) = False Then
    '                Sqlcmd.Connection.Dispose()
    '                Sqlcmd.Dispose()
    '            End If
    '            If IsNothing(SqlTrans) = False Then SqlTrans.Dispose()
    '        End Try
    '    End Function
    '    Public Function CancelTradingInvoice(ByVal strInvoiceNo As String) As Boolean
    '        Dim Sqlcmd As New SqlCommand
    '        Dim ObjVal As Object = Nothing
    '        Dim SqlTrans As SqlTransaction
    '        Dim IsTrans As Boolean = False
    '        Dim SQLRd As SqlDataReader
    '        Dim builder As New StringBuilder

    '        Dim strDespatchQty As String = String.Empty
    '        Dim strCustDrgno As String = String.Empty
    '        Dim strInternalCode As String = String.Empty
    '        Try
    '            CancelTradingInvoice = False
    '            Sqlcmd.CommandTimeout = 0

    '            Sqlcmd.Connection = SqlConnectionclass.GetConnection
    '            SqlTrans = Sqlcmd.Connection.BeginTransaction(System.Data.IsolationLevel.Serializable)
    '            Sqlcmd.Transaction = SqlTrans
    '            Sqlcmd.CommandType = CommandType.Text


    '            builder.Remove(0, builder.ToString.Length)
    '            builder.AppendLine("SELECT ITEM_CODE,CUST_ITEM_CODE,SALES_QUANTITY FROM SALES_DTL WHERE UNIT_CODE='" + gstrUNITID + "'  AND DOC_NO=" + Val(strInvoiceNo).ToString)
    '            Sqlcmd.CommandText = builder.ToString
    '            SQLRd = Sqlcmd.ExecuteReader
    '            If SQLRd.HasRows Then
    '                SQLRd.Read()
    '                strInternalCode = SQLRd("ITEM_CODE").ToString
    '                strCustDrgno = SQLRd("CUST_ITEM_CODE").ToString
    '                strDespatchQty = SQLRd("SALES_QUANTITY").ToString
    '            End If
    '            If SQLRd.IsClosed = False Then SQLRd.Close()


    '            builder.Remove(0, builder.ToString.Length)
    '            builder.AppendLine("UPDATE CUST_ORD_DTL SET DESPATCH_QTY = DESPATCH_QTY - " + Val(strDespatchQty).ToString)
    '            builder.AppendLine("WHERE UNIT_CODE='" + gstrUNITID + "'  AND ACCOUNT_CODE ='" + Me.txtCustCode.Text.Trim + "' AND CUST_DRGNO = '" + strCustDrgno + "'  AND CUST_REF = '" + Me.txtRefNo.Text.Trim + "'  AND AMENDMENT_NO = '" + Me.txtAmendNo.Text.Trim + "' AND ACTIVE_FLAG ='A'")
    '            Sqlcmd.CommandText = builder.ToString
    '            Sqlcmd.Parameters.Clear()
    '            Sqlcmd.ExecuteNonQuery()

    '            Dim strRetMessage As String
    '            If UpdateForSchedules("-", Sqlcmd, strInternalCode, strCustDrgno, strDespatchQty, strRetMessage) = False Then
    '                IsTrans = False
    '                SqlTrans.Rollback()
    '                MsgBox(strRetMessage, MsgBoxStyle.Information, ResolveResString(100))
    '                Exit Function
    '            End If

    '            Sqlcmd.CommandType = CommandType.Text
    '            Sqlcmd.Parameters.Clear()

    '            builder.Remove(0, builder.ToString.Length)
    '            builder.AppendLine("UPDATE ITEMBAL_MST SET CUR_BAL=CUR_BAL+" + Val(strDespatchQty).ToString + " WHERE ITEM_CODE='" + strInternalCode + "' AND UNIT_CODE='" + gstrUNITID + "' ")
    '            Sqlcmd.CommandText = builder.ToString
    '            Sqlcmd.Parameters.Clear()
    '            Sqlcmd.ExecuteNonQuery()

    '            builder.Remove(0, builder.ToString.Length)
    '            builder.AppendLine("UPDATE A SET DESPATCH_QTY_TRADING=ISNULL(DESPATCH_QTY_TRADING,0)-B.KNOCKOFFQTY")
    '            builder.AppendLine("FROM GRN_DTL A,")
    '            builder.AppendLine("(	")
    '            builder.AppendLine("SELECT GRIN_NO,KNOCKOFFQTY,GRIN_DOC_TYPE,ITEM_CODE,UNIT_CODE ")
    '            builder.AppendLine("FROM SALES_TRADING_GRIN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND DOC_NO=" + Val(strInvoiceNo).ToString)
    '            builder.AppendLine(") B")
    '            builder.AppendLine("WHERE A.DOC_NO=B.GRIN_NO")
    '            builder.AppendLine("AND A.DOC_TYPE=B.GRIN_DOC_TYPE")
    '            builder.AppendLine("AND A.ITEM_CODE=B.ITEM_CODE  ")
    '            builder.AppendLine("AND A.UNIT_CODE=B.UNIT_CODE  ")
    '            builder.AppendLine("AND A.UNIT_CODE='" + gstrUNITID + "'")
    '            Sqlcmd.CommandText = builder.ToString
    '            Sqlcmd.Parameters.Clear()
    '            Sqlcmd.ExecuteNonQuery()

    '            '=== NEW
    '            builder.Remove(0, builder.ToString.Length)
    '            builder.AppendLine("UPDATE SALES_TRADING_GRIN_DTL SET GRIN_KNOCKOFF_SLNO=-1 WHERE UNIT_CODE='" + gstrUNITID + "' AND DOC_NO=" + Val(strInvoiceNo).ToString)
    '            Sqlcmd.CommandText = builder.ToString
    '            Sqlcmd.Parameters.Clear()
    '            Sqlcmd.ExecuteNonQuery()
    '            '=== NEW

    '            builder.Remove(0, builder.ToString.Length)
    '            builder.AppendLine("UPDATE SALESCHALLAN_DTL SET CANCEL_FLAG=1 WHERE UNIT_CODE='" + gstrUNITID + "'  AND DOC_NO=" + Val(strInvoiceNo).ToString)
    '            Sqlcmd.CommandText = builder.ToString
    '            Sqlcmd.Parameters.Clear()
    '            Sqlcmd.ExecuteNonQuery()

    '            SqlTrans.Commit()
    '            IsTrans = False
    '            CancelTradingInvoice = True
    '        Catch ex As Exception
    '            If IsTrans = True Then SqlTrans.Rollback()
    '            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '        Finally
    '            If IsTrans = True Then SqlTrans.Rollback()
    '            If IsNothing(Sqlcmd.Connection.State) = False Then
    '                If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
    '            End If
    '            If IsNothing(Sqlcmd) = False Then
    '                Sqlcmd.Connection.Dispose()
    '                Sqlcmd.Dispose()
    '            End If
    '            If IsNothing(SqlTrans) = False Then SqlTrans.Dispose()
    '        End Try
    '    End Function
    '    Private Function DeleteData() As Boolean
    '        Dim Sqlcmd As New SqlCommand
    '        Dim strSql As String
    '        Dim ObjVal As Object = Nothing
    '        Dim SqlTrans As SqlTransaction
    '        Dim IsTrans As Boolean = False
    '        Dim builder As New StringBuilder
    '        Dim strInvType As String = String.Empty
    '        Dim strInvSubType As String = String.Empty
    '        Dim strInvoiceNo As String = "0"
    '        Dim strDespatchQty As String
    '        Dim strCustDrgno As String
    '        Dim strInternalCode As String
    '        Try
    '            DeleteData = False
    '            If blnBillFlag = True Then
    '                MsgBox("Deletion Not Allowed For Locked Invoice(s).", MsgBoxStyle.Information, ResolveResString(100))
    '                Exit Function
    '            End If
    '            Sqlcmd.CommandTimeout = 0
    '            Sqlcmd.Connection = SqlConnectionclass.GetConnection
    '            Sqlcmd.CommandType = CommandType.Text

    '            Sqlcmd.CommandText = "SELECT TOP 1 1 FROM SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND DOC_NO>" + Val(Me.txtChallanNo.Text).ToString + " AND ISNULL(BILL_FLAG,0)=0"
    '            ObjVal = Sqlcmd.ExecuteScalar()
    '            If IsNothing(ObjVal) = True Then ObjVal = "0"
    '            If ObjVal = "1" Then
    '                MsgBox("Please Delete Invoice(s) Made After - " + Me.txtChallanNo.Text, MsgBoxStyle.Information, ResolveResString(100))
    '                Exit Function
    '            End If
    '            ObjVal = Nothing

    '            Sqlcmd.CommandText = "SELECT TOP 1 1 FROM SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND DOC_NO=" + Me.txtChallanNo.Text + " AND ISNULL(BILL_FLAG,0)=1"
    '            ObjVal = Sqlcmd.ExecuteScalar()
    '            If IsNothing(ObjVal) = True Then ObjVal = "0"
    '            If ObjVal = "1" Then
    '                MsgBox("This Invoice Has Been Locked.Can't Delete.", MsgBoxStyle.Information, ResolveResString(100))
    '                Exit Function
    '            End If
    '            ObjVal = Nothing

    '            SqlTrans = Sqlcmd.Connection.BeginTransaction(System.Data.IsolationLevel.Serializable)
    '            Sqlcmd.Transaction = SqlTrans
    '            Sqlcmd.CommandType = CommandType.Text

    '            SpChEntry.Col = EnumInv.ENUMQUANTITY
    '            strDespatchQty = SpChEntry.Text

    '            SpChEntry.Col = EnumInv.CUSTPARTNO
    '            strCustDrgno = SpChEntry.Text

    '            SpChEntry.Col = EnumInv.ENUMITEMCODE
    '            strInternalCode = SpChEntry.Text


    '            strInvoiceNo = Me.txtChallanNo.Text
    '            Sqlcmd.CommandText = "DELETE FROM SALES_TRADING_GRIN_DTL WHERE DOC_NO=" + strInvoiceNo + " AND UNIT_CODE='" + gstrUNITID + "' "
    '            Sqlcmd.ExecuteNonQuery()

    '            Sqlcmd.CommandText = "DELETE FROM SALES_DTL				 WHERE DOC_NO=" + strInvoiceNo + "  AND UNIT_CODE='" + gstrUNITID + "' "
    '            Sqlcmd.ExecuteNonQuery()

    '            Sqlcmd.CommandText = "DELETE FROM SALESCHALLAN_DTL		 WHERE DOC_NO=" + strInvoiceNo + "  AND UNIT_CODE='" + gstrUNITID + "' "
    '            Sqlcmd.ExecuteNonQuery()

    '            'builder.Remove(0, builder.ToString.Length)
    '            'builder.AppendLine("UPDATE CUST_ORD_DTL SET DESPATCH_QTY = DESPATCH_QTY - " + Val(strDespatchQty).ToString)
    '            'builder.AppendLine("WHERE UNIT_CODE='" + gstrUNITID + "'  AND ACCOUNT_CODE ='" + Me.txtCustCode.Text.Trim + "' AND CUST_DRGNO = '" + strCustDrgno + "'  AND CUST_REF = '" + Me.txtRefNo.Text.Trim + "'  AND AMENDMENT_NO = '" + Me.txtAmendNo.Text.Trim + "' AND ACTIVE_FLAG ='A'")
    '            'Sqlcmd.CommandText = builder.ToString
    '            'Sqlcmd.Parameters.Clear()
    '            'Sqlcmd.ExecuteNonQuery()

    '            Dim strRetMessage As String
    '            If UpdateForSchedules("-", Sqlcmd, strInternalCode, strCustDrgno, strDespatchQty, strRetMessage) = False Then
    '                IsTrans = False
    '                SqlTrans.Rollback()
    '                MsgBox(strRetMessage, MsgBoxStyle.Information, ResolveResString(100))
    '                Exit Function
    '            End If

    '            SqlTrans.Commit()
    '            IsTrans = False
    '            DeleteData = True
    '        Catch ex As Exception
    '            If IsTrans = True Then SqlTrans.Rollback()
    '            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '        Finally
    '            If IsTrans = True Then SqlTrans.Rollback()
    '            If IsNothing(Sqlcmd.Connection) = False Then
    '                If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
    '                Sqlcmd.Connection.Dispose()
    '                Sqlcmd.Dispose()
    '            End If
    '            If IsNothing(SqlTrans) = False Then SqlTrans.Dispose()
    '        End Try
    '    End Function
    '    Private Function GetAvailableStockAndDespatchQty(ByRef SQLCmd As SqlCommand, ByVal IsStockOrDespatch As Boolean, ByVal strItemCode As String) As Double
    '        'IsStockOrDespatch=True for Stock , False For Despatch
    '        Dim ObjVal As Object
    '        Dim builder As New StringBuilder
    '        SQLCmd.CommandType = CommandType.Text
    '        Try
    '            If IsStockOrDespatch = True Then
    '                builder.Remove(0, builder.ToString.Length)
    '                builder.AppendLine("SELECT 'AVAILABLEQTY'=CUR_BAL-")
    '                builder.AppendLine("ISNULL((")
    '                builder.AppendLine("SELECT SUM(B.SALES_QUANTITY) FROM SALESCHALLAN_DTL A,SALES_DTL B")
    '                builder.AppendLine("WHERE A.UNIT_CODE=B.UNIT_CODE")
    '                builder.AppendLine("AND   A.DOC_NO=B.DOC_NO")
    '                builder.AppendLine("AND   A.UNIT_CODE='" + gstrUNITID + "'  AND B.ITEM_CODE='" + strItemCode + "'")
    '                builder.AppendLine("AND   ISNULL(A.BILL_FLAG,0)=0 AND ISNULL(A.CANCEL_FLAG,0)=0")
    '                builder.AppendLine("),0) ")
    '                builder.AppendLine("FROM ITEMBAL_MST ")
    '                builder.AppendLine("WHERE UNIT_CODE='" + gstrUNITID + "'  AND ITEM_CODE='" + strItemCode + "' AND LOCATION_CODE='01T1'")

    '                SQLCmd.CommandText = builder.ToString
    '                ObjVal = SQLCmd.ExecuteScalar
    '                If IsNothing(ObjVal) = True Then ObjVal = "0"
    '                GetAvailableStockAndDespatchQty = Val(ObjVal.ToString)
    '            Else
    '                SQLCmd.CommandText = "SELECT ISNULL(DESPATCH_QTY,0) AS RESTQTY FROM CUST_ORD_DTL WHERE UNIT_CODE='" + gstrUNITID + "'  AND ACCOUNT_CODE='" + Me.txtCustCode.Text.Trim + "' AND CUST_REF='" + Me.txtRefNo.Text.Trim + "' AND AMENDMENT_NO='" + Me.txtAmendNo.Text.Trim + "' AND ITEM_CODE='" + strItemCode + "'"
    '                ObjVal = SQLCmd.ExecuteScalar
    '                If IsNothing(ObjVal) = True Then ObjVal = "0"
    '                GetAvailableStockAndDespatchQty = Val(ObjVal.ToString)
    '            End If
    '        Catch EX As Exception
    '            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
    '            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
    '        Finally

    '        End Try
    '    End Function
    '    Private Sub GetItemDescription()
    '        Dim strSql As String
    '        Dim Sqlcmd As New SqlCommand
    '        Dim SQLCon As SqlConnection
    '        Dim ObjVal As Object

    '        SQLCon = SqlConnectionclass.GetConnection()
    '        Sqlcmd.Connection = SQLCon
    '        Sqlcmd.CommandType = CommandType.Text
    '        Try


    '            Sqlcmd.CommandText = "SELECT UNT_UNITNAME FROM GEN_UNITMASTER WHERE UNT_CODEID='" + Me.txtLocationCode.Text + "'"
    '            ObjVal = Sqlcmd.ExecuteScalar()
    '            If IsNothing(ObjVal) = True Then ObjVal = String.Empty
    '            Me.lblLocCodeDes.Text = ObjVal.ToString


    '            Me.SpChEntry.Row = Me.SpChEntry.MaxRows
    '            Me.SpChEntry.Col = EnumInv.ENUMITEMCODE
    '            Sqlcmd.CommandText = "SELECT DESCRIPTION FROM ITEM_MST WHERE ITEM_CODE='" + Me.SpChEntry.Text.Trim + "' AND UNIT_CODE='" + gstrUNITID + "' "
    '            ObjVal = Sqlcmd.ExecuteScalar()
    '            If IsNothing(ObjVal) = True Then ObjVal = String.Empty
    '            Me.lblInternalPartDesc.Text = ObjVal.ToString

    '            Sqlcmd.CommandText = "SELECT DRG_DESC FROM CUSTITEM_MST WHERE UNIT_CODE='" + gstrUNITID + "'  AND ACCOUNT_CODE='" + Me.txtCustCode.Text + "' AND ITEM_CODE='" + SpChEntry.Text + "'"
    '            ObjVal = Sqlcmd.ExecuteScalar()
    '            If IsNothing(ObjVal) = True Then ObjVal = String.Empty
    '            Me.lblCustPartDesc.Text = ObjVal.ToString
    '            Me.lblCurrentStock.Text = GetAvailableStockAndDespatchQty(Sqlcmd, True, Me.SpChEntry.Text.Trim)
    '            Me.lblDespetchQty.Text = GetAvailableStockAndDespatchQty(Sqlcmd, False, Me.SpChEntry.Text.Trim)
    '        Catch EX As Exception
    '            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
    '            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
    '        Finally
    '            If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
    '            If SQLCon.State = ConnectionState.Open Then SQLCon.Close()
    '            Sqlcmd.Connection.Dispose()
    '            Sqlcmd.Dispose()
    '            SQLCon.Dispose()
    '        End Try
    '    End Sub
    '    Private Function UpdateForSchedules(ByVal pstrUpdType As String, ByRef SQLCmd As SqlCommand, ByVal Item_Code As String, ByVal CustDrawingNo As String, ByVal Quantity As String, ByRef ReturnMessage As String) As Boolean
    '        On Error GoTo ErrHandler
    '        Dim strsql As String
    '        Dim intCtr As Integer
    '        Dim strMSG As String
    '        Dim strYYYYmm As String
    '        Dim curQty As Decimal
    '        Dim varItemCode As Object
    '        Dim varDrgNo As Object
    '        Dim varItemQty As Object

    '        UpdateForSchedules = True
    '        strYYYYmm = Me.dtpDateDesc.Value.ToString("yyyy")  'Year(ConvertToDate(lblDateDes.Text)) & VB.Right("0" & Month(ConvertToDate(lblDateDes.Text)), 2)
    '        With SpChEntry

    '            varItemCode = Item_Code
    '            varDrgNo = CustDrawingNo
    '            varItemQty = Quantity

    '            SQLCmd.CommandType = CommandType.StoredProcedure
    '            SQLCmd.CommandText = "MKT_SCHEDULE_KNOCKOFF_NORTH"
    '            SQLCmd.Parameters.Add("@UNITCODE", SqlDbType.VarChar).Value = gstrUNITID
    '            SQLCmd.Parameters.Add("@CUSTOMER_CODE", SqlDbType.VarChar).Value = Trim(txtCustCode.Text)
    '            SQLCmd.Parameters.Add("@ITEM_CODE", SqlDbType.VarChar).Value = varItemCode
    '            SQLCmd.Parameters.Add("@CUSTDRG_NO", SqlDbType.VarChar).Value = varDrgNo
    '            SQLCmd.Parameters.Add("@FLAG", SqlDbType.VarChar).Value = pstrUpdType
    '            SQLCmd.Parameters.Add("@SCH_TYPE", SqlDbType.VarChar).Value = "D"
    '            SQLCmd.Parameters.Add("@YYYYMM", SqlDbType.VarChar).Value = strYYYYmm
    '            SQLCmd.Parameters.Add("@REQ_QTY", SqlDbType.Decimal).Value = Val(varItemQty)
    '            SQLCmd.Parameters.Add("@DATE", SqlDbType.VarChar).Value = getDateForDB(Me.dtpDateDesc.Value.ToString("dd/MM/yyyy"))
    '            SQLCmd.Parameters.Add("@MSG", SqlDbType.VarChar).Value = String.Empty
    '            SQLCmd.Parameters.Add("@ERR", SqlDbType.VarChar).Value = String.Empty
    '            SQLCmd.Parameters(9).Direction = ParameterDirection.Output
    '            SQLCmd.Parameters(10).Direction = ParameterDirection.Output
    '            SQLCmd.ExecuteNonQuery()
    '            If Len(SQLCmd.Parameters(9).Value) > 0 Then
    '                ReturnMessage = SQLCmd.Parameters(9).Value
    '                UpdateForSchedules = False
    '                Exit Function
    '            End If
    '            If Len(SQLCmd.Parameters(10).Value) > 0 Then
    '                ReturnMessage = SQLCmd.Parameters(10).Value
    '                UpdateForSchedules = False
    '                Exit Function
    '            End If
    '        End With
    '        Exit Function
    'ErrHandler:
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '    End Function
    '       '    Private Function CalculatePackingValue(ByVal pintRowNo As Short, ByVal blnRoundoff As Boolean, ByRef grid As AxFPSpreadADO.AxfpSpread) As Double
    '        Dim strPkg_Type As String
    '        Dim ldblPkg_Per As Double
    '        Dim ldblRate As Double
    '        Dim lintQty As Double
    '        Dim rsTaxRate As ClsResultSetDB
    '        Dim intPackingRoundoff_Decimal As Short
    '        On Error GoTo ErrHandler
    '        With grid

    '            .Row = pintRowNo

    '            .Col = EnumInv.RATE_PERUNIT
    '            ldblRate = 0 'Val(.Text) / Val(ctlPerValue.Text)

    '            .Col = EnumInv.PACKING
    '            strPkg_Type = Trim(.Text)

    '            .Col = EnumInv.ENUMQUANTITY
    '            lintQty = Val(.Text)
    '            intPackingRoundoff_Decimal = Val(Find_Value("select PackingRoundoff_Decimal from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'"))
    '            rsTaxRate = New ClsResultSetDB
    '            rsTaxRate.GetResult("Select Txrt_Rate_no,TxRt_Percentage from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  Tx_TaxeID = 'PKT' and Txrt_Rate_no = '" & Trim(strPkg_Type) & "'")
    '            If rsTaxRate.GetNoRows > 0 Then

    '                ldblPkg_Per = rsTaxRate.GetValue("TxRt_Percentage")
    '            Else
    '                ldblPkg_Per = 0
    '            End If
    '            rsTaxRate.ResultSetClose()
    '            If blnRoundoff = True Then
    '                CalculatePackingValue = System.Math.Round((ldblRate * lintQty) * ldblPkg_Per / 100, 0)
    '            Else
    '                CalculatePackingValue = System.Math.Round((ldblRate * lintQty) * ldblPkg_Per / 100, intPackingRoundoff_Decimal)
    '            End If
    '        End With
    '        Exit Function 'This is to avoid the execution of the error handler
    'ErrHandler:
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '        CalculatePackingValue = 0
    '    End Function
    '    Public Function ReturnCustomerLocation() As String
    '        On Error GoTo Errorhandler
    '        Dim rsObject As New ClsResultSetDB
    '        Call rsObject.GetResult("Select Cust_Location=isnull(Cust_Location,'') from Customer_mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Customer_Code = '" & Trim(Me.txtCustCode.Text) & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '        If rsObject.RowCount > 0 Then

    '            ReturnCustomerLocation = rsObject.GetValue("Cust_Location")
    '        Else
    '            ReturnCustomerLocation = ""
    '        End If

    '        rsObject.ResultSetClose()
    '        Exit Function
    'Errorhandler:
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '    End Function
    '    Private Sub CalculateTaxes()
    '        If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then

    '            If blnISBasicRoundOff = False Then
    '                dblBasicValue = Math.Round(dblBasicValue, intExciseRoundOffDecimal).ToString("############.00")
    '            End If

    '            lblExciseValue.Text = Math.Round(Val(lblExciseValue.Text)).ToString("############.00")
    '            lblAEDValue.Text = Math.Round(Val(lblAEDValue.Text)).ToString("############.00")

    '            lblEcessValue.Text = (Val(lblExciseValue.Text) * (Val(lblECSStax_Per.Text) / 100.0)).ToString("############.00")
    '            lblHCessValue.Text = (Val(lblExciseValue.Text) * (Val(lblSECSStax_Per.Text) / 100.0)).ToString("############.00")

    '            If blnECSSTax = False Then
    '                lblEcessValue.Text = Math.Round(Val(lblEcessValue.Text), intECSRoundOffDecimal).ToString("############.00")
    '                lblHCessValue.Text = Math.Round(Val(lblHCessValue.Text), intECSRoundOffDecimal).ToString("############.00")
    '            End If


    '            If blnISSalesTaxRoundOff = False Then
    '                lblSalesTaxValue.Text = Math.Round(Val(lblSalesTaxValue.Text), intSaleTaxRoundOffDecimal).ToString("############.00")
    '                lblAddVATValue.Text = Math.Round(Val(lblAddVATValue.Text), intSaleTaxRoundOffDecimal).ToString("############.00")
    '            End If

    '            Me.lblBasicExciseAndCess.Text = (dblBasicValue + Val(lblExciseValue.Text) + Val(lblAEDValue.Text) + Val(lblEcessValue.Text) + Val(lblHCessValue.Text)).ToString("############.00")

    '            If OptDiscountValue.Checked = True Then
    '                Me.lblAssValue.Text = Val(Me.lblBasicExciseAndCess.Text) - Val(txtDiscountAmt.Text)
    '            ElseIf OptDiscountPercentage.Checked = True Then
    '                Me.lblAssValue.Text = Val(Me.lblBasicExciseAndCess.Text) - (Val(Me.lblBasicExciseAndCess.Text) * Val(txtDiscountAmt.Text) / 100)
    '            Else
    '                Me.lblAssValue.Text = Me.lblBasicExciseAndCess.Text
    '            End If

    '            lblSalesTaxValue.Text = ((Val(Me.lblAssValue.Text)) * Val(lblSaltax_Per.Text) / 100.0).ToString("############.00")
    '            lblAddVATValue.Text = ((Val(Me.lblSalesTaxValue.Text)) * Val(lblAddVAT.Text) / 100.0).ToString("############.00")


    '            LblNetInvoiceValue.Text = (Val(Me.lblAssValue.Text) + Val(lblSalesTaxValue.Text) + Val(lblAddVATValue.Text) + Val(ctlInsurance.Text) + Val(txtFreight.Text)).ToString("###########.00")
    '            LblNetInvoiceValue.Tag = LblNetInvoiceValue.Text
    '            LblNetInvoiceValue.Text = Math.Round(Val(LblNetInvoiceValue.Text)).ToString("###########.00")
    '            lblRoundOff.Text = (Val(Me.LblNetInvoiceValue.Tag.ToString) - Val(Me.LblNetInvoiceValue.Text)).ToString("##########0.00")

    '        End If
    '    End Sub
    '    Private Sub IncludeDefaultTaxes()
    '        Dim strSql As String
    '        Dim Sqlcmd As New SqlCommand
    '        Dim Dr As SqlDataReader
    '        Dim SQLCon As SqlConnection
    '        Dim intX As Integer = 0
    '        SQLCon = SqlConnectionclass.GetConnection()
    '        Sqlcmd.Connection = SQLCon
    '        Sqlcmd.CommandType = CommandType.Text

    '        Try
    '            strSql = "SELECT SLNO,TXRT_RATE_NO,TXRT_PERCENTAGE "
    '            strSql = strSql + " FROM"
    '            strSql = strSql + " ("
    '            strSql = strSql + " SELECT 1 SLNO,TXRT_RATE_NO,TXRT_PERCENTAGE FROM GEN_TAXRATE "
    '            strSql = strSql + " WHERE UNIT_CODE='" + gstrUNITID + "' AND TX_TAXEID='ECT' AND TXRT_RATE_NO='ECT2' AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) )"
    '            strSql = strSql + " UNION ALL"
    '            strSql = strSql + " SELECT 2 SLNO,TXRT_RATE_NO,TXRT_PERCENTAGE FROM GEN_TAXRATE "
    '            strSql = strSql + " WHERE UNIT_CODE='" + gstrUNITID + "' AND (TX_TAXEID='ECSST') AND TXRT_RATE_NO='ECSST1' AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) "
    '            strSql = strSql + " OR (CAST(GETDATE() AS DATE) <= DEACTIVE_DATE))"
    '            strSql = strSql + " ) ABCD order by 1 "
    '            Sqlcmd.CommandText = strSql
    '            Dr = Sqlcmd.ExecuteReader
    '            If Dr.HasRows = True Then
    '                While Dr.Read
    '                    If intX = 0 Then
    '                        txtECSSTaxType.Text = Dr("TXRT_RATE_NO")
    '                        lblECSStax_Per.Text = Dr("TXRT_PERCENTAGE")
    '                    Else
    '                        txtSECSSTaxType.Text = Dr("TXRT_RATE_NO")
    '                        lblSECSStax_Per.Text = Dr("TXRT_PERCENTAGE")
    '                    End If
    '                    intX = intX + 1
    '                End While
    '            End If
    '            If Dr.IsClosed = False Then Dr.Close()
    '        Catch EX As Exception
    '            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
    '            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
    '        Finally
    '            If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
    '            If SQLCon.State = ConnectionState.Open Then SQLCon.Close()
    '            Sqlcmd.Connection.Dispose()
    '            Sqlcmd.Dispose()
    '            SQLCon.Dispose()
    '        End Try
    '    End Sub
    '    Private Function CheckExistanceOfFieldData(ByRef pstrFieldText As String, ByRef pstrColumnName As String, ByRef pstrTableName As String, Optional ByRef pstrCondition As String = "") As Boolean
    '        On Error GoTo ErrHandler
    '        CheckExistanceOfFieldData = False
    '        Dim strTableSql As String 'Declared To Make Select Query
    '        Dim rsExistData As ClsResultSetDB
    '        If Len(Trim(pstrCondition)) > 0 Then
    '            strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and " & pstrCondition
    '        Else
    '            strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "'"
    '        End If
    '        rsExistData = New ClsResultSetDB
    '        rsExistData.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '        If rsExistData.GetNoRows > 0 Then
    '            CheckExistanceOfFieldData = True
    '        Else
    '            CheckExistanceOfFieldData = False
    '        End If
    '        rsExistData.ResultSetClose()
    '        Exit Function
    'ErrHandler:  'The Error Handling Code Starts here
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    '    End Function
    '    Private Function GetTaxRate(ByRef pstrFieldText As String, ByRef pstrColumnName As String, ByRef pstrTableName As String, ByRef pstrFieldName_WhichValueRequire As String, Optional ByRef pstrCondition As String = "") As Double
    '        On Error GoTo ErrHandler
    '        GetTaxRate = 0
    '        Dim strTableSql As String 'Declared To Make Select Query
    '        Dim rsExistData As ClsResultSetDB
    '        If Len(Trim(pstrCondition)) > 0 Then
    '            strTableSql = "select " & Trim(pstrFieldName_WhichValueRequire) & " from " & Trim(pstrTableName) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and " & pstrCondition
    '        Else
    '            strTableSql = "select " & Trim(pstrFieldName_WhichValueRequire) & " from " & Trim(pstrTableName) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "'"
    '        End If
    '        rsExistData = New ClsResultSetDB
    '        rsExistData.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '        If rsExistData.GetNoRows > 0 Then

    '            GetTaxRate = rsExistData.GetValue(Trim(pstrFieldName_WhichValueRequire))
    '        Else
    '            GetTaxRate = 0
    '        End If
    '        rsExistData.ResultSetClose()
    '        Exit Function
    'ErrHandler:  'The Error Handling Code Starts here
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    '    End Function
    '    Public Function ReturnNoOfDecimals(ByRef pstrItemCode As String) As Short
    '        '***************************************************************************************
    '        'Name       :   ReturnNoOfDecimals
    '        'Type       :   Function
    '        'Author     :   Nisha Rai
    '        'Arguments  :   Item Code
    '        'Return     :   No of Decimals as intiger
    '        'Purpose    :   Fetches measure code of given item code and then according to decimal
    '        '               allowed fkag it returns No decimals allowed
    '        '***************************************************************************************
    '        Dim rsMeasurementUnit As ClsResultSetDB
    '        Dim rsNoOfDecimal As ClsResultSetDB
    '        Dim strMeasurementUnit As String
    '        Dim intNoofDecimals As Short
    '        On Error GoTo ErrHandler
    '        rsMeasurementUnit = New ClsResultSetDB
    '        rsMeasurementUnit.GetResult("Select Cons_Measure_Code,Pur_Measure_Code from Item_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  item_code = '" & pstrItemCode & "'")
    '        If rsMeasurementUnit.GetNoRows > 0 Then
    '            rsMeasurementUnit.MoveFirst()

    '            If UCase(CmbInvType.Text) = "REJECTION" Then
    '                strMeasurementUnit = rsMeasurementUnit.GetValue("Pur_Measure_Code")
    '            Else
    '                strMeasurementUnit = rsMeasurementUnit.GetValue("Cons_Measure_Code")
    '            End If
    '            rsNoOfDecimal = New ClsResultSetDB
    '            rsNoOfDecimal.GetResult("select Decimal_Allowed_Flag,NoOFDecimal from Measure_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Measure_Code = '" & strMeasurementUnit & "'")
    '            If rsNoOfDecimal.GetNoRows > 0 Then
    '                rsNoOfDecimal.MoveFirst()

    '                If rsNoOfDecimal.GetValue("Decimal_Allowed_Flag") = True Then

    '                    intNoofDecimals = Val(rsNoOfDecimal.GetValue("NoOFDecimal"))
    '                    If intNoofDecimals = 0 Then
    '                        intNoofDecimals = 2
    '                    End If
    '                    ReturnNoOfDecimals = intNoofDecimals
    '                Else
    '                    ReturnNoOfDecimals = 0
    '                End If
    '            End If
    '            rsNoOfDecimal.ResultSetClose()
    '        End If
    '        rsMeasurementUnit.ResultSetClose()
    '        Exit Function
    'ErrHandler:
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '    End Function
    '    Private Sub ChangeCellTypeStaticText()
    '        On Error GoTo ErrHandler
    '        Dim intRow As Short
    '        Dim intcol As Short
    '        Dim varItemCode As Object
    '        Dim blnQtyChkAccToMeasureCode As Boolean
    '        Dim rsSalesParameter As ClsResultSetDB
    '        Dim strMin As String
    '        Dim strMax As String
    '        Dim intDecimal As Short
    '        Dim intloopcounter1 As Short
    '        Dim blnTrfInvoiceWithSO As Boolean
    '        If mblnBatchTrack = True And Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
    '            If Me.SpChEntry.MaxRows > 0 Then
    '                ReDim mBatchData(Me.SpChEntry.MaxRows)
    '            End If
    '            For intloopcounter1 = 1 To Me.SpChEntry.MaxRows
    '                ReDim mBatchData(intloopcounter1).Batch_No(0)
    '                ReDim mBatchData(intloopcounter1).Batch_Date(0)
    '                ReDim mBatchData(intloopcounter1).Batch_Quantity(0)
    '            Next
    '        End If
    '        Dim strQry As String
    '        Dim rsSOReq As ClsResultSetDB
    '        Dim dblMaxLimit As Double
    '        Dim str_NewCurrencyCode As String
    '        Dim int_NoOfDecimal As Short
    '        Dim str_Min As String
    '        Dim str_Max As String
    '        Dim intLoopCounter As Short
    '        Dim rs_currencycode As ClsResultSetDB
    '        Dim strInvSubType As String
    '        Dim rsSOReq1 As ClsResultSetDB
    '        Dim strqry1 As String
    '        With Me.SpChEntry
    '            Select Case Me.CmdGrpChEnt.Mode
    '                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
    '                    blnTrfInvoiceWithSO = False

    '                    If (UCase(Trim(CmbInvType.Text)) = "NORMAL INVOICE") Or ((UCase(Trim(CmbInvType.Text)) = "TRANSFER INVOICE") And blnTrfInvoiceWithSO) Or (UCase(Trim(CmbInvType.Text)) = "EXPORT INVOICE") Or (UCase(Trim(CmbInvType.Text)) = "SERVICE INVOICE") Then

    '                        If UCase(Trim(CmbInvSubType.Text)) <> "SCRAP" Then
    '                            For intRow = 1 To .MaxRows
    '                                .Row = intRow
    '                                For intcol = 1 To .MaxCols
    '                                    .Col = intcol
    '                                    If intcol = EnumInv.ENUMQUANTITY Then
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                                        .Lock = True
    '                                    ElseIf intcol = EnumInv.BINQTY Then
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                                    ElseIf intcol = EnumInv.SelectGrin Then
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
    '                                        .TypeButtonPicture = My.Resources.ico111.ToBitmap
    '                                        .Text = "..."
    '                                    Else
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
    '                                    End If
    '                                Next intcol
    '                            Next intRow
    '                        Else
    '                            For intRow = 1 To .MaxRows
    '                                .Row = intRow
    '                                For intcol = 1 To .MaxCols
    '                                    .Col = intcol
    '                                    If intcol = EnumInv.ENUMQUANTITY Or intcol = EnumInv.FROMBOX Or intcol = EnumInv.TOBOX Then
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                                    ElseIf intcol = EnumInv.RATE_PERUNIT Then
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                                    ElseIf intcol = EnumInv.BATCHCOL And mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
    '                                    ElseIf intcol = EnumInv.BINQTY Then
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                                    ElseIf intcol = EnumInv.SelectGrin Then
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
    '                                        .TypeButtonPicture = My.Resources.ico111.ToBitmap
    '                                        .Text = "..."
    '                                    Else
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
    '                                    End If
    '                                Next intcol
    '                            Next intRow
    '                        End If
    '                    Else
    '                        For intRow = 1 To .MaxRows
    '                            .Row = intRow
    '                            For intcol = 1 To .MaxCols
    '                                .Col = intcol
    '                                If intcol = EnumInv.ENUMQUANTITY Or intcol = EnumInv.FROMBOX Or intcol = EnumInv.TOBOX Then
    '                                    If mblnRejTracking = True Then
    '                                        If intcol = EnumInv.ENUMQUANTITY And CmbInvType.Text = "REJECTION" Then
    '                                            SpChEntry.Col = EnumInv.ENUMQUANTITY
    '                                            SpChEntry.Row = intRow
    '                                            dblMaxLimit = .TypeFloatMax
    '                                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = dblMaxLimit
    '                                        Else
    '                                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                                        End If
    '                                    Else
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                                    End If
    '                                ElseIf intcol = EnumInv.RATE_PERUNIT Then
    '                                    rs_currencycode = New ClsResultSetDB
    '                                    rs_currencycode.GetResult("Select currency_code from cust_ord_hdr WHERE UNIT_CODE='" + gstrUNITID + "' AND  account_code = " & "'" & Me.txtCustCode.Text & "'" & "and cust_ref = " & "'" & Me.txtRefNo.Text & "'")
    '                                    str_NewCurrencyCode = rs_currencycode.GetValue("Currency_code")
    '                                    rs_currencycode.ResultSetClose()
    '                                    int_NoOfDecimal = ToGetDecimalPlaces(Trim(str_NewCurrencyCode))
    '                                    If int_NoOfDecimal < 2 Then
    '                                        int_NoOfDecimal = 2
    '                                    End If
    '                                    str_Min = "0." : str_Max = "99999999999999."
    '                                    .Col = EnumInv.RATE_PERUNIT : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = int_NoOfDecimal : .TypeFloatMin = CDbl(str_Min) : .TypeFloatMax = CDbl(str_Max)

    '                                ElseIf intcol = EnumInv.BATCHCOL And mblnBatchTrack = True And UCase(CmbInvType.Text) <> "REJECTION" Then
    '                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = "Batch Details"
    '                                ElseIf intcol = EnumInv.BINQTY Then
    '                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                                Else
    '                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
    '                                End If
    '                            Next intcol
    '                        Next intRow
    '                    End If
    '                Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
    '                    If Trim(strInvType) = "" Then
    '                        rsSalesParameter = New ClsResultSetDB
    '                        rsSalesParameter.GetResult("Select invoice_type from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no='" & Trim(txtChallanNo.Text) & "'")
    '                        If rsSalesParameter.GetNoRows > 0 Then
    '                            strInvType = rsSalesParameter.GetValue("invoice_type")
    '                        Else
    '                            strInvType = ""
    '                        End If
    '                        rsSalesParameter.ResultSetClose()
    '                    End If
    '                    rsSOReq1 = New ClsResultSetDB
    '                    strInvSubType = Find_Value("Select Sub_category from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no='" & Trim(txtChallanNo.Text) & "'")
    '                    strqry1 = "Select isnull(SORequired,0) as SORequired from saleConf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type ='" & strInvType & "' and Sub_Type='" & Trim(strInvSubType) & "' and  (fin_start_date <= getdate() and fin_end_date >= getdate())"
    '                    rsSOReq1.GetResult(strqry1)
    '                    If rsSOReq1.GetNoRows > 0 Then
    '                        blnTrfInvoiceWithSO = rsSOReq1.GetValue("SORequired")
    '                    Else
    '                        blnTrfInvoiceWithSO = False
    '                    End If
    '                    rsSOReq1.ResultSetClose()
    '                    If (UCase(strInvType) = "INV") Or ((UCase(strInvType) = "TRF") And blnTrfInvoiceWithSO) Or (UCase(strInvType) = "EXP") Or (UCase(strInvType) = "SRC") Then
    '                        If (UCase(strInvType) = "SRC") And mblnServiceInvoiceWithoutSO Then
    '                            For intRow = 1 To .MaxRows
    '                                .Row = intRow
    '                                For intcol = 1 To .MaxCols
    '                                    .Col = intcol
    '                                    If intcol = EnumInv.ENUMQUANTITY Or intcol = EnumInv.FROMBOX Or intcol = EnumInv.TOBOX Then
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                                    ElseIf intcol = EnumInv.BATCHCOL And mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
    '                                    ElseIf intcol = EnumInv.BINQTY Then
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                                    Else
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
    '                                    End If
    '                                    .Col = EnumInv.RATE_PERUNIT : .Lock = False
    '                                    .CtlEditMode = True : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                                    .Col = EnumInv.ENUMQUANTITY : .Lock = True

    '                                Next intcol
    '                            Next intRow
    '                        ElseIf (UCase(strInvSubType) <> "L") Then
    '                            For intRow = 1 To .MaxRows
    '                                .Row = intRow
    '                                For intcol = 1 To .MaxCols
    '                                    .Col = intcol
    '                                    If intcol = EnumInv.ENUMQUANTITY Or intcol = EnumInv.FROMBOX Or intcol = EnumInv.TOBOX Then
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat

    '                                    ElseIf intcol = EnumInv.BATCHCOL And mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
    '                                    ElseIf intcol = EnumInv.BINQTY Then
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                                    Else
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
    '                                    End If
    '                                Next intcol
    '                            Next intRow
    '                        Else
    '                            For intRow = 1 To .MaxRows
    '                                .Row = intRow
    '                                For intcol = 1 To .MaxCols
    '                                    .Col = intcol
    '                                    If intcol = EnumInv.ENUMQUANTITY Or intcol = EnumInv.FROMBOX Or intcol = EnumInv.TOBOX Then
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                                    ElseIf intcol = EnumInv.RATE_PERUNIT Then
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                                    ElseIf intcol = EnumInv.BATCHCOL And mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
    '                                    ElseIf intcol = EnumInv.BINQTY Then
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                                    Else
    '                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
    '                                    End If
    '                                Next intcol
    '                            Next intRow
    '                        End If
    '                    Else
    '                        For intRow = 1 To .MaxRows
    '                            .Row = intRow
    '                            For intcol = 1 To .MaxCols
    '                                .Col = intcol
    '                                If intcol = EnumInv.ENUMQUANTITY Or intcol = EnumInv.FROMBOX Or intcol = EnumInv.TOBOX Then
    '                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                                ElseIf intcol = EnumInv.RATE_PERUNIT Then
    '                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                                ElseIf intcol = EnumInv.BATCHCOL And mblnBatchTrack = True And UCase(CmbInvType.Text) <> "REJECTION" Then
    '                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
    '                                ElseIf intcol = EnumInv.BINQTY Then
    '                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                                Else
    '                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
    '                                End If
    '                            Next intcol
    '                        Next intRow
    '                    End If
    '                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
    '                    If mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
    '                        With Me.SpChEntry
    '                            .Enabled = True
    '                            SpChEntry.Row = 1 : SpChEntry.Row2 = SpChEntry.MaxRows : SpChEntry.Col = 0 : SpChEntry.Col2 = SpChEntry.MaxCols - 1
    '                            SpChEntry.BlockMode = True : SpChEntry.Lock = True : SpChEntry.BlockMode = False
    '                            SpChEntry.Row = 1 : SpChEntry.Row2 = SpChEntry.MaxRows : SpChEntry.Col = EnumInv.BATCHCOL : SpChEntry.Col2 = EnumInv.BATCHCOL
    '                            SpChEntry.BlockMode = True : SpChEntry.Lock = False : SpChEntry.BlockMode = False
    '                        End With
    '                    Else
    '                        For intRow = 1 To .MaxRows
    '                            SpChEntry.Row = intRow
    '                            For intcol = 1 To SpChEntry.MaxCols
    '                                SpChEntry.Col = intcol
    '                                If intcol = EnumInv.ENUMQUANTITY Then
    '                                    SpChEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                                    SpChEntry.Lock = True
    '                                ElseIf intcol = EnumInv.BINQTY Then
    '                                    SpChEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                                ElseIf intcol = EnumInv.SelectGrin Then
    '                                    SpChEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
    '                                    SpChEntry.TypeButtonPicture = My.Resources.ico111.ToBitmap
    '                                    SpChEntry.Text = "..."
    '                                Else
    '                                    SpChEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
    '                                End If
    '                            Next intcol
    '                        Next intRow
    '                    End If
    '            End Select
    '            rsSalesParameter = New ClsResultSetDB
    '            rsSalesParameter.GetResult("Select QtyChkAccToMeasureCode from Sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")
    '            If rsSalesParameter.GetNoRows > 0 Then
    '                If rsSalesParameter.GetValue("QtyChkAccToMeasureCode") = False Then
    '                    blnQtyChkAccToMeasureCode = False
    '                Else
    '                    blnQtyChkAccToMeasureCode = True
    '                End If
    '            End If
    '            rsSalesParameter.ResultSetClose()
    '            If blnQtyChkAccToMeasureCode = True Then
    '                For intRow = 1 To .MaxRows
    '                    varItemCode = Nothing
    '                    Call .GetText(EnumInv.ENUMITEMCODE, intRow, varItemCode)
    '                    If Len(Trim(varItemCode)) > 0 Then
    '                        intDecimal = ReturnNoOfDecimals(CStr(varItemCode))
    '                        strMin = "0." : strMax = "99999999999999."
    '                        For intloopcounter1 = 1 To intDecimal
    '                            strMin = strMin & "0"
    '                            strMax = strMax & "9"
    '                        Next
    '                        If intDecimal = 0 Then
    '                            strMin = "0" : strMax = "99999999999999"
    '                        End If
    '                        .Row = intRow : .Row2 = intRow : .Col = EnumInv.ENUMQUANTITY : .Col2 = EnumInv.ENUMQUANTITY : .BlockMode = True '.CellType = CellTypeFloat
    '                        .TypeFloatDecimalPlaces = intDecimal
    '                        .BlockMode = False
    '                    End If
    '                Next
    '            End If
    '        End With
    '        Exit Sub
    'ErrHandler:  'The Error Handling Code Starts here
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    '    End Sub
    '    Public Function Find_Value(ByRef strField As String) As String
    '        On Error GoTo ErrHandler
    '        Dim Rs As New ADODB.Recordset
    '        Rs = New ADODB.Recordset
    '        Rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    '        Rs.Open(strField, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
    '        If Rs.RecordCount > 0 Then
    '            If IsDBNull(Rs.Fields(0).Value) = False Then
    '                Find_Value = Rs.Fields(0).Value
    '            Else
    '                Find_Value = ""
    '            End If
    '        Else
    '            Find_Value = ""
    '        End If
    '        Rs.Close()
    '        Exit Function
    'ErrHandler:
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '    End Function
    '    Public Function GetDocumentDetail(ByRef strItem_code As String) As String
    '        On Error GoTo ErrHandler
    '        Dim strCompileString As String
    '        Dim vardoc_no As Object
    '        Dim varBatch_No As Object
    '        Dim varBatch_Date As Object
    '        Dim varBatchReq As Object
    '        Dim varQty As Object
    '        Dim intRow As Short
    '        Dim rsTmp As New ClsResultSetDB
    '        Dim strSql As String
    '        strSql = "Select REF_DOC_NO, Batch_No, Quantity from MKT_INVREJ_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_No=" & Trim(txtChallanNo.Text) & " and Item_code='" & strItem_code & "'"
    '        rsTmp.GetResult(strSql)
    '        strCompileString = ""
    '        If rsTmp.RowCount > 0 Then
    '            Do While Not rsTmp.EOFRecord
    '                strCompileString = strCompileString & rsTmp.GetValue("Ref_Doc_no") & "§" & rsTmp.GetValue("Batch_No") & "§§" & rsTmp.GetValue("Quantity") & "¶"
    '                rsTmp.MoveNext()
    '            Loop
    '        End If
    '        rsTmp.ResultSetClose()
    '        GetDocumentDetail = strCompileString
    '        Exit Function
    'ErrHandler:
    '        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '    End Function
    '    Sub AddMaxAllowedQuanity(ByRef intRow As Short)
    '        On Error GoTo ErrHandler
    '        Dim dblqty As Double
    '        Dim dblStock As Double
    '        Dim varItemCode As Object
    '        Dim varMaxQuanity As Object
    '        Dim strsaledtl As String
    '        varItemCode = Nothing
    '        SpChEntry.GetText(EnumInv.ENUMITEMCODE, intRow, varItemCode)
    '        If Find_Value("Select top 1 REJ_TYPE from MKT_INVREJ_DTL  WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_No=" & txtChallanNo.Text) = "1" Then
    '            strsaledtl = "select MaxAllowedQty = SUM( ((a.Rejected_Quantity + a.excess_po_quantity) - (isnull(a.Despatch_Quantity,0) + isnull(a.Inspected_Quantity,0) + isnull(a.RGP_Quantity,0)))) from grn_Dtl a, grn_hdr b Where "
    '            strsaledtl = strsaledtl & " a.Doc_type = b.Doc_type and a.unit_code=b.unit_code and a.unit_code='" + gstrUNITID + "' And a.Doc_No = b.Doc_No and "
    '            strsaledtl = strsaledtl & " a.From_Location = b.From_Location and a.From_Location ='01R1'"
    '            strsaledtl = strsaledtl & " and a.Rejected_quantity > 0 and b.Vendor_code = '" & txtCustCode.Text
    '            strsaledtl = strsaledtl & "' and a.Doc_No in (" & Trim(txtRefNo.Text) & ") and a.Item_code = '" & varItemCode & "' AND ISNULL(b.GRN_Cancelled,0) = 0"
    '            strsaledtl = strsaledtl & " Group by a.Item_code "
    '            dblqty = Val(Find_Value(strsaledtl))
    '        ElseIf Find_Value("Select top 1 REJ_TYPE from MKT_INVREJ_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_No=" & txtChallanNo.Text) = "2" Then
    '            strsaledtl = "Select   MaxAllowedQty = Sum(rejected_Quantity) from LRN_HDR as a " & " Inner Join LRN_DTL as b on a.doc_No=b.doc_no and a.unit_code=b.unit_code and a.unit_code='" + gstrUNITID + "' and a.Doc_Type=b.doc_Type and a.from_Location=b.from_location " & " Where Authorized_Code Is Not Null " & " and a.Doc_No IN (" & Trim(txtRefNo.Text) & ") and ITem_code = '" & varItemCode & "'" & " Group by B.Item_Code "
    '            dblqty = Val(Find_Value(strsaledtl))
    '        End If
    '        dblStock = Val(Find_Value("Select Cur_Bal From ItemBal_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Location_code='01J1' and Item_code='" & varItemCode & "'"))
    '        If dblStock < dblqty Then
    '            varMaxQuanity = dblStock
    '        Else
    '            varMaxQuanity = dblqty
    '        End If

    '        Exit Sub
    'ErrHandler:
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '    End Sub
    '    Private Function SetMaxLengthInSpread(ByRef pintDecimalSize As Short) As Object
    '        On Error GoTo ErrHandler
    '        Dim intRow As Short
    '        Dim strMin As String
    '        Dim strMax As String
    '        Dim intLoopCounter As Short
    '        Dim intDecimal As Short
    '        If pintDecimalSize < 2 Then
    '            pintDecimalSize = 2
    '        End If
    '        strMin = "0." : strMax = "99999999999999."
    '        For intLoopCounter = 1 To intDecimal
    '            strMin = strMin & "0"
    '            strMax = strMax & "9"
    '        Next
    '        With Me.SpChEntry
    '            For intRow = 1 To .MaxRows
    '                .Row = intRow
    '                .Col = EnumInv.ENUMITEMCODE : .TypeMaxEditLen = 16
    '                .Col = EnumInv.CUSTPARTNO : .TypeMaxEditLen = 30
    '                .Col = EnumInv.RATE_PERUNIT : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax)
    '                .Col = EnumInv.CUSTSUPPMAT_PERUNIT : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax)
    '                .Col = EnumInv.ENUMQUANTITY : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("99999999999999.99")
    '                .Col = EnumInv.PACKING : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit

    '                .CtlEditMode = False
    '                If CmbInvType.Text = "NORMAL INVOICE" Or CmbInvType.Text = "JOBWORK INVOICE" Or CmbInvType.Text = "EXPORT INVOICE" Or (CmbInvType.Text = "SERVICE INVOICE" And Not mblnServiceInvoiceWithoutSO) Then
    '                    If UCase(Trim(CmbInvSubType.Text)) <> "SCRAP" Then
    '                        .CtlEditMode = False
    '                    Else
    '                        .CtlEditMode = True
    '                    End If
    '                Else
    '                    .CtlEditMode = True
    '                End If
    '                .Col = EnumInv.OTHERS_PERUNIT : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("99999999999999.99")
    '                .Col = EnumInv.CUMULATIVEBOXES : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 0 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("99999999999999.99")
    '                .Col = EnumInv.FROMBOX : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 0 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("999999.99")
    '                .Col = EnumInv.TOBOX : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 0 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("999999.99")
    '                .Col = EnumInv.BINQTY : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("999999.99")

    '                .CtlEditMode = True : .Enabled = True
    '                If mblnRejTracking = False Then
    '                    If mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
    '                        If mblnServiceInvoiceWithoutSO Then
    '                            .Col = EnumInv.BATCHCOL : .ColHidden = True
    '                        Else
    '                            .Col = EnumInv.BATCHCOL : .ColHidden = False : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = "Batch Details"
    '                        End If
    '                    End If
    '                Else
    '                    If UCase(CmbInvType.Text) = "REJECTION" Or UCase(strInvType) = "REJ" Then
    '                        .Col = EnumInv.BATCHCOL : .ColHidden = True
    '                    Else
    '                        If mblnBatchTrack = True And UCase(CmbInvType.Text) <> "REJECTION" Then
    '                            If mblnServiceInvoiceWithoutSO Then
    '                                .Col = EnumInv.BATCHCOL : .ColHidden = True
    '                            Else
    '                                .Col = EnumInv.BATCHCOL : .ColHidden = False : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = "Batch Details" : .set_ColWidth(EnumInv.BATCHCOL, 1200)
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '            Next intRow
    '        End With
    '        Exit Function
    'ErrHandler:  'The Error Handling Code Starts here
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    '    End Function
    '    Private Sub addRowAtEnterKeyPress(ByRef pintRows As Short)
    '        On Error GoTo ErrHandler
    '        Dim intRowHeight As Short
    '        With Me.SpChEntry
    '            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
    '            For intRowHeight = 1 To pintRows
    '                .MaxRows = .MaxRows + 1
    '                .Row = .MaxRows
    '                .set_RowHeight(.Row, 300)
    '            Next intRowHeight
    '            If .MaxRows > 3 Then .ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsBoth
    '        End With
    '        Exit Sub
    'ErrHandler:  'The Error Handling Code Starts here
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    '        Exit Sub
    '    End Sub

    '    Public Function ToGetDecimalPlaces(ByRef pstrCurrency As String) As Short
    '        Dim rscurrency As ClsResultSetDB
    '        rscurrency = New ClsResultSetDB
    '        rscurrency.GetResult("Select Decimal_Place from Currency_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Currency_code ='" & pstrCurrency & "'")

    '        ToGetDecimalPlaces = Val(rscurrency.GetValue("Decimal_Place"))
    '        rscurrency.ResultSetClose()
    '    End Function

    '    Public Sub displayDeatilsfromCustOrdHdrandDtl()
    '        On Error GoTo ErrHandler
    '        Dim strCustOrdHdr As String
    '        Dim rsCustOrdHdr As ClsResultSetDB
    '        Dim strCurrency As String
    '        Dim intDecimalPlace As Short
    '        'To Get Data from Cusft_Ord_hdr
    '        Dim rsSOReq As ClsResultSetDB
    '        Dim blnSoReq As Boolean
    '        Select Case CmdGrpChEnt.Mode
    '            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
    '                strCustOrdHdr = "Select max(Order_date), SalesTax_Type,AddVAT_Type"
    '                strCustOrdHdr = strCustOrdHdr & "Currency_Code,PerValue from Cust_ord_hdr"
    '                strCustOrdHdr = strCustOrdHdr & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & txtCustCode.Text & "' and Cust_Ref ='"
    '                strCustOrdHdr = strCustOrdHdr & mstrRefNo & "'and Amendment_No ='" & mstrAmmNo & "' Group By salestax_type,AddVAT_Type,currency_code"
    '                rsCustOrdHdr = New ClsResultSetDB
    '                rsCustOrdHdr.GetResult(strCustOrdHdr, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '                strCurrency = rsCustOrdHdr.GetValue("Currency_code")
    '                intDecimalPlace = ToGetDecimalPlaces(strCurrency)
    '                If intDecimalPlace < 2 Then
    '                    intDecimalPlace = 2
    '                End If

    '                txtSaleTaxType.Text = rsCustOrdHdr.GetValue("SalesTax_Type")

    '                Call txtSaleTaxType_Validating(txtSaleTaxType, New System.ComponentModel.CancelEventArgs(False))


    '                rsCustOrdHdr.ResultSetClose()
    '            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
    '                rsSOReq = New ClsResultSetDB
    '                rsSOReq.GetResult("Select isnull(SORequired,0) as SORequired from saleConf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type ='TRF' and Sub_Type_Description='" & Trim(CmbInvSubType.Text) & "' and  (fin_start_date <= getdate() and fin_end_date >= getdate())")
    '                If rsSOReq.GetNoRows > 0 Then
    '                    blnSoReq = rsSOReq.GetValue("SORequired")
    '                Else
    '                    blnSoReq = False
    '                End If
    '                rsSOReq.ResultSetClose()
    '                If UCase(CStr((Trim(CmbInvType.Text)) = "NORMAL INVOICE")) Or ((Trim(CmbInvType.Text) = "TRANSFER INVOICE") And blnSoReq) Or UCase(CStr((Trim(CmbInvType.Text)) = "JOBWORK INVOICE")) Or UCase(CStr((Trim(CmbInvType.Text)) = "EXPORT INVOICE")) Or (UCase(CStr((Trim(CmbInvType.Text)) = "SERVICE INVOICE")) And Not mblnServiceInvoiceWithoutSO) Then
    '                    If CBool(UCase(CStr((Trim(CmbInvSubType.Text)) <> "SCRAP"))) Then
    '                        If Len(Trim(txtRefNo.Text)) Then
    '                            strCustOrdHdr = "Select max(Order_date),SalesTax_Type,AddVAT_Type,Currency_code,PerValue,term_payment, surcharge_code from Cust_ord_hdr"
    '                            strCustOrdHdr = strCustOrdHdr & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & txtCustCode.Text & "' and Cust_Ref ='"
    '                            strCustOrdHdr = strCustOrdHdr & mstrRefNo & "'and Amendment_No ='" & mstrAmmNo & "'"
    '                            strCustOrdHdr = strCustOrdHdr & " and active_flag = 'A' Group by salestax_type,currency_code,AddVAT_Type,PerValue,term_payment, surcharge_code"
    '                            rsCustOrdHdr = New ClsResultSetDB
    '                            rsCustOrdHdr.GetResult(strCustOrdHdr, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '                            strSOSaleTaxType = rsCustOrdHdr.GetValue("SalesTax_Type")
    '                            'txtSaleTaxType.Text = rsCustOrdHdr.GetValue("SalesTax_Type")

    '                            strCurrency = rsCustOrdHdr.GetValue("Currency_code")
    '                            txtCreditTerms.Text = rsCustOrdHdr.GetValue("term_payment")

    '                            Dim RsTermMst As New ClsResultSetDB
    '                            RsTermMst.GetResult("SELECT CRTRM_TERMID,CRTRM_DESC FROM GEN_CREDITTRMMASTER WHERE UNIT_CODE='" + gstrUNITID + "'  AND CRTRM_TERMID='" + txtCreditTerms.Text + "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '                            If RsTermMst.GetNoRows > 0 Then
    '                                Me.lblCreditTermDesc.Text = RsTermMst.GetValue("CRTRM_DESC")
    '                            End If
    '                            RsTermMst.ResultSetClose()
    '                            RsTermMst = Nothing


    '                            intDecimalPlace = ToGetDecimalPlaces(strCurrency)
    '                            If intDecimalPlace < 2 Then
    '                                intDecimalPlace = 2
    '                            End If

    '                            rsCustOrdHdr.ResultSetClose()
    '                        End If
    '                    End If
    '                End If
    '        End Select
    '        Call DisplayDetailsInSpread(strCurrency)
    '        Exit Sub
    'ErrHandler:  'The Error Handling Code Starts here
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    '    End Sub

    '    Private Function DisplayDetailsInSpread(ByRef pstrCurrency As String) As Boolean
    '        'Description    -  To display Details From Sales_Dtl Acc To Location Code,Challan No and Drawing No
    '        On Error GoTo ErrHandler
    '        Dim intLoopCounter As Short
    '        Dim intRecordCount As Short
    '        Dim Intcounter As Short
    '        Dim inti As Short
    '        Dim strsaledtl As String
    '        Dim dblPacking As Double
    '        Dim varItem_Code As Object
    '        Dim varCustItemCode As Object
    '        Dim varItemAlready As Object
    '        Dim rsSalesDtl As ClsResultSetDB
    '        Dim rsBatch As ClsResultSetDB
    '        Dim rsSOReq As ClsResultSetDB
    '        Dim blnQtyChkAccToMeasureCode As Boolean
    '        Dim intDecimal As Short
    '        Dim strMin As String
    '        Dim strMax As String
    '        Dim intloopcounter1 As Short
    '        Dim strCompileString As String
    '        Dim blnSoReq As Boolean


    '        Select Case Me.CmdGrpChEnt.Mode
    '            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
    '                strsaledtl = ""
    '                strsaledtl = "SELECT Location_Code,Doc_No,Suffix,Item_Code,Sales_Quantity,From_Box,To_Box,Rate,Sales_Tax,Excise_Tax,Packing,Others,Cust_Mtrl,Year,Cust_Item_Code,Cust_Item_Desc,Tool_Cost,Measure_Code,Excise_type,SalesTax_type,CVD_type,SAD_type,GL_code,SL_code,Basic_Amount,Accessible_amount,CVD_Amount,SVD_amount,Excise_per,CVD_per,SVD_per,CustMtrl_Amount,ToolCost_amount,pervalue,TotalExciseAmount,SupplementaryInvoiceFlag,To_Location,Discount_type,Discount_amt,Discount_perc,From_Location,Cust_ref,Amendment_No,SRVDINO,SRVLocation,USLOC,SchTime,BinQuantity,Packing_Type,ItemPacking_Amount,Item_remark,pkg_amount,csiexcise_amount,ADD_EXCISE_TYPE,ADD_EXCISE_PER,ADD_EXCISE_AMOUNT from Sales_Dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Location_Code='" & Trim(txtLocationCode.Text) & "'"
    '                strsaledtl = strsaledtl & " and Doc_No=" & Val(txtChallanNo.Text) & " and Cust_Item_Code in(" & Trim(mstrItemCode) & ")"
    '            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
    '                rsSOReq = New ClsResultSetDB
    '                rsSOReq.GetResult("Select isnull(SORequired,0) as SORequired from saleConf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type ='TRF' and Sub_Type_Description='" & Trim(CmbInvSubType.Text) & "' and  (fin_start_date <= getdate() and fin_end_date >= getdate())")
    '                If rsSOReq.GetNoRows > 0 Then
    '                    blnSoReq = rsSOReq.GetValue("SORequired")
    '                Else
    '                    blnSoReq = False
    '                End If
    '                rsSOReq.ResultSetClose()
    '                If UCase(CStr(Trim(CmbInvType.Text) = "NORMAL INVOICE")) Or (UCase(CStr(Trim(CmbInvType.Text) = "TRANSFER INVOICE")) And blnSoReq) Or UCase(CStr(Trim(CmbInvType.Text) = "JOBWORK INVOICE")) Or UCase(CStr(Trim(CmbInvType.Text) = "EXPORT INVOICE")) Or UCase(CStr(Trim(CmbInvType.Text) = "SERVICE INVOICE")) Then
    '                    strsaledtl = "Select Item_Code,Cust_DrgNo,Rate,Cust_Mtrl,Packing,Packing_Type,Others,tool_Cost,Excise_Duty from Cust_ord_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
    '                    strsaledtl = strsaledtl & "Account_Code ='" & txtCustCode.Text & "'and Cust_ref ='"
    '                    strsaledtl = strsaledtl & txtRefNo.Text & "' and Amendment_No = '" & mstrAmmNo & "'and "
    '                    strsaledtl = strsaledtl & " Active_flag ='A' and Cust_DrgNo in(" & mstrItemCode & ")"
    '                Else
    '                    strsaledtl = ""
    '                    strsaledtl = "SELECT Item_Code,standard_Rate from Item_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
    '                    strsaledtl = strsaledtl & " Status = 'A' and Hold_flag <> 1 and Item_Code in (" & mstrItemCode & ")"
    '                End If
    '        End Select
    '        rsSalesDtl = New ClsResultSetDB
    '        rsSalesDtl.GetResult(strsaledtl, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '        Dim intLoopCount As Short
    '        Dim varCumulative As Object
    '        Dim intcnt As Short
    '        Dim dblqty As Double
    '        Dim dblStock As Double
    '        Dim strpono As String
    '        Dim intMaxSerial_No As Short
    '        Dim strCustDrgNo As Object
    '        Dim strSqlBins As String
    '        Dim dblBins As Double
    '        Dim rsBinQty As ClsResultSetDB
    '        If rsSalesDtl.GetNoRows > 0 Then
    '            intRecordCount = rsSalesDtl.GetNoRows
    '            ReDim mdblPrevQty(intRecordCount - 1) ' To get value of Quantity in Arrey for updation in despatch
    '            ReDim mdblToolCost(intRecordCount - 1) ' To get value of Quantity i


    '            ' Call Me.SpChEntry.SetText(EnumInv.ENUMQUANTITY, 1, 0)
    '            '-----------------------
    '            SpChEntry.MaxRows = 0
    '            SpChEntry.MaxRows = SpChEntry.MaxRows + 1
    '            BlankTaxDetails()
    '            GetDefaultTaxexFromSO()
    '            FRMMKTTRN0076A.DeleteTmpTable()
    '            '-----------------------

    '            'If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
    '            '    If SpChEntry.MaxRows > 0 Then
    '            '        varItemAlready = Nothing
    '            '        Call SpChEntry.GetText(EnumInv.ENUMITEMCODE, 1, varItemAlready)
    '            '        If Len(Trim(varItemAlready)) = 0 Then
    '            '            Call addRowAtEnterKeyPress(intRecordCount - 1)
    '            '        End If
    '            '    Else
    '            '        Call addRowAtEnterKeyPress(intRecordCount)
    '            '    End If
    '            'Else
    '            '    Call addRowAtEnterKeyPress(intRecordCount - 1)
    '            'End If



    '            rsSalesDtl.MoveFirst()
    '            If CmbInvType.Text = "NORMAL INVOICE" Or CmbInvType.Text = "JOBWORK INVOICE" Or CmbInvType.Text = "EXPORT INVOICE" Or (CmbInvType.Text = "SERVICE INVOICE" And Not mblnServiceInvoiceWithoutSO) Then
    '                If UCase(Trim(CmbInvSubType.Text)) <> "SCRAP" Then
    '                    For intLoopCount = 1 To intRecordCount
    '                        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
    '                            mdblToolCost(intLoopCount - 1) = Val(rsSalesDtl.GetValue("Tool_Cost"))
    '                        Else
    '                            mdblToolCost(intLoopCount - 1) = Val(rsSalesDtl.GetValue("Tool_Cost"))
    '                        End If
    '                        rsSalesDtl.MoveNext()
    '                    Next
    '                End If
    '            End If
    '            rsSalesDtl.MoveFirst()
    '            intDecimal = ToGetDecimalPlaces(pstrCurrency)
    '            Call SetMaxLengthInSpread(intDecimal)



    '            'If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
    '            '    If SpChEntry.MaxRows > 0 Then
    '            '        varItemAlready = Nothing
    '            '        Call SpChEntry.GetText(EnumInv.ENUMITEMCODE, 1, varItemAlready)
    '            '        If Len(Trim(varItemAlready)) > 0 Then
    '            '            inti = SpChEntry.MaxRows + 1
    '            '            SpChEntry.MaxRows = SpChEntry.MaxRows + intRecordCount
    '            '            intRecordCount = SpChEntry.MaxRows
    '            '        Else
    '            '            inti = 1
    '            '        End If
    '            '    Else
    '            '        inti = 1
    '            '        SpChEntry.MaxRows = intRecordCount
    '            '    End If
    '            'Else
    '            '    inti = 1
    '            'End If

    '            inti = SpChEntry.MaxRows

    '            For intLoopCounter = inti To intRecordCount
    '                With Me.SpChEntry
    '                    Select Case Me.CmdGrpChEnt.Mode
    '                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
    '                            .Row = 1 : .Row2 = .MaxRows : .Col = 0 : .Col2 = .MaxCols
    '                            .Enabled = True : .BlockMode = True : .Lock = True : .BlockMode = False
    '                            Call .SetText(EnumInv.ENUMITEMCODE, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
    '                            Call .SetText(EnumInv.CUSTPARTNO, intLoopCounter, rsSalesDtl.GetValue("Cust_Item_Code"))
    '                            Call .SetText(EnumInv.RATE_PERUNIT, intLoopCounter, rsSalesDtl.GetValue("Rate") * Val("CTLPERVALUE"))


    '                            Call .SetText(EnumInv.ENUMQUANTITY, intLoopCounter, rsSalesDtl.GetValue("Sales_Quantity"))
    '                            mdblPrevQty(intLoopCounter - 1) = Nothing
    '                            Call .GetText(EnumInv.ENUMQUANTITY, intLoopCounter, mdblPrevQty(intLoopCounter - 1))
    '                            If mblnRejTracking = True Then
    '                                Call AddMaxAllowedQuanity(intLoopCounter)
    '                            End If
    '                            Call .SetText(EnumInv.PACKING, intLoopCounter, rsSalesDtl.GetValue("Packing_Type"))



    '                            Call .SetText(EnumInv.FROMBOX, intLoopCounter, rsSalesDtl.GetValue("From_Box"))
    '                            Call .SetText(EnumInv.TOBOX, intLoopCounter, rsSalesDtl.GetValue("To_Box"))
    '                            Call .SetText(EnumInv.BINQTY, intLoopCounter, rsSalesDtl.GetValue("BinQuantity"))

    '                            If mblnRejTracking = True And mblnBatchTracking = True Then
    '                                strCompileString = GetDocumentDetail(rsSalesDtl.GetValue("Item_Code"))
    '                            End If
    '                            If intLoopCounter = 1 Then
    '                                Call .SetText(EnumInv.CUMULATIVEBOXES, intLoopCounter, (rsSalesDtl.GetValue("To_Box") - rsSalesDtl.GetValue("From_Box")) + 1)
    '                            Else
    '                                varCumulative = Nothing
    '                                Call .GetText(EnumInv.CUMULATIVEBOXES, intLoopCounter - 1, varCumulative)
    '                                Call .SetText(EnumInv.CUMULATIVEBOXES, intLoopCounter, varCumulative + ((rsSalesDtl.GetValue("To_Box") - rsSalesDtl.GetValue("From_Box")) + 1))
    '                            End If
    '                            If mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
    '                                varItem_Code = Nothing
    '                                Call .GetText(EnumInv.ENUMITEMCODE, intLoopCounter, varItem_Code)
    '                                varCustItemCode = Nothing
    '                                Call .GetText(EnumInv.CUSTPARTNO, intLoopCounter, varCustItemCode)
    '                                rsBatch = New ClsResultSetDB
    '                                Call rsBatch.GetResult("Select Batch_No,Batch_Date,Batch_Qty from ItemBatch_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Item_Code = '" & varItem_Code & "' and Cust_Item_Code = '" & varCustItemCode & "' and  From_Location = '" & mstrLocationCode & "' and Doc_No = " & Trim(Me.txtChallanNo.Text) & " and Doc_type = 9999 ", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '                                If rsBatch.RowCount > 0 Then
    '                                    ReDim Preserve mBatchData(intLoopCounter)
    '                                    ReDim Preserve mBatchData(intLoopCounter).Batch_No(rsBatch.RowCount)
    '                                    ReDim Preserve mBatchData(intLoopCounter).Batch_Date(rsBatch.RowCount)
    '                                    ReDim Preserve mBatchData(intLoopCounter).Batch_Quantity(rsBatch.RowCount)
    '                                    Intcounter = 1
    '                                    While Not rsBatch.EOFRecord
    '                                        mBatchData(intLoopCounter).Batch_No(Intcounter) = rsBatch.GetValue("Batch_No")
    '                                        mBatchData(intLoopCounter).Batch_Date(Intcounter) = ConvertToDate(VB6.Format(rsBatch.GetValue("Batch_Date"), gstrDateFormat))
    '                                        mBatchData(intLoopCounter).Batch_Quantity(Intcounter) = rsBatch.GetValue("Batch_Qty")
    '                                        Intcounter = Intcounter + 1
    '                                        rsBatch.MoveNext()
    '                                    End While
    '                                End If
    '                                rsBatch.ResultSetClose()
    '                            End If
    '                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
    '                            .Enabled = True
    '                            .Row = 1 : .Row2 = .MaxRows : .Col = 0 : .Col2 = .MaxCols : .BlockMode = True : .Lock = False : .set_RowHeight(.MaxRows, 12) : .BlockMode = False
    '                            If (Trim(CmbInvType.Text) = "NORMAL INVOICE") Or (UCase(CStr(Trim(CmbInvType.Text) = "TRANSFER INVOICE")) And blnSoReq) Or (Trim(CmbInvType.Text) = "JOBWORK INVOICE") Or (Trim(CmbInvType.Text) = "EXPORT INVOICE") Or (Trim(CmbInvType.Text) = "SERVICE INVOICE") Then

    '                                Call .SetText(EnumInv.ENUMITEMCODE, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
    '                                Call .SetText(EnumInv.CUSTPARTNO, intLoopCounter, rsSalesDtl.GetValue("Cust_DrgNo"))
    '                                Call .SetText(EnumInv.RATE_PERUNIT, intLoopCounter, (Val(rsSalesDtl.GetValue("Rate"))))

    '                                Call Me.SpChEntry.SetText(EnumInv.ENUMQUANTITY, intLoopCounter, 0)
    '                                Call .SetText(EnumInv.CUSTSUPPMAT_PERUNIT, intLoopCounter, (Val(rsSalesDtl.GetValue("Cust_Mtrl")) * Val(0)))

    '                                Call .SetText(EnumInv.PACKING, intLoopCounter, rsSalesDtl.GetValue("Packing_Type"))
    '                                Call .SetText(EnumInv.OTHERS_PERUNIT, intLoopCounter, (Val(rsSalesDtl.GetValue("Others")) * Val(0)))




    '                            Else
    '                                Call .SetText(EnumInv.ENUMITEMCODE, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
    '                                Call .SetText(EnumInv.CUSTPARTNO, intLoopCounter, rsSalesDtl.GetValue("Item_code"))
    '                                Call .SetText(EnumInv.RATE_PERUNIT, intLoopCounter, (rsSalesDtl.GetValue("Standard_Rate") * Val("CTLPERVALUE")))
    '                            End If
    '                            rsBinQty = New ClsResultSetDB
    '                            strCustDrgNo = Nothing
    '                            Call SpChEntry.GetText(EnumInv.CUSTPARTNO, intLoopCounter, strCustDrgNo)
    '                            strSqlBins = "Select isnull(BinQuantity,1) as BinQuantity from custitem_mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  cust_drgno= '" & strCustDrgNo & "' and Account_code='" & Trim(Me.txtCustCode.Text) & "' "
    '                            rsBinQty.GetResult(strSqlBins)
    '                            If rsBinQty.GetNoRows > 0 Then
    '                                If rsBinQty.GetValue("BinQuantity") = 0 Then
    '                                    dblBins = 1
    '                                Else
    '                                    dblBins = rsBinQty.GetValue("BinQuantity")
    '                                End If
    '                            Else
    '                                dblBins = 1
    '                            End If
    '                            rsBinQty.ResultSetClose()
    '                            Call SpChEntry.SetText(EnumInv.BINQTY, intLoopCounter, dblBins)
    '                    End Select
    '                End With
    '                rsSalesDtl.MoveNext()
    '            Next intLoopCounter

    '        End If

    '        If SpChEntry.MaxRows > 3 Then
    '            SpChEntry.ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsBoth
    '        End If
    '        rsSalesDtl.ResultSetClose()
    '        Exit Function
    'ErrHandler:  'The Error Handling Code Starts here
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    '    End Function


    '    Private Function GetMode() As String
    '        On Error GoTo ErrHandler
    '        Select Case CmdGrpChEnt.Mode
    '            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
    '                GetMode = "ADD"
    '            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
    '                GetMode = "EDIT"
    '            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
    '                GetMode = "VIEW"
    '        End Select
    '        Exit Function
    'ErrHandler:
    '        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '    End Function

    '    Private Function SelectDataFromTable(ByRef mstrFieldName As String, ByRef mstrTableName As String, ByRef mstrCondition As String) As String
    '        Dim StrSQLQuery As String
    '        Dim GetDataFromTable As ClsResultSetDB
    '        On Error GoTo ErrHandler
    '        If UCase(mstrTableName) = "SALESCHALLAN_DTL" Then
    '            StrSQLQuery = "Select TOP 1 " & mstrFieldName & " From " & mstrTableName & mstrCondition
    '        Else
    '            StrSQLQuery = "Select " & mstrFieldName & " From " & mstrTableName & mstrCondition
    '        End If
    '        GetDataFromTable = New ClsResultSetDB
    '        If GetDataFromTable.GetResult(StrSQLQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic) Then
    '            If GetDataFromTable.GetNoRows > 0 Then
    '                GetDataFromTable.MoveFirst()
    '                SelectDataFromTable = GetDataFromTable.GetValue(mstrFieldName)
    '            Else
    '                SelectDataFromTable = ""
    '            End If
    '        Else
    '            SelectDataFromTable = ""
    '        End If
    '        GetDataFromTable.ResultSetClose()
    '        Exit Function
    'ErrHandler:
    '        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '    End Function

    '    Public Sub ShowMultipleSOItemHelp()
    '        Dim frmMKTTRN0021B As New frmMKTTRN0021B
    '        On Error GoTo ErrHandler
    '        If Trim(txtCustCode.Text) = "" Then
    '            MsgBox("Select Customer Code before selecting items.", MsgBoxStyle.Information, ResolveResString(100))
    '            CmdCustCodeHelp.Focus()
    '            mblnCheckArray = False
    '            Exit Sub
    '        Else
    '            With frmMKTTRN0021B

    '                .mdtInvoiceDate = GetServerDate()

    '                .mstrCustomerCode = Trim(txtCustCode.Text)

    '                .mstrInvType = mstrInvTypenew

    '                .mstrInvSubType = mstrInvSubTypenew

    '                .mstrLocationCode = Trim(txtLocationCode.Text)

    '                .mstrStockLocation = mstrStockLocation

    '                .mstrDocNo = Trim(txtChallanNo.Text)

    '                .mstrMode = GetMode()

    '                If .FillSOItemHelp = True Then mblnCheckArray = True Else mblnCheckArray = False
    '            End With
    '        End If
    '        Exit Sub
    'ErrHandler:
    '        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '    End Sub

    '    Private Function OriginalRefNoOVER(ByVal strRefNumber As String) As Boolean
    '        On Error GoTo ErrHandler
    '        '1st Check if Any Blank Amendment no for Ref No. Exists
    '        If SelectDataFromTable("Active_Flag", "Cust_ORD_HDR", " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code ='" & Trim(txtCustCode.Text) & "' AND Cust_Ref = '" & txtRefNo.Text & "' AND Amendment_No = ''") = "O" Then
    '            OriginalRefNoOVER = True
    '        Else
    '            OriginalRefNoOVER = False
    '        End If
    '        Exit Function
    'ErrHandler:  'The Error Handling Code Starts here
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    '    End Function

    '    Private Sub GetInvoiceType()
    '        Dim Sqlcmd As New SqlCommand
    '        Dim Rd As SqlDataReader
    '        Dim strSql As String
    '        Sqlcmd.Connection = SqlConnectionclass.GetConnection()
    '        Sqlcmd.CommandType = CommandType.Text
    '        Try
    '            strSql = "SELECT DISTINCT DESCRIPTION FROM SALECONF WHERE UNIT_CODE=@UNITCODE AND  INVOICE_TYPE  IN('INV') AND (FIN_START_DATE <= GETDATE() AND FIN_END_DATE >= GETDATE())  "
    '            Sqlcmd.CommandText = strSql
    '            Sqlcmd.Parameters.Clear()
    '            Sqlcmd.Parameters.Add("@UNITCODE", SqlDbType.VarChar).Value = gstrUNITID
    '            Rd = Sqlcmd.ExecuteReader()
    '            If Rd.HasRows = True Then
    '                While Rd.Read
    '                    CmbInvType.Items.Add(Rd("DESCRIPTION"))
    '                End While
    '            End If
    '            If Rd.IsClosed = False Then Rd.Close()
    '            CmbInvType.SelectedIndex = 0

    '            strSql = "SELECT SUB_TYPE_DESCRIPTION FROM SALECONF WHERE UNIT_CODE=@UNITCODE AND  INVOICE_TYPE  IN('INV') AND SUB_TYPE='T' AND (FIN_START_DATE <= GETDATE() AND FIN_END_DATE >= GETDATE())"
    '            Sqlcmd.CommandText = strSql
    '            Sqlcmd.Parameters.Clear()
    '            Sqlcmd.Parameters.Add("@UNITCODE", SqlDbType.VarChar).Value = gstrUNITID
    '            Rd = Sqlcmd.ExecuteReader()
    '            If Rd.HasRows = True Then
    '                While Rd.Read
    '                    CmbInvSubType.Items.Add(Rd("SUB_TYPE_DESCRIPTION"))
    '                End While
    '            End If
    '            If Rd.IsClosed = False Then Rd.Close()
    '            CmbInvSubType.SelectedIndex = 0

    '            CmbTransType.Items.Add("R - Road")
    '            CmbTransType.Items.Add("L - Rail")
    '            CmbTransType.Items.Add("S - Sea")
    '            CmbTransType.Items.Add("A - Air")
    '            CmbTransType.Items.Add("H - Hand")
    '            CmbTransType.Items.Add("C - Courier")
    '            CmbTransType.SelectedIndex = 0

    '        Catch Ex As Exception
    '            MsgBox(Ex.Message.ToString, MsgBoxStyle.Information, ResolveResString(100))
    '        Finally
    '            If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
    '            Sqlcmd.Connection.Dispose()
    '            Sqlcmd.Dispose()
    '        End Try
    '    End Sub


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
        lblNetAdvAmt.Text = 0
        lblCancelledInvoice.Visible = False
        blnBillFlag = False
        dblBasicValue = 0
        mstrItemCode = String.Empty
        dblGrinQuantityForSale = 0
        ' txtLocationCode.Text = gstrUNITID
        'lblLocCodeDes.Text = String.Empty
        txtChallanNo.Text = String.Empty
        txtCustCode.Text = String.Empty
        lblCustCodeDes.Text = String.Empty
        txtRefNo.Text = String.Empty
        dtpDateSO.Text = Date.Today.ToString
        lblRoundOff.Text = String.Empty
        lblRoundOff.Text = 0
        ctlInsurance.Text = String.Format("{0:n2}", 0)
        txtFreight.Text = String.Format("{0:n2}", 0)
        'String.Format("{0:n2}", 0)
        LblNetInvoiceValue.Text = 0


        Me.lblInternalPartDesc.Text = String.Empty
        Me.lblCustPartDesc.Text = String.Empty
        ' Me.dtpDateDesc.Value = GetServerDate.ToString("dd/MMM/yyyy")


        sspr.MaxRows = 0
        ' FRMMKTTRN0076A.DeleteTmpTable()
        ' GetItemDescription()

    End Sub
    '    Private Sub BlankTaxDetails()
    '        dblBasicValue = 0
    '        dblGrinQuantityForSale = 0
    '        ctlInsurance.Text = "0.00"
    '        txtFreight.Text = "0.00"
    '        txtSaleTaxType.Text = String.Empty
    '        lblSaltax_Per.Text = "0.00"
    '        lblSalesTaxValue.Text = "0.00"
    '        txtECSSTaxType.Text = String.Empty
    '        lblECSStax_Per.Text = "0.00"
    '        lblEcessValue.Text = "0.00"
    '        txtSECSSTaxType.Text = String.Empty
    '        lblSECSStax_Per.Text = "0.00"
    '        lblHCessValue.Text = "0"
    '        txtDiscountAmt.Text = "0.00"
    '        lblBasicValue.Text = "0.00"
    '        lblAssValue.Text = "0.00"
    '        lblExciseValue.Text = "0.00"
    '        lblEcessValue.Text = "0.00"
    '        lblHCessValue.Text = "0.00"
    '        lblSalesTaxValue.Text = "0.00"
    '        lblBasicExciseAndCess.Text = "0.00"
    '        lblAEDValue.Text = "0.00"

    '        txtAddVAT.Text = String.Empty
    '        lblAddVAT.Text = "0.00"
    '        lblAddVATValue.Text = "0.00"
    '    End Sub
    '    Private Function StockLocationSalesConf(ByRef pstrInvType As String, ByRef pstrInvSubtype As String, ByRef pstrFeild As String) As String
    '        Dim rsSalesConf As ClsResultSetDB
    '        Dim StockLocation As String
    '        On Error GoTo ErrHandler
    '        rsSalesConf = New ClsResultSetDB
    '        Select Case pstrFeild
    '            Case "DESCRIPTION"
    '                rsSalesConf.GetResult("Select Stock_Location from SaleConf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Description ='" & Trim(pstrInvType) & "' and Sub_type_Description ='" & Trim(pstrInvSubtype) & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
    '            Case "TYPE"
    '                rsSalesConf.GetResult("Select Stock_Location from SaleConf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_type ='" & Trim(pstrInvType) & "' and Sub_type ='" & Trim(pstrInvSubtype) & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
    '        End Select
    '        If rsSalesConf.GetNoRows > 0 Then
    '            StockLocation = rsSalesConf.GetValue("Stock_Location")
    '        End If
    '        rsSalesConf.ResultSetClose()
    '        StockLocationSalesConf = StockLocation
    '        Exit Function
    'ErrHandler:  'The Error Handling Code Starts here
    '        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    '    End Function

    '#End Region
    Private Sub TRADINGINVOICE_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = "'" Then e.Handled = True
    End Sub
    Private Sub TRADINGINVOICE_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Label9.Visible = False
            ctlInsurance.Visible = False
            ctlInsurance.Text = 0
            Label10.Visible = False
            txtFreight.Text = 0
            txtFreight.Visible = False
            Me.Visible = False
            Me.KeyPreview = True
            Me.MdiParent = mdifrmMain
            Me.Icon = mdifrmMain.Icon
            Call FitToClient(Me, PnlMain, ctlHeader, CmdGrpChEnt, 250)
            Me.Group1.Enabled = False
            Me.Group2.Enabled = False
            Me.Group3.Enabled = False
            Me.Group4.Enabled = False
            SetPRGridHeading()
            'GetInvoiceType()
            BlankFields()
            Me.Cmditems.Visible = False
            CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
            CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
            CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
            CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
            CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
            CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
            CmdChallanNo.Enabled = True

            'SubGetRoundoffConfig()
            'CmdGrpChEnt_ButtonClick(sender, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL))
        Catch Ex As Exception
            MsgBox(Ex.Message, MsgBoxStyle.Information, ResolveResString(100))
        Finally
            Me.Visible = True
        End Try
    End Sub
#Region "GRID FORMATTING"
    Private Sub SetPRGridHeading()
        ' On Error GoTo ErrHandler
        Try


            With Me.sspr

                .UnitType = FPSpreadADO.UnitTypeConstants.UnitTypeTwips
                .MaxRows = 0
                .MaxCols = EnumInv.PrevQty
                .RowHeaderCols = 0
                .Row = 0
                .set_RowHeight(0, 400)
                .ColHeaderRows = 2
                '---- SL NO----
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.SL
                .Text = "S.l."
                .set_ColWidth(EnumInv.SL, 400)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5


                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "S.l."

                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.SL, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                '--FREEZ COLLUMNS
                ' .ColsFrozen = 7


                '-- Item Code-------

                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.ItemCode
                .Text = "Item Code"
                .set_ColWidth(EnumInv.ItemCode, 2000)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5


                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "Item Code"

                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.ItemCode, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                '.ColsFrozen = 6
                '.CellType = FPSpreadADO.CellTypeConstants.CellTypeButton



                ''- HSN/SAC No
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.HSN_SAC_No
                .Text = "HSN/SAC No"
                .set_ColWidth(EnumInv.HSN_SAC_No, 1000)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "HSN/SAC No"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.HSN_SAC_No, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)

                ''- HSN/SAC Type
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.HSN_SAC_Type
                .Text = "HSN/SAC Type"
                .set_ColWidth(EnumInv.HSN_SAC_Type, 1200)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "HSN/SAC Type"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.HSN_SAC_Type, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                ''- Quantity
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.Quantity
                .Text = "Quantity"
                .set_ColWidth(EnumInv.Quantity, 900)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "Quantity"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.Quantity, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                ''- Rate
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.Rate
                .Text = "Rate"
                .set_ColWidth(EnumInv.Rate, 700)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "Rate"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.Rate, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)

                ''- Basic Value
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.Basic_value
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "PI Basic Value"
                .set_ColWidth(EnumInv.Basic_value, 1200)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.LightPink

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "PI Basic Value"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.Basic_value, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                .TypeButtonColor = Color.LightPink
                ''- Discount%
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.DiscountPer
                .Text = "Discount%"
                .set_ColWidth(EnumInv.DiscountPer, 1100)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "Discount%"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.DiscountPer, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)

                ''- Discount Value
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.DiscountVal
                .Text = "Discount Val"
                .set_ColWidth(EnumInv.DiscountVal, 1100)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "Discount Val"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.DiscountVal, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                ''- Assable Value
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.Assable_Value
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "Assable Value"
                .set_ColWidth(EnumInv.Assable_Value, 1300)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.LightGreen

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "Assable Value"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.Assable_Value, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                .TypeButtonColor = Color.LightGreen

                ''- Advance Amount
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.Advance_Amt
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "PI Value for Advance"
                .set_ColWidth(EnumInv.Advance_Amt, 1750)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.LightGreen

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "PI Value for Advance"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.Advance_Amt, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                .TypeButtonColor = Color.LightGreen

                ''- IGST----- IGST TAX TYPE
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.IGST_Tax_type
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "IGST"
                .set_ColWidth(EnumInv.IGST_Tax_type, 1300)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.IGST_Tax_type, FPSpreadADO.CoordConstants.SpreadHeader, 3, 1)
                .TypeButtonColor = Color.SkyBlue


                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "IGST Tax Type"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                ''- IGST Tax %
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.IGST_Tax_Per
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "IGST"
                .set_ColWidth(EnumInv.IGST_Tax_Per, 1000)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.SkyBlue

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "IGST Tax %"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                ''- IGST Tax Value
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.IGST_Tax_Value
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "IGST"
                .set_ColWidth(EnumInv.IGST_Tax_Value, 1000)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.SkyBlue

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "IGST Tax Val"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.SkyBlue

                ''- CGST-----CGST TAX Type
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.CGST_Tax_type
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "CGST"
                .set_ColWidth(EnumInv.CGST_Tax_type, 1300)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .BackColor = Color.Yellow
                .AddCellSpan(EnumInv.CGST_Tax_type, FPSpreadADO.CoordConstants.SpreadHeader, 3, 1)
                .TypeButtonColor = Color.BlanchedAlmond

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "CGST Tax Type"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                '- CGST Tax %
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.CGST_Tax_Per
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "CGST"
                .set_ColWidth(EnumInv.CGST_Tax_Per, 1000)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.BlanchedAlmond

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "CGST Tax %"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                '- CGST TAx Value
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.CGST_Tax_Value
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "CGST"
                .set_ColWidth(EnumInv.CGST_Tax_Value, 1100)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.BlanchedAlmond

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "CGST Tax Val"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.BlanchedAlmond

                '- SGST--SGST Tax Type
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.SGST_Tax_type
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "SGST"
                .set_ColWidth(EnumInv.SGST_Tax_type, 1200)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.SkyBlue
                .AddCellSpan(EnumInv.SGST_Tax_type, FPSpreadADO.CoordConstants.SpreadHeader, 3, 1)
                .TypeButtonColor = Color.SkyBlue

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "SGST Tax Type"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                '- SGST Tax Per
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.SGST_Tax_Per
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "SGST"
                .set_ColWidth(EnumInv.SGST_Tax_Per, 1000)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.SkyBlue

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "SGST Tax %"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                '- SGST Tax Value
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.SGST_Tax_Value
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "SGST"
                .set_ColWidth(EnumInv.SGST_Tax_Value, 1100)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.SkyBlue

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "SGST Tax Val"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.SkyBlue


                ''- UTGST----UTGST Tax Type
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.UTGST_Tax_type
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "UTGST"
                .set_ColWidth(EnumInv.UTGST_Tax_type, 1300)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .BackColor = Color.Yellow
                .AddCellSpan(EnumInv.UTGST_Tax_type, FPSpreadADO.CoordConstants.SpreadHeader, 3, 1)
                .TypeButtonColor = Color.BlanchedAlmond

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "UTGST Tax Type"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5


                '- UTGST Tax %
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.UTGST_Tax_Per
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "UTGST"
                .set_ColWidth(EnumInv.UTGST_Tax_Per, 1000)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.BlanchedAlmond

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "UTGST Tax %"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5


                '- UTGST Tax Value
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.UTGST_Tax_Value
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "UTGST"
                .set_ColWidth(EnumInv.UTGST_Tax_Value, 1200)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.BlanchedAlmond

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "UTGST Tax Val"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.BlanchedAlmond

                ''- CSESS--CSESS Tax Type
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.CSESS_Tax_type
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "CSESS"
                .set_ColWidth(EnumInv.CSESS_Tax_type, 1300)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.SkyBlue

                .AddCellSpan(EnumInv.CSESS_Tax_type, FPSpreadADO.CoordConstants.SpreadHeader, 3, 1)

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "CSESS Tax Type"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                '- CSESS Tax %
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.CSESS_Tax_Per
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "CSESS"
                .set_ColWidth(EnumInv.CSESS_Tax_Per, 1000)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.SkyBlue

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "CSESS Tax %"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                '- CSESS Tax Value
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.CSESS_Tax_Value
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "CSESS"
                .set_ColWidth(EnumInv.CSESS_Tax_Value, 1200)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.SkyBlue

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "CSESS Tax Val"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.SkyBlue

                ''- ITEM Total
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.ItemTotal
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "Item Total"
                .set_ColWidth(EnumInv.ItemTotal, 1200)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .TypeButtonColor = Color.Green

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "Item Total"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.ItemTotal, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                .TypeButtonColor = Color.Green
                ''- Remarks.
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.Remarks
                .Text = "Remarks"
                .set_ColWidth(EnumInv.Remarks, 1500)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "Remarks"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.Remarks, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)

                ''- Internal Item Code Description.
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.Internal_Item_Desc
                .Text = "Internal Item Desc"
                .set_ColWidth(EnumInv.Internal_Item_Desc, 1500)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .ColHidden = True

                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "Internal Item Desc"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.Internal_Item_Desc, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                .ColHidden = True

                ''- Customer Drawing No .
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.Cust_Drgno
                .Text = "Customer DrgNo"
                .set_ColWidth(EnumInv.Cust_Drgno, 1500)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5


                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "Customer DrgNo"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.Cust_Drgno, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                .ColHidden = True

                ''- Customer Drawing No .
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.Cust_DrgNo_Desc
                .Text = "Customer DrgNo Desc"
                .set_ColWidth(EnumInv.Cust_DrgNo_Desc, 1500)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5


                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "Customer DrgNo Desc"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.Cust_DrgNo_Desc, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                .ColHidden = True



                ''- PREV SO QTY
                .Row = FPSpreadADO.CoordConstants.SpreadHeader
                .Col = EnumInv.PrevQty
                .Text = "Prev Qty"
                .set_ColWidth(EnumInv.PrevQty, 1500)
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5


                .Row = FPSpreadADO.CoordConstants.SpreadHeader + 1
                .Text = "Prev Qty"
                .FontName = "Arial"
                .FontBold = True
                .FontSize = 8.5
                .AddCellSpan(EnumInv.PrevQty, FPSpreadADO.CoordConstants.SpreadHeader, 1, 2)
                .ColHidden = True



            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try

        'ErrHandler:
        '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
#End Region



    '    Private Sub PrintInvoice()
    '        Dim strSql As String
    '        Dim Sqlcmd As New SqlCommand
    '        Dim SQLCon As SqlConnection
    '        Dim ObjVal As Object
    '        Dim RdAddSold As ReportDocument
    '        Dim RepPath As String


    '        SQLCon = SqlConnectionclass.GetConnection()
    '        Sqlcmd.Connection = SQLCon
    '        Sqlcmd.CommandType = CommandType.Text
    '        Try
    '            Sqlcmd.CommandText = "SELECT REPORT_FILENAME FROM SALECONF WHERE UNIT_CODE='" + gstrUNITID + "' AND INVOICE_TYPE='INV' AND SUB_TYPE='T' AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE"
    '            ObjVal = Sqlcmd.ExecuteScalar()
    '            If IsNothing(ObjVal) = True Then ObjVal = String.Empty
    '            If ObjVal.ToString = String.Empty Then
    '                MsgBox("Report File Not Found In Sale Configuration.", MsgBoxStyle.Information, ResolveResString(100))
    '                Exit Sub
    '            End If

    '            Sqlcmd.CommandType = CommandType.StoredProcedure
    '            Sqlcmd.CommandText = "TRADING_INVOICE_PRINT"
    '            Sqlcmd.Parameters.Add("@UNITCODE", SqlDbType.VarChar).Value = gstrUNITID
    '            Sqlcmd.Parameters.Add("@IPADDRESS", SqlDbType.VarChar).Value = gstrIpaddressWinSck
    '            Sqlcmd.Parameters.Add("@INVOICENO", SqlDbType.Decimal).Value = Val(Me.txtChallanNo.Text).ToString
    '            Sqlcmd.ExecuteNonQuery()

    '            Dim Frm As New eMProCrystalReportViewer_Inv
    '            RdAddSold = Frm.GetReportDocument()
    '            ' RepPath = "C:\Documents and Settings\amitrana\Desktop\" + ObjVal.ToString.Trim & ".rpt"
    '            RepPath = My.Application.Info.DirectoryPath & "\Reports\" & ObjVal.ToString.Trim & ".rpt"
    '            RdAddSold.Load(RepPath)
    '            RdAddSold.DataDefinition.RecordSelectionFormula = "{TRADING_TMP_INVOICE_PRINT_HDR.UNIT_CODE}='" + gstrUNITID + "' And {TRADING_TMP_INVOICE_PRINT_HDR.IPADDRESS}='" + gstrIpaddressWinSck + "'"
    '            Dim c As New CrystalDecisions.Shared.PageMargins
    '            c.rightMargin = 0
    '            c.leftMargin = 0
    '            c.topMargin = 0

    '            RdAddSold.PrintOptions.ApplyPageMargins(c)
    '            Dim Section As Section              'Defining the section of report
    '            Dim objFieldObject As PictureObject 'For storing field object of report
    '            Dim intSectionCount As Integer      'Counter for Sections in report    
    '            Dim intSectionFieldCount As Integer 'Counter for fields in the section

    '            Try

    '                For intSectionCount = 0 To RdAddSold.ReportDefinition.Sections.Count - 1
    '                    ' Get the Section object by name.
    '                    Section = RdAddSold.ReportDefinition.Sections.Item(intSectionCount)
    '                    ' Get the ReportObject by name and cast it as a FieldObject.
    '                    For intSectionFieldCount = 0 To Section.ReportObjects.Count - 1
    '                        If Section.ReportObjects(intSectionFieldCount).Kind = ReportObjectKind.PictureObject Then
    '                            objFieldObject = Section.ReportObjects(intSectionFieldCount)
    '                            If objFieldObject.Name.ToUpper = "INVIMAGE" Then
    '                                objFieldObject.Height = 0
    '                            End If

    '                        End If

    '                    Next
    '                Next intSectionCount

    '            Catch ex As Exception
    '                MsgBox(ex.Message)
    '            End Try

    '            Frm.SetReportDocument()
    '            Frm.Show()
    '        Catch EX As Exception
    '            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
    '            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
    '        Finally
    '            If Sqlcmd.Connection.State = ConnectionState.Open Then Sqlcmd.Connection.Close()
    '            If SQLCon.State = ConnectionState.Open Then SQLCon.Close()
    '            Sqlcmd.Connection.Dispose()
    '            Sqlcmd.Dispose()
    '            SQLCon.Dispose()
    '        End Try
    '    End Sub
    Private Sub CmdGrpChEnt_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles CmdGrpChEnt.ButtonClick
        Try
            Select Case e.Button
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                    Me.Cmditems.Visible = False
                    BlankFields()
                    'FRMMKTTRN0076A.DeleteTmpTable()
                    Call SetPRGridHeading()
                    Me.Group1.Enabled = True
                    Me.Group2.Enabled = False
                    Me.Group3.Enabled = True
                    Me.Group4.Enabled = True
                    Me.CmdChallanNo.Enabled = False
                    Me.Cmditems.Enabled = True
                    Me.CmdCustCodeHelp.Enabled = True
                    Me.CmdRefNoHelp.Enabled = True
                    Me.txtRemarks.Text = String.Empty

                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = True
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
                    'Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                    '    Me.Group1.Enabled = True
                    '    Me.Group2.Enabled = True
                    '    Me.Group3.Enabled = True
                    '    Me.Group4.Enabled = True
                    '    Me.Cmditems.Enabled = False
                    '    Me.CmdCustCodeHelp.Enabled = False
                    '    Me.CmdRefNoHelp.Enabled = False
                    '    Me.CmdChallanNo.Enabled = False


                    '    If blnBillFlag = True Then ' IF INVOICE IS LOCKED
                    '        Me.Group1.Enabled = False
                    '        Me.Group2.Enabled = False
                    '        Me.Group3.Enabled = False
                    '        Me.SpChEntry.Row = 1
                    '        Me.SpChEntry.Col = EnumInv.BINQTY
                    '        Me.SpChEntry.Lock = True
                    '    Else
                    '        Me.SpChEntry.Row = 1
                    '        Me.SpChEntry.Col = EnumInv.BINQTY
                    '        Me.SpChEntry.Lock = False
                    '    End If


                    '    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = True
                    '    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    '    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    '    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    '    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = False
                    '    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                    If MsgBox("Are you Sure? ", vbYesNo, "eMPro") = vbYes Then
                    Else
                        Exit Sub
                    End If
                    If SaveData() = True Then
                        Me.Group1.Enabled = False
                        Me.Group2.Enabled = False
                        Me.Group3.Enabled = False
                        Me.Group4.Enabled = False
                        Me.CmdChallanNo.Enabled = True
                        CmdGrpChEnt.Revert()
                        'CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                        'CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                        'CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                        'CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                        'CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                        'CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = False


                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = False



                        MsgBox("Invoice Saved Successfully." + vbCrLf + "Invoice No-" + txtChallanNo.Text.ToString, MsgBoxStyle.Information, ResolveResString(100))
                    End If
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                    BlankFields()
                    Call SetPRGridHeading()
                    Me.Group1.Enabled = False
                    Me.Group2.Enabled = False
                    Me.Group3.Enabled = False
                    Me.Group4.Enabled = False
                    Me.CmdChallanNo.Enabled = True
                    Me.Cmditems.Enabled = False
                    Me.CmdCustCodeHelp.Enabled = False
                    Me.CmdRefNoHelp.Enabled = False
                    CmdGrpChEnt.Revert()

                    'CmdGrpChEnt.Top = 578
                    'CmdGrpChEnt.Left = 201

                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = False
                    'Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
                    '    If MsgBox("Are You Sure To Delete Select Invoice?", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                    '        If Me.txtChallanNo.Text.Trim = String.Empty Then
                    '            MsgBox("Please Select Invoice Number To Delete.", MsgBoxStyle.Information, ResolveResString(100))
                    '            Exit Sub
                    '        End If
                    '        If DeleteData() = True Then
                    '            CmdGrpChEnt_ButtonClick(Sender, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL))
                    '        End If
                    '    End If
                    'Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
                    '   PrintInvoice()
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                    Me.Close()
            End Select
        Catch Ex As Exception
            MsgBox(Ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub
    '    Private Sub cmdCreditHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCreditHelp.Click
    '        Try
    '            Dim strHelp() As String
    '            Dim strQuery As String
    '            strQuery = "SELECT CRTRM_TERMID,CRTRM_DESC FROM GEN_CREDITTRMMASTER WHERE UNIT_CODE='" + gstrUNITID + "'  AND CRTRM_STATUS=1 "
    '            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Credit Terms")
    '            If UBound(strHelp) > 0 Then
    '                If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
    '                    MsgBox("Credit Terms Found.", MsgBoxStyle.Information, ResolveResString(100))
    '                    Exit Sub
    '                End If
    '                txtCreditTerms.Text = strHelp(0)
    '                Me.lblCreditTermDesc.Text = strHelp(1)
    '            End If
    '        Catch Ex As Exception
    '            MsgBox(Ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
    '        Finally
    '        End Try
    '    End Sub
    '    Private Sub Cmditems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmditems.Click
    '        Dim frmMKTTRN0021A As New frmMKTTRN0021A
    '        Dim frmMKTTRN0021B As New frmMKTTRN0021B
    '        Dim frmMKTTRN0021 As New frmMKTTRN0021
    '        On Error GoTo ErrHandler
    '        Dim rssalechallan As ClsResultSetDB
    '        Dim salechallan As String
    '        Dim strItemNotIn As String
    '        Dim varItemCode As Object
    '        Dim rsSaleConf As ClsResultSetDB
    '        Dim strStockLocation As String
    '        Dim rsCurrencyType As ClsResultSetDB
    '        Dim intLoopCounter As Short
    '        Dim intMaxLoop As Short
    '        Dim blnTrfInvoiceWithSO As Boolean
    '        With Me.SpChEntry
    '            If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
    '                .MaxRows = 1
    '                .Row = 1 : .Row2 = .MaxRows : .Col = EnumInv.ENUMITEMCODE : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
    '            End If
    '        End With

    '        Dim strQry As String
    '        Dim rsSOReq As ClsResultSetDB
    '        Dim strtemp() As String
    '        frmMKTTRN0021.IsTradingInVoice = True
    '        Select Case Me.CmdGrpChEnt.Mode
    '            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
    '                rssalechallan = New ClsResultSetDB
    '                salechallan = ""
    '                salechallan = "SELECT Invoice_type,SUB_CATEGORY FROM saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_No = "
    '                salechallan = salechallan & Val(txtChallanNo.Text)
    '                rssalechallan.GetResult(salechallan, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
    '                If rssalechallan.GetNoRows > 0 Then
    '                    rssalechallan.MoveFirst()
    '                    strInvType = rssalechallan.GetValue("Invoice_type")
    '                    strInvSubType = rssalechallan.GetValue("sub_category")
    '                End If
    '                rssalechallan.ResultSetClose()
    '                strStockLocation = ""
    '                strStockLocation = StockLocationSalesConf(strInvType, strInvSubType, "TYPE")
    '                mstrLocationCode = Trim(strStockLocation)
    '                If (UCase(strInvType) = "SRC") And mblnServiceInvoiceWithoutSO Then
    '                    mstrItemCode = frmMKTTRN0021.SelectDatafromsaleDtl(Trim(txtChallanNo.Text))
    '                    If Len(Trim(mstrItemCode)) = 0 Then
    '                        SpChEntry.MaxRows = 0
    '                        frmMKTTRN0021 = Nothing
    '                    End If
    '                Else
    '                    If Len(Trim(strStockLocation)) > 0 Then

    '                        If (UCase(strInvType) = "INV") Or (UCase(strInvType) = "EXP") Or (UCase(strInvType) = "SRC") Then
    '                            mstrItemCode = frmMKTTRN0021.SelectDatafromsaleDtl(Trim(txtChallanNo.Text))
    '                            If Len(Trim(mstrItemCode)) = 0 Then
    '                                SpChEntry.MaxRows = 0
    '                                frmMKTTRN0021 = Nothing
    '                            End If
    '                        Else
    '                            mstrItemCode = frmMKTTRN0021.SelectDatafromsaleDtl(Trim(txtChallanNo.Text))
    '                            If Len(Trim(mstrItemCode)) = 0 Then
    '                                SpChEntry.MaxRows = 0
    '                                frmMKTTRN0021 = Nothing
    '                            End If
    '                        End If
    '                    Else
    '                        MsgBox("Please Define Stock Location in Sales Conf")
    '                        Exit Sub
    '                    End If
    '                End If
    '            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
    '                If mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
    '                    If MsgBox(" Do You Want To Follow FIFO Wise Batch Tracking ", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
    '                        mblnbatchfifomode = True
    '                    Else
    '                        mblnbatchfifomode = False
    '                    End If
    '                End If

    '                rsSOReq = New ClsResultSetDB
    '                strQry = "Select isnull(SORequired,0) as SORequired from saleConf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type ='TRF' and Sub_Type_Description='" & Trim(CmbInvSubType.Text) & "' and  (fin_start_date <= getdate() and fin_end_date >= getdate())"
    '                rsSOReq.GetResult(strQry)
    '                If rsSOReq.GetNoRows > 0 Then
    '                    blnTrfInvoiceWithSO = rsSOReq.GetValue("SORequired")
    '                Else
    '                    blnTrfInvoiceWithSO = False
    '                End If
    '                rsSOReq.ResultSetClose()
    '                If UCase(CStr((Trim(CmbInvType.Text)) = "NORMAL INVOICE")) Or (UCase(CStr(Trim(CmbInvType.Text) = "TRANSFER INVOICE")) And blnTrfInvoiceWithSO) Or UCase(CStr((Trim(CmbInvType.Text)) = "JOBWORK INVOICE")) Or UCase(CStr((Trim(CmbInvType.Text)) = "EXPORT INVOICE")) Or (UCase(CStr((Trim(CmbInvType.Text)) = "SERVICE INVOICE")) And Not mblnServiceInvoiceWithoutSO) Then
    '                    If (UCase(CmbInvSubType.Text) <> "SCRAP" And mblnMultipleSOAllowed = False) Then
    '                        If Len(Trim(txtRefNo.Text)) = 0 Then
    '                            Call ConfirmWindow(10240, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
    '                            If txtRefNo.Enabled Then txtRefNo.Focus()
    '                            Exit Sub
    '                        ElseIf Len(Trim(txtAmendNo.Text)) = 0 Then
    '                            'User Can Enter Ref Code And Amendment From KeyBoard 1.Check If Ref No with Blank Amend is Over Or NOT
    '                            '   2.If Over Then see y No Amendments are added
    '                            If OriginalRefNoOVER(Trim(txtRefNo.Text)) Then
    '                                'Orig Ref No is OVER , So Amendment Number should be added
    '                                MsgBox("Enter Amendment No.", MsgBoxStyle.Information, ResolveResString(100))
    '                                If txtAmendNo.Enabled Then txtAmendNo.Focus() : Exit Sub
    '                            Else
    '                                'Original Ref No is Still Active
    '                                mstrRefNo = Trim(txtRefNo.Text)
    '                                mstrAmmNo = "" 'As Amend No is Blank
    '                            End If
    '                        Else
    '                            'Reference Number is Already Verfied in Validate Event of txtRefNo Amend No. is Also Verified in Its Validate Event
    '                            'Then Pass It to Form variables
    '                            mstrRefNo = Trim(txtRefNo.Text)
    '                            mstrAmmNo = Trim(txtAmendNo.Text)
    '                        End If
    '                    End If
    '                End If

    '                If SpChEntry.MaxRows > 0 Then
    '                    intMaxLoop = SpChEntry.MaxRows
    '                    strItemNotIn = ""
    '                    For intLoopCounter = 1 To intMaxLoop
    '                        With SpChEntry
    '                            varItemCode = Nothing
    '                            Call .GetText(EnumInv.ENUMITEMCODE, intLoopCounter, varItemCode)
    '                            If Len(Trim(strItemNotIn)) > 0 Then
    '                                strItemNotIn = Trim(strItemNotIn) & ",'" & Trim(varItemCode) & "'"
    '                            Else
    '                                strItemNotIn = "'" & Trim(varItemCode) & "'"
    '                            End If
    '                        End With
    '                    Next
    '                End If
    '                If False = True And 2 = 1 Then
    '                    ' to be removed
    '                Else
    '                    If UCase(CStr(Trim(CmbInvType.Text))) = "NORMAL INVOICE" Or UCase(CStr(Trim(CmbInvType.Text))) = "EXPORT INVOICE" Or UCase(CStr(Trim(CmbInvType.Text))) = "SERVICE INVOICE" Then
    '                        If UCase(Trim(CmbInvType.Text)) = "SERVICE INVOICE" And mblnServiceInvoiceWithoutSO Then
    '                            'If Len(Trim(strItemNotIn)) > 0 Then
    '                            '    mstrItemCode = frmMKTTRN0021.SelectDatafromItem_Mst(Trim(CmbInvType.Text), Trim(CmbInvSubType.Text), strStockLocation, , strItemNotIn, SpChEntry.MaxRows)
    '                            'Else
    '                            '    mstrItemCode = frmMKTTRN0021.SelectDatafromItem_Mst(Trim(CmbInvType.Text), Trim(CmbInvSubType.Text), strStockLocation)
    '                            'End If
    '                        Else
    '                            strStockLocation = ""
    '                            strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
    '                            mstrLocationCode = strStockLocation
    '                            If Len(Trim(strStockLocation)) > 0 Then

    '                                If Len(Trim(strItemNotIn)) > 0 Then
    '                                    mstrItemCode = frmMKTTRN0021.SelectDataFromCustOrd_Dtl(Trim(txtCustCode.Text), Trim(txtRefNo.Text), mstrAmmNo, Trim(CmbInvSubType.Text), Trim(CmbInvType.Text), strStockLocation, strItemNotIn, SpChEntry.MaxRows)
    '                                Else
    '                                    mstrItemCode = frmMKTTRN0021.SelectDataFromCustOrd_Dtl(Trim(txtCustCode.Text), Trim(txtRefNo.Text), mstrAmmNo, Trim(CmbInvSubType.Text), Trim(CmbInvType.Text), strStockLocation)
    '                                End If
    '                                BlankTaxDetails()

    '                            Else
    '                                MsgBox("Please Define Stock Location in Sales Conf", MsgBoxStyle.Information, ResolveResString(100))
    '                                Exit Sub
    '                            End If
    '                            If Len(Trim(mstrItemCode)) = 0 Then SpChEntry.MaxRows = 0
    '                        End If

    '                    Else
    '                        rsSaleConf = New ClsResultSetDB
    '                        rsSaleConf.GetResult("select Stock_Location From saleconf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Description ='" & Trim(CmbInvType.Text) & "' and sub_type_description ='" & Trim(CmbInvSubType.Text) & "' AND Location_Code ='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
    '                        If ((Len(Trim(rsSaleConf.GetValue("Stock_Location"))) = 0) Or (Trim(CStr(rsSaleConf.GetValue("Stock_Location") = "Unknown")))) Then
    '                            MsgBox("Plese Select Stock Location in SalesConf first", MsgBoxStyle.Information, ResolveResString(100))
    '                            If Cmditems.Enabled Then Cmditems.Focus()
    '                            Exit Sub
    '                        End If
    '                        mstrLocationCode = rsSaleConf.GetValue("Stock_Location")
    '                        rsSaleConf.ResultSetClose()
    '                        If Len(Trim(mstrItemCode)) = 0 And Len(Trim(strItemNotIn)) = 0 Then
    '                            SpChEntry.MaxRows = 0
    '                        Else
    '                            If Len(Trim(mstrItemCode)) = 0 Then
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '        End Select
    '        Dim intDecimalPlace As Short
    '        Dim strCurrency As String
    '        If Len(mstrItemCode) > 0 Then
    '            mstrItemCode = Mid(mstrItemCode, 1, Len(mstrItemCode) - 1)
    '            Select Case Me.CmdGrpChEnt.Mode
    '                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
    '                    rsCurrencyType = New ClsResultSetDB
    '                    rsCurrencyType.GetResult("Select Currency_code from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_No = " & Val(txtChallanNo.Text))
    '                    If rsCurrencyType.GetNoRows > 0 Then
    '                        rsCurrencyType.MoveFirst()
    '                        strCurrency = rsCurrencyType.GetValue("Currency_code")
    '                    End If
    '                    rsCurrencyType.ResultSetClose()
    '                    intDecimalPlace = ToGetDecimalPlaces(strCurrency)
    '                    If intDecimalPlace < 2 Then
    '                        intDecimalPlace = 2
    '                    End If
    '                    DisplayDetailsInSpread(strCurrency) 'Procedure Call To Select Data >From Sales_Dtl
    '                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
    '                    Call displayDeatilsfromCustOrdHdrandDtl()
    '                    System.Windows.Forms.Application.DoEvents()
    '            End Select
    '            If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
    '                If CDbl(txtChallanNo.Text.Trim.Substring(0, 2)) = 99 Then
    '                    Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
    '                    Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
    '                End If
    '            End If

    '        Else
    '            frmMKTTRN0021 = Nothing
    '        End If
    '        'Set Cell Type In Spread
    '        Call ChangeCellTypeStaticText()
    '        Call GetItemDescription()
    '        Exit Sub
    'ErrHandler:  'The Error Handling Code Starts here
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    '        Exit Sub
    '    End Sub

    '    Private Sub txtSaleTaxType_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs)
    '        '        Dim Cancel As Boolean = e.Cancel
    '        '        Dim strInvoiceType, strsql As String
    '        '        Dim rsChallanEntry, rsadditionaltax, rsadditionalVattax, rsadditionalsurcharge, rsadditionalVatsurcharge As ClsResultSetDB
    '        '        On Error GoTo ErrHandler
    '        '        If Len(txtSaleTaxType.Text) > 0 Then
    '        '            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
    '        '                strInvoiceType = UCase(Trim(CmbInvType.Text))
    '        '            ElseIf (CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT) Or (CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW) Then
    '        '                rsChallanEntry = New ClsResultSetDB
    '        '                rsChallanEntry.GetResult("Select a.Description,a.Sub_Type_Description from SaleConf a,SalesChallan_Dtl b where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND Doc_No = " & txtChallanNo.Text & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and a.Location_code = b.Location_code and (fin_start_date <= getdate() and fin_end_date >= getdate())")
    '        '                strInvoiceType = UCase(rsChallanEntry.GetValue("Description"))
    '        '                rsChallanEntry.ResultSetClose()
    '        '            End If
    '        '            If UCase(Trim(strInvoiceType)) <> "SERVICE INVOICE" Then
    '        '                If UCase(Trim(strInvoiceType)) <> "JOBWORK INVOICE" Then
    '        '                    If CheckExistanceOfFieldData((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST' OR Tx_TaxeID='VAT')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
    '        '                        lblSaltax_Per.Text = CStr(GetTaxRate((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST' OR Tx_TaxeID='VAT')"))

    '        '                    Else
    '        '                        Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
    '        '                        Cancel = True
    '        '                        txtSaleTaxType.Text = ""
    '        '                        If txtSaleTaxType.Enabled Then txtSaleTaxType.Focus()
    '        '                    End If
    '        '                Else
    '        '                    If CheckExistanceOfFieldData((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST' Or Tx_TaxeID='SRT' OR Tx_TaxeID='VAT')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
    '        '                        lblSaltax_Per.Text = CStr(GetTaxRate((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST' Or Tx_TaxeID='SRT' OR Tx_TaxeID='VAT')"))

    '        '                    Else
    '        '                        Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
    '        '                        Cancel = True
    '        '                        txtSaleTaxType.Text = ""
    '        '                        If txtSaleTaxType.Enabled Then txtSaleTaxType.Focus()
    '        '                    End If
    '        '                End If
    '        '            Else
    '        '                If CheckExistanceOfFieldData((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='SRT')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
    '        '                    lblSaltax_Per.Text = CStr(GetTaxRate((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='SRT')"))

    '        '                Else
    '        '                    Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
    '        '                    Cancel = True
    '        '                    txtSaleTaxType.Text = ""
    '        '                    If txtSaleTaxType.Enabled Then txtSaleTaxType.Focus()
    '        '                End If
    '        '            End If
    '        '        End If
    '        '        If UCase(Trim(GetPlantName)) = "MATM" And UCase(strInvoiceType) = "NORMAL INVOICE" Then
    '        '            strsql = " select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  (Tx_TaxeID='CST' OR Tx_TaxeID='LST') and txrt_percentage > 2.0 and TxRt_Rate_No='" & txtSaleTaxType.Text & "'  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date)) "
    '        '            rsadditionaltax = New ClsResultSetDB
    '        '            rsadditionaltax.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
    '        '            If rsadditionaltax.GetNoRows > 0 Then
    '        '                rsadditionalsurcharge = New ClsResultSetDB
    '        '                strsql = " select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  Tx_TaxeID='SsT' and txrt_percentage=5.0"
    '        '                rsadditionalsurcharge.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
    '        '                If rsadditionalsurcharge.GetNoRows > 0 Then
    '        '                    'txtSurchargeTaxType.Text = rsadditionalsurcharge.GetValue("TxRt_Rate_No")
    '        '                    'lblSurcharge_Per.Text = rsadditionalsurcharge.GetValue("TxRt_Percentage")
    '        '                End If
    '        '                rsadditionalsurcharge.ResultSetClose()
    '        '                rsadditionalsurcharge = Nothing
    '        '            End If
    '        '            rsadditionaltax.ResultSetClose()
    '        '            rsadditionaltax = Nothing
    '        '            strsql = " select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  (Tx_TaxeID='VAT') and TxRt_Rate_No='" & txtSaleTaxType.Text & "'  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
    '        '            rsadditionalVattax = New ClsResultSetDB
    '        '            rsadditionalVattax.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
    '        '            If rsadditionalVattax.GetNoRows > 0 Then
    '        '                rsadditionalVatsurcharge = New ClsResultSetDB
    '        '                strsql = " select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  Tx_TaxeID='SsT' and txrt_percentage=5.0  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
    '        '                rsadditionalVatsurcharge.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
    '        '                If rsadditionalVatsurcharge.GetNoRows > 0 Then
    '        '                    'txtSurchargeTaxType.Text = rsadditionalVatsurcharge.GetValue("TxRt_Rate_No")
    '        '                    'lblSurcharge_Per.Text = rsadditionalVatsurcharge.GetValue("TxRt_Percentage")
    '        '                End If
    '        '                rsadditionalVatsurcharge.ResultSetClose()
    '        '                rsadditionalVatsurcharge = Nothing
    '        '            End If
    '        '            rsadditionalVattax.ResultSetClose()
    '        '            rsadditionalVattax = Nothing
    '        '        End If
    '        '        GoTo EventExitSub
    '        'ErrHandler:  'The Error Handling Code Starts here
    '        '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    '        'EventExitSub:
    '        '        e.Cancel = Cancel
    '    End Sub

    '    Private Sub SpChEntry_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles SpChEntry.ButtonClicked
    '        Try
    '            If e.col = EnumInv.SelectGrin Then
    '                Dim FRM As New FRMMKTTRN0076A
    '                Dim dblItemRate As Double
    '                Dim dblBinQty As Double


    '                SpChEntry.Row = e.row
    '                SpChEntry.Col = EnumInv.RATE_PERUNIT
    '                dblItemRate = SpChEntry.Text

    '                SpChEntry.Col = EnumInv.ENUMITEMCODE
    '                FRM.strInternalPartNo = SpChEntry.Text

    '                SpChEntry.Col = EnumInv.CUSTPARTNO
    '                FRM.strCustPartNo = SpChEntry.Text

    '                FRM.strInternalPartDesc = Me.lblInternalPartDesc.Text
    '                FRM.strCustomerPartDesc = Me.lblCustPartDesc.Text

    '                FRM.strCurrentStockQty = Me.lblCurrentStock.Text

    '                If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
    '                    FRM.ParentFormOperationMode = "ADD"
    '                Else
    '                    FRM.ParentFormOperationMode = "EDIT"
    '                    FRM.strInvoiceNo = Me.txtChallanNo.Text
    '                End If
    '                FRM.blnBillFlag = blnBillFlag
    '                FRM.ShowDialog()

    '                If strGrinAllocationOKCancel = True Then
    '                    SpChEntry.Col = EnumInv.ENUMQUANTITY
    '                    SpChEntry.Text = dblGrinQuantityForSale.ToString
    '                    dblBasicValue = (dblGrinQuantityForSale * dblItemRate)
    '                    Me.lblBasicValue.Text = dblBasicValue.ToString("#############.00")
    '                    Me.lblAssValue.Text = dblBasicValue.ToString("#############.00")
    '                    IncludeDefaultTaxes()
    '                    CalculateTaxes()
    '                    SpChEntry_EditChange(Me, New AxFPSpreadADO._DSpreadEvents_EditChangeEvent(EnumInv.BINQTY, 1))
    '                    SpChEntry.Col = EnumInv.CUMULATIVEBOXES : SpChEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : SpChEntry.TypeFloatDecimalPlaces = 0 : SpChEntry.TypeFloatMin = CDbl("0.00") : SpChEntry.TypeFloatMax = CDbl("99999999999999.99") : SpChEntry.Lock = True
    '                    SpChEntry.Col = EnumInv.FROMBOX : SpChEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : SpChEntry.TypeFloatDecimalPlaces = 0 : SpChEntry.TypeFloatMin = CDbl("0.00") : SpChEntry.TypeFloatMax = CDbl("999999.99") : SpChEntry.Lock = True
    '                    SpChEntry.Col = EnumInv.TOBOX : SpChEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : SpChEntry.TypeFloatDecimalPlaces = 0 : SpChEntry.TypeFloatMin = CDbl("0.00") : SpChEntry.TypeFloatMax = CDbl("999999.99") : SpChEntry.Lock = True

    '                End If
    '            End If
    '        Catch Ex As Exception
    '            MsgBox(Ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
    '        End Try

    '    End Sub
    '    Private Sub SpChEntry_Change(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ChangeEvent) Handles SpChEntry.Change

    '    End Sub

    '    Private Sub txtECSSTaxType_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs)
    '        '        Dim Cancel As Boolean = e.Cancel
    '        '        On Error GoTo ErrHandler
    '        '        If Len(txtECSSTaxType.Text) > 0 Then
    '        '            If CheckExistanceOfFieldData((txtECSSTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='ECS')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
    '        '                lblECSStax_Per.Text = CStr(GetTaxRate((txtECSSTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='ECS')"))

    '        '                If OptDiscountValue.Enabled Then OptDiscountValue.Focus()


    '        '            Else
    '        '                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
    '        '                Cancel = True
    '        '                txtECSSTaxType.Text = ""
    '        '                If txtECSSTaxType.Enabled Then txtECSSTaxType.Focus()
    '        '            End If
    '        '        End If
    '        '        GoTo EventExitSub
    '        'ErrHandler:  'The Error Handling Code Starts here
    '        '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    '        'EventExitSub:
    '        '        e.Cancel = Cancel
    '    End Sub
    '    Private Sub CmdECSSTaxType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        Try
    '            Dim strHelp() As String
    '            Dim strQuery As String
    '            strQuery = "SELECT TXRT_RATE_NO,TXRT_PERCENTAGE FROM GEN_TAXRATE WHERE UNIT_CODE='" + gstrUNITID + "' AND TX_TAXEID='ECT' AND ((ISNULL(DEACTIVE_FLAG,0) <> 1))"
    '            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "E Cess")
    '            If IsNothing(strHelp) = True Then
    '                txtECSSTaxType.Text = ""
    '                lblECSStax_Per.Text = ""
    '                CalculateTaxes()
    '                Exit Sub
    '            End If

    '            If UBound(strHelp) > 0 Then
    '                If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
    '                    MsgBox("Tax Rate Not Found.", MsgBoxStyle.Information, ResolveResString(100))
    '                    Exit Sub
    '                End If
    '                txtECSSTaxType.Text = strHelp(0)
    '                lblECSStax_Per.Text = strHelp(1)
    '                CalculateTaxes()
    '            End If
    '        Catch Ex As Exception
    '            MsgBox(Ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
    '        Finally
    '        End Try
    '    End Sub
    '    Private Sub BtnHCess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        Try
    '            Dim strHelp() As String
    '            Dim strQuery As String
    '            strQuery = "SELECT TXRT_RATE_NO,TXRT_PERCENTAGE FROM GEN_TAXRATE WHERE UNIT_CODE='" + gstrUNITID + "' AND TX_TAXEID='ECSST' AND ((ISNULL(DEACTIVE_FLAG,0) <> 1))"
    '            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "H Cess")
    '            If IsNothing(strHelp) = True Then
    '                txtSECSSTaxType.Text = ""
    '                lblSECSStax_Per.Text = ""
    '                CalculateTaxes()
    '                Exit Sub
    '            End If

    '            If UBound(strHelp) > 0 Then
    '                If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
    '                    MsgBox("Tax Rate Not Found.", MsgBoxStyle.Information, ResolveResString(100))
    '                    Exit Sub
    '                End If
    '                txtSECSSTaxType.Text = strHelp(0)
    '                lblSECSStax_Per.Text = strHelp(1)
    '                CalculateTaxes()
    '            End If
    '        Catch Ex As Exception
    '            MsgBox(Ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
    '        Finally
    '        End Try
    '    End Sub

    '    Private Sub CmdSaleTaxType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        ' To Display Help From SaleTax Master
    '        On Error GoTo ErrHandler
    '        Dim strHelp() As String
    '        Dim rssalechallan, rsadditionaltax, rsadditionalVattax, rsadditionalsurcharge, rsadditionalVatsurcharge As ClsResultSetDB
    '        Dim salechallan, strsql As String
    '        Dim strInvoiceType As Object
    '        Select Case Me.CmdGrpChEnt.Mode
    '            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
    '                If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
    '                    rssalechallan = New ClsResultSetDB
    '                    salechallan = ""
    '                    salechallan = "select b.Description, b.Sub_type_Description from SalesChallan_dtl a,saleconf b where doc_no = " & Trim(txtChallanNo.Text)
    '                    salechallan = salechallan & " and a.Location_code = b.Location_code and a.unit_code=b.unit_code and a.unit_code='" + gstrUNITID + "' and a.Invoice_type = b.invoice_type and a.sub_category = b.Sub_type"
    '                    rssalechallan.GetResult(salechallan, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
    '                    If rssalechallan.GetNoRows > 0 Then
    '                        rssalechallan.MoveFirst()
    '                        strInvoiceType = rssalechallan.GetValue("Description")
    '                    End If
    '                    rssalechallan.ResultSetClose()
    '                Else
    '                    strInvoiceType = CmbInvType.Text
    '                End If



    '                Dim strQuery As String
    '                strQuery = "SELECT TXRT_RATE_NO,TXRT_PERCENTAGE FROM GEN_TAXRATE WHERE UNIT_CODE='" + gstrUNITID + "' And (TX_TAXEID='CSTT' OR TX_TAXEID='LSTT' OR TX_TAXEID='VATT')  AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) OR (CAST(GETDATE() AS DATE) <= DEACTIVE_DATE))"
    '                strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Sales Tax")

    '                If IsNothing(strHelp) = True Then
    '                    txtSaleTaxType.Text = ""
    '                    lblSaltax_Per.Text = ""
    '                    CalculateTaxes()
    '                    Exit Sub
    '                End If

    '                If UBound(strHelp) > 0 Then
    '                    If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
    '                        MsgBox("Tax Rate Not Found.", MsgBoxStyle.Information, ResolveResString(100))
    '                        Exit Sub
    '                    End If
    '                    txtSaleTaxType.Text = strHelp(0)
    '                    lblSaltax_Per.Text = strHelp(1)
    '                    CalculateTaxes()
    '                End If

    '        End Select
    '        Exit Sub
    'ErrHandler:  'The Error Handling Code Starts here
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    '        Exit Sub
    '    End Sub
    '    Private Sub OptDiscountValue_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
    '            If OptDiscountValue.Checked = True Then
    '                OptDiscountPercentage.Checked = False
    '            ElseIf OptDiscountPercentage.Checked = True Then
    '                OptDiscountValue.Checked = False
    '            End If
    '            If Me.OptDiscountPercentage.Checked = False And Me.OptDiscountValue.Checked = False Then
    '                txtDiscountAmt.Text = "0.00"

    '                txtDiscountAmt.Enabled = False
    '            Else
    '                txtDiscountAmt.Enabled = True
    '            End If
    '            CalculateTaxes()
    '        End If
    '    End Sub
    '    Private Sub txtDiscountAmt_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
    '        AllowNumericValueInTextBox(txtDiscountAmt, e)
    '    End Sub
    '    Private Sub txtDiscountAmt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        If Me.OptDiscountPercentage.Checked = False And Me.OptDiscountValue.Checked = False Then
    '            txtDiscountAmt.Enabled = False
    '        End If
    '        CalculateTaxes()
    '    End Sub

    '    Private Sub OptDiscountPercentage_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
    '            If OptDiscountValue.Checked = True Then
    '                OptDiscountPercentage.Checked = False
    '            ElseIf OptDiscountPercentage.Checked = True Then
    '                OptDiscountValue.Checked = False
    '            End If
    '            If Me.OptDiscountPercentage.Checked = False And Me.OptDiscountValue.Checked = False Then
    '                txtDiscountAmt.Text = "0.00"

    '                txtDiscountAmt.Enabled = False
    '            Else
    '                txtDiscountAmt.Enabled = True
    '            End If
    '            CalculateTaxes()
    '        End If
    '    End Sub

    Private Sub ctlInsurance_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ctlInsurance.KeyPress
        AllowNumericValueInTextBox(ctlInsurance, e)
    End Sub

    Private Sub ctlInsurance_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ctlInsurance.TextChanged
        If Val(ctlInsurance.Text) > 0 Then
            Call Calculate_GridTotal()
        End If

    End Sub

    Private Sub txtFreight_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFreight.KeyPress
        AllowNumericValueInTextBox(txtFreight, e)
    End Sub

    Private Sub txtFreight_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFreight.TextChanged
        If Val(txtFreight.Text) > 0 Then
            Call Calculate_GridTotal()
        End If
    End Sub
    '    Private Sub CmdChallanNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdChallanNo.Click
    '        Try
    '            Dim strHelp() As String
    '            Dim strQuery As String
    '            strQuery = "SELECT DOC_NO INVOICE_NO,CONVERT(CHAR(11),INVOICE_DATE,106) INVOICE_DATE,ACCOUNT_CODE CUSTOMER_CODE,CASE WHEN ISNULL(CANCEL_FLAG,0)=1 THEN 'CANCELLED' ELSE '' END as Status FROM SALESCHALLAN_DTL WHERE INVOICE_TYPE='INV' AND SUB_CATEGORY='T' AND UNIT_CODE='" + gstrUNITID + "'  ORDER BY ENT_DT DESC"
    '            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Invoice List")
    '            If IsNothing(strHelp) = True Then Exit Sub
    '            If UBound(strHelp) > 0 Then
    '                If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
    '                    MsgBox("No Any Saved Invoice Found.", MsgBoxStyle.Information, ResolveResString(100))
    '                    Exit Sub
    '                End If
    '                GetSavedData(strHelp(0))
    '            End If
    '        Catch Ex As Exception
    '            MsgBox(Ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
    '        Finally
    '        End Try
    '    End Sub
    '    Private Sub cmdAddVAT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        Dim strHelp As String
    '        Dim strSTaxHelp() As String
    '        On Error GoTo ErrHandler
    '        Select Case Me.CmdGrpChEnt.Mode
    '            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
    '                strHelp = "Select TxRT_Rate_No,TxRt_Percentage from Gen_taxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  Tx_TaxeID in('ADVAT','ADCST')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
    '                strSTaxHelp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, "Add. VAT/CST Tax Help")
    '                If IsNothing(strSTaxHelp) = True Then
    '                    txtAddVAT.Text = ""
    '                    lblAddVAT.Text = ""
    '                    CalculateTaxes()
    '                    Exit Sub
    '                End If
    '                If UBound(strSTaxHelp) <= 0 Then Exit Sub
    '                If strSTaxHelp(0) = "0" Then
    '                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtAddVAT.Text = "" : txtAddVAT.Focus() : Exit Sub
    '                Else
    '                    txtAddVAT.Text = strSTaxHelp(0)
    '                    lblAddVAT.Text = strSTaxHelp(1)
    '                End If
    '        End Select
    '        CalculateTaxes()
    '        Exit Sub
    'ErrHandler:  'The Error Handling Code Starts here
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    '        Exit Sub
    '    End Sub
    '    Private Sub SpChEntry_EditChange(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles SpChEntry.EditChange
    '        If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
    '            If e.col = EnumInv.BINQTY Then
    '                Dim strSalqQuantity As String
    '                Dim strBinQty As String

    '                SpChEntry.Row = e.row
    '                SpChEntry.Col = EnumInv.ENUMQUANTITY
    '                strSalqQuantity = SpChEntry.Text

    '                SpChEntry.Col = EnumInv.BINQTY
    '                strBinQty = SpChEntry.Text

    '                If VAL(strBinQty) > 0 Then
    '                    SpChEntry.Col = EnumInv.FROMBOX
    '                    SpChEntry.Text = "1"

    '                    SpChEntry.Col = EnumInv.TOBOX
    '                    SpChEntry.Text = Math.Ceiling(Val(strSalqQuantity) / Val(strBinQty))

    '                    SpChEntry.Col = EnumInv.CUMULATIVEBOXES
    '                    SpChEntry.Text = Math.Ceiling(Val(strSalqQuantity) / Val(strBinQty))
    '                Else
    '                    SpChEntry.Col = EnumInv.FROMBOX
    '                    SpChEntry.Text = "0"

    '                    SpChEntry.Col = EnumInv.TOBOX
    '                    SpChEntry.Text = "0"

    '                End If
    '            End If
    '        End If
    '    End Sub

    '    Private Sub CmdGrpChEnt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdGrpChEnt.Load

    '    End Sub

    '    Private Sub SpChEntry_Advance(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_AdvanceEvent) Handles SpChEntry.Advance

    '    End Sub

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
            Me.CmdRefNoHelp.Enabled = True
            Me.CmdRefNoHelp.Focus()
        End If
        rsCustMst = Nothing
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdRefNoHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdRefNoHelp.Click
        If Len(Trim(txtCustCode.Text)) = 0 Then
            MessageBox.Show("Select Customer First.", "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.txtCustCode.Focus()
            Exit Sub
        End If
        On Error GoTo ErrHandler
        Dim strQselect As String
        Dim SO_detail As String = String.Empty

        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If Len(Trim(txtCustCode.Text)) > 0 Then
                    strQselect = " SELECT  CUST_REF+','+AMENDMENT_NO AS SO_NO,CASE WHEN CAST(Order_date AS DATE)='1900-01-01' THEN '' ELSE CONVERT(VARCHAR(11), Order_date,106) END SO_DATE, CUST_REF,AMENDMENT_NO " & _
                                    " FROM CUST_ORD_HDR WHERE UNIT_CODE='" & gstrUNITID & "'  AND ACCOUNT_CODE='" & txtCustCode.Text.ToString.Trim & "' AND ACTIVE_FLAG='A' AND (EFFECT_DATE >=GETDATE() AND EFFECT_DATE<=VALID_DATE )"

                    SO_detail = GetDocumentNo(strQselect, "SO_No")
                    If SO_detail = Nothing Then
                        MessageBox.Show("Operation Cancelled", "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                        Exit Sub
                    End If
                    Dim Split() As String = SO_detail.Split("~")
                    If Split(0).Contains("No record found.") Then
                        MessageBox.Show("No Record Found", "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                        Exit Sub
                    End If
                    If Not String.IsNullOrEmpty(Split(0)) Then
                        txtRefNo.Text = Split(0).ToString
                        Dim SONoSPlit() As String = txtRefNo.Text.Split(",")
                        Call Fill_Grid(SONoSPlit(0).ToString, SONoSPlit(1).ToString, "ADD")  '--- FILL GRID
                        If IsDate(Split(1).ToString) Then
                            dtpDateSO.Text = Split(1).ToString
                        Else
                            dtpDateSO.Text = Now.Date.ToString
                        End If

                    Else
                        CmdRefNoHelp.Enabled = False
                        txtRefNo.Enabled = False
                        dtpDateSO.Text = Now.Date.ToString
                        Call ConfirmWindow(10225, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    End If

                Else
                    CmdRefNoHelp.Enabled = False
                    txtRefNo.Enabled = False
                    dtpDateSO.Text = Now.Date.ToString
                End If
        End Select
        If Len(Trim(txtCustCode.Text)) > 0 Then
            Me.txtRefNo.Focus()
        End If

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Function GetDocumentNo(ByVal Qselect As String, ByVal HelpFor As String) As String

        Dim Result As String = ""
        Dim strHelp() As String = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, Qselect)
        If UBound(strHelp) = -1 Then Exit Function
        If Not IsNothing(strHelp) AndAlso strHelp.Length = 0 Then
            Result = "No record found."
            Return Result
        ElseIf String.IsNullOrEmpty(strHelp(1)) Then
            Result = "No record found."
            Return Result
        Else
            If HelpFor = "SO_No" Then
                If IsNothing(strHelp(0)) Or IsNothing(strHelp(1)) Then
                    Result = "No record found."
                    Return Result
                End If
                Result = strHelp(0).ToString + "~" + strHelp(1).ToString
            End If
            If HelpFor = "PI_No" Then
                If IsNothing(strHelp(0)) Or IsNothing(strHelp(1)) Or IsNothing(strHelp(2)) Or IsNothing(strHelp(3)) Or IsNothing(strHelp(4)) Or IsNothing(strHelp(5)) Or IsNothing(strHelp(6)) Then
                    Result = "No record found."
                    Return Result
                End If
                Result = strHelp(0).ToString + "~" + strHelp(1).ToString + "~" + strHelp(2).ToString + "~" + strHelp(3).ToString + "~" + strHelp(4).ToString + "~" + strHelp(5).ToString + "~" + strHelp(6).ToString
            End If
            Return Result
        End If

    End Function
    Private Sub Fill_Grid(ByVal SO_NO As String, ByVal Amendment_No As String, ByVal Mode As String)
        Try
            Dim Qselect As String = String.Empty
            If Mode = "ADD" Then
                Qselect = " SELECT A.ITEM_CODE,B.DESCRIPTION,A.CUST_DRGNO,CUST_DRG_DESC,A.HSNSACCODE AS HSNNO,A.ISHSNORSAC  AS HSN_TYPE,ISNULL(A.RATE,0) AS RATE, " & _
                 " A.ORDER_QTY AS QUANTITY,CAST( ISNULL(A.RATE,0)*A.ORDER_QTY AS NUMERIC(18,4)) AS BASIC_VALUE" & _
                 " ,A.IGSTTXRT_TYPE AS IGST_TAX_TYPE,ISNULL((SELECT TXRT_PERCENTAGE  FROM GEN_TAXRATE(NOLOCK) WHERE UNIT_CODE=A.UNIT_CODE AND TXRT_RATE_NO=A.IGSTTXRT_TYPE ),0)AS IGST_TAX_PER " & _
                 " ,A.CGSTTXRT_TYPE AS CGST_TAX_TYPE,ISNULL((SELECT TXRT_PERCENTAGE  FROM GEN_TAXRATE(NOLOCK) WHERE UNIT_CODE=A.UNIT_CODE AND TXRT_RATE_NO=A.CGSTTXRT_TYPE ),0)AS CGST_TAX_PER  " & _
                 " ,A.SGSTTXRT_TYPE AS SGST_TAX_TYPE,ISNULL((SELECT TXRT_PERCENTAGE  FROM GEN_TAXRATE(NOLOCK) WHERE UNIT_CODE=A.UNIT_CODE AND TXRT_RATE_NO=A.SGSTTXRT_TYPE ),0)AS SGST_TAX_PER" & _
                 " ,A.UTGSTTXRT_TYPE AS UTGST_TAX_TYPE,ISNULL((SELECT TXRT_PERCENTAGE  FROM GEN_TAXRATE(NOLOCK) WHERE UNIT_CODE=A.UNIT_CODE AND TXRT_RATE_NO=A.UTGSTTXRT_TYPE ),0)AS UTGST_TAX_PER " & _
                 "  ,A.COMPENSATION_CESS AS CSESS_TAX_TYPE,ISNULL((SELECT TXRT_PERCENTAGE  FROM GEN_TAXRATE(NOLOCK) WHERE UNIT_CODE=A.UNIT_CODE AND TXRT_RATE_NO=A.COMPENSATION_CESS ),0)AS CSESS_TAX_PER " & _
                 " FROM  CUST_ORD_DTL(NOLOCK) A" & _
                 " INNER JOIN ITEM_MST(NOLOCK) B " & _
                 " ON A.UNIT_CODE=B.UNIT_CODE AND A.ITEM_CODE=B.ITEM_CODE" & _
                 " WHERE A.UNIT_CODE='" & gstrUNITID & "'  AND A.ACCOUNT_CODE='" & txtCustCode.Text & "'" & _
                 " AND A.ACTIVE_FLAG='A' AND A.CUST_REF='" & SO_NO.ToString & "' AND B.STATUS='A' AND RTRIM(LTRIM(A.AMENDMENT_NO))='" & Amendment_No.ToString & "'"
            End If
            If Mode = "SHOW" Then
                Qselect = "	   SELECT DISTINCT  B.Item_Code,C.Description,B.Cust_Item_Code,B.Cust_Item_Desc,HSNSACCODE AS HSNNO,ISHSNORSAC  AS HSN_TYPE,ISNULL(B.RATE,0) AS RATE" & _
                             " ,B.Sales_Quantity AS QUANTITY,B.Basic_Amount AS BASIC_VALUE,B.ACCESSIBLE_AMOUNT,B.IGSTTXRT_TYPE AS IGST_TAX_TYPE,B.IGST_PERCENT AS IGST_TAX_PER,B.IGST_AMT ,B.CGSTTXRT_TYPE," & _
                             " B.CGST_PERCENT ,B.CGST_AMT ,IGSTTXRT_TYPE,IGST_PERCENT,IGST_AMT ,UTGSTTXRT_TYPE,UTGST_PERCENT,UTGST_AMT ,SGSTTXRT_TYPE,SGST_PERCENT,SGST_AMT,COMPENSATION_CESS_TYPE AS CESS_TYPE," & _
                             "  COMPENSATION_CESS_PERCENT AS CESS_PERCENT,COMPENSATION_CESS_AMT  AS CESS_AMT,B.DISCOUNT_PERC,B.DISCOUNT_AMT,B.ITEM_VALUE,B.ADVANCE,B.ITEM_REMARK AS REMARKS,A.INSURANCE,A.FRIEGHT_AMOUNT,A.Remarks  " & _
                             " FROM PI_SALESCHALLAN_DTL(NOLOCK) AS A" & _
                             " INNER JOIN PI_SALES_DTL (NOLOCK) AS B" & _
                             " ON A.UNIT_CODE=B.UNIT_CODE AND A.CUST_REF=B.CUST_REF AND A.AMENDMENT_NO=B.AMENDMENT_NO AND A.DOC_NO=B.DOC_NO " & _
                             " INNER JOIN ITEM_MST(NOLOCK) C" & _
                             " ON B.UNIT_CODE=C.UNIT_CODE AND B.ITEM_CODE=C.ITEM_CODE" & _
                             " WHERE B.UNIT_CODE='" & gstrUNITID & "'  AND B.Doc_No='" & txtChallanNo.Text & "'"
            End If



            Dim da As New SqlDataAdapter(Qselect, SqlConnectionclass.GetConnection)
            Dim dt As New DataTable
            da.Fill(dt)
            If dt.Rows.Count = 0 Then
                MessageBox.Show("No Record Found for this SO.", "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                If Mode = "ADD" Then
                    Call PopulateGridData(dt)
                End If
                If Mode = "SHOW" Then
                    Call PopulateGridData_ONSHOW(dt)
                    Me.ctlInsurance.Text = Val(dt.Rows(0)("INSURANCE").ToString)
                    Me.txtFreight.Text = Val(dt.Rows(0)("FRIEGHT_AMOUNT").ToString)
                    Me.txtRemarks.Text = dt.Rows(0)("REMARKS1").ToString
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub
    Private Sub PopulateGridData(ByVal dtRec As DataTable)
        If dtRec.Rows.Count > 0 Then
            For i As Integer = 0 To dtRec.Rows.Count - 1
                With sspr
                    .MaxRows = i + 1

                    .Row = .MaxRows : .Col = EnumInv.SL : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 15 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .Text = i + 1

                    .Row = .MaxRows : .Col = EnumInv.ItemCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 30 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("ITEM_CODE").ToString <> "", dtRec.Rows(i)("ITEM_CODE").ToString, String.Empty)
                    .Row = .MaxRows : .Col = EnumInv.HSN_SAC_No : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 30 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("HSNNO").ToString <> "", dtRec.Rows(i)("HSNNO").ToString, String.Empty)
                    .Row = .MaxRows : .Col = EnumInv.HSN_SAC_Type : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 30 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("HSN_TYPE").ToString <> "", dtRec.Rows(i)("HSN_TYPE").ToString, String.Empty)
                    .Row = .MaxRows : .Col = EnumInv.Quantity : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = False : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = IIf(dtRec.Rows(i)("QUANTITY").ToString <> "0", dtRec.Rows(i)("QUANTITY").ToString, 0)
                    .Row = .MaxRows : .Col = EnumInv.Rate : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = IIf(dtRec.Rows(i)("RATE").ToString <> "0", dtRec.Rows(i)("RATE").ToString, 0)
                    .Row = .MaxRows : .Col = EnumInv.Basic_value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = IIf(dtRec.Rows(i)("BASIC_VALUE").ToString <> "0", dtRec.Rows(i)("BASIC_VALUE").ToString, 0)
                    .Row = .MaxRows : .Col = EnumInv.DiscountPer : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = False : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = 0.0
                    .Row = .MaxRows : .Col = EnumInv.DiscountVal : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = 0.0
                    .Row = .MaxRows : .Col = EnumInv.Assable_Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = dtRec.Rows(i)("BASIC_VALUE").ToString
                    .Row = .MaxRows : .Col = EnumInv.Advance_Amt : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = False : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = 0
                    '--IGST
                    .Row = .MaxRows : .Col = EnumInv.IGST_Tax_type : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 15 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("IGST_TAX_TYPE").ToString <> "", dtRec.Rows(i)("IGST_TAX_TYPE").ToString, String.Empty)
                    .Row = .MaxRows : .Col = EnumInv.IGST_Tax_Per : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = dtRec.Rows(i)("IGST_TAX_PER").ToString
                    .Row = .MaxRows : .Col = EnumInv.IGST_Tax_Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = 0
                    '--CGST
                    .Row = .MaxRows : .Col = EnumInv.CGST_Tax_type : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 15 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("CGST_TAX_TYPE").ToString <> "", dtRec.Rows(i)("CGST_TAX_TYPE").ToString, String.Empty)
                    .Row = .MaxRows : .Col = EnumInv.CGST_Tax_Per : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = IIf(dtRec.Rows(i)("CGST_TAX_PER").ToString <> "", dtRec.Rows(i)("CGST_TAX_PER").ToString, 0.0)
                    .Row = .MaxRows : .Col = EnumInv.CGST_Tax_Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = 0
                    '--SGST
                    .Row = .MaxRows : .Col = EnumInv.SGST_Tax_type : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 15 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("SGST_TAX_TYPE").ToString <> "", dtRec.Rows(i)("SGST_TAX_TYPE").ToString, String.Empty)
                    .Row = .MaxRows : .Col = EnumInv.SGST_Tax_Per : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = IIf(dtRec.Rows(i)("SGST_TAX_PER").ToString <> "", dtRec.Rows(i)("SGST_TAX_PER").ToString, 0.0)
                    .Row = .MaxRows : .Col = EnumInv.SGST_Tax_Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = 0
                    '--UTGST
                    .Row = .MaxRows : .Col = EnumInv.UTGST_Tax_type : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 15 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("UTGST_TAX_TYPE").ToString <> "", dtRec.Rows(i)("UTGST_TAX_TYPE").ToString, String.Empty)
                    .Row = .MaxRows : .Col = EnumInv.UTGST_Tax_Per : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = dtRec.Rows(i)("UTGST_TAX_PER").ToString
                    .Row = .MaxRows : .Col = EnumInv.UTGST_Tax_Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = 0

                    '--UTGST
                    .Row = .MaxRows : .Col = EnumInv.CSESS_Tax_type : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 15 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("CSESS_TAX_TYPE").ToString <> "", dtRec.Rows(i)("CSESS_TAX_TYPE").ToString, String.Empty)
                    .Row = .MaxRows : .Col = EnumInv.CSESS_Tax_Per : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = dtRec.Rows(i)("CSESS_TAX_PER").ToString
                    .Row = .MaxRows : .Col = EnumInv.CSESS_Tax_Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = 0

                    '--- ITEMTOTAL
                    .Row = .MaxRows : .Col = EnumInv.ItemTotal : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = 0 ' IIf(dtRec.Rows(i)("BASIC_VALUE").ToString <> "", dtRec.Rows(i)("BASIC_VALUE").ToString, 0)
                    .Row = .MaxRows : .Col = EnumInv.Remarks : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 200 : .Lock = False : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .Text = String.Empty
                    '----INTERNAL ITEM DESCRIPTION
                    .Row = .MaxRows : .Col = EnumInv.Internal_Item_Desc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("DESCRIPTION").ToString <> "", dtRec.Rows(i)("DESCRIPTION").ToString, String.Empty)
                    '--- CUSTDRGNO DESC
                    .Row = .MaxRows : .Col = EnumInv.Cust_DrgNo_Desc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("CUST_DRG_DESC").ToString <> "", dtRec.Rows(i)("CUST_DRG_DESC").ToString, String.Empty)
                    '--- CUST DRGNO
                    .Row = .MaxRows : .Col = EnumInv.Cust_Drgno : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("CUST_DRGNO").ToString <> "", dtRec.Rows(i)("CUST_DRGNO").ToString, String.Empty)
                    '-- Prev Qty
                    .Row = .MaxRows : .Col = EnumInv.PrevQty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = False : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = IIf(dtRec.Rows(i)("QUANTITY").ToString <> "0", dtRec.Rows(i)("QUANTITY").ToString, 0)
                    .set_RowHeight(.MaxRows, 300)
                End With
            Next
            Call Calculate_GridTotal()
        End If
    End Sub

    Private Sub PopulateGridData_ONSHOW(ByVal dtRec As DataTable)
        If dtRec.Rows.Count > 0 Then
            For i As Integer = 0 To dtRec.Rows.Count - 1
                With sspr
                    .MaxRows = i + 1

                    .Row = .MaxRows : .Col = EnumInv.SL : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 15 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .Text = i + 1

                    .Row = .MaxRows : .Col = EnumInv.ItemCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 30 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("ITEM_CODE").ToString <> "", dtRec.Rows(i)("ITEM_CODE").ToString, String.Empty)
                    .Row = .MaxRows : .Col = EnumInv.HSN_SAC_No : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 30 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("HSNNO").ToString <> "", dtRec.Rows(i)("HSNNO").ToString, String.Empty)
                    .Row = .MaxRows : .Col = EnumInv.HSN_SAC_Type : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 30 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("HSN_TYPE").ToString <> "", dtRec.Rows(i)("HSN_TYPE").ToString, String.Empty)
                    .Row = .MaxRows : .Col = EnumInv.Quantity : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = IIf(dtRec.Rows(i)("QUANTITY").ToString <> "0", dtRec.Rows(i)("QUANTITY").ToString, 0)
                    .Row = .MaxRows : .Col = EnumInv.Rate : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = IIf(dtRec.Rows(i)("RATE").ToString <> "0", dtRec.Rows(i)("RATE").ToString, 0)
                    .Row = .MaxRows : .Col = EnumInv.Basic_value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = IIf(dtRec.Rows(i)("BASIC_VALUE").ToString <> "0", dtRec.Rows(i)("BASIC_VALUE").ToString, 0)
                    .Row = .MaxRows : .Col = EnumInv.DiscountPer : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = IIf(dtRec.Rows(i)("DISCOUNT_PERC").ToString <> "0", dtRec.Rows(i)("DISCOUNT_PERC").ToString, 0)
                    .Row = .MaxRows : .Col = EnumInv.DiscountVal : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = IIf(dtRec.Rows(i)("DISCOUNT_AMT").ToString <> "0", dtRec.Rows(i)("DISCOUNT_AMT").ToString, 0)
                    .Row = .MaxRows : .Col = EnumInv.Assable_Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = dtRec.Rows(i)("ACCESSIBLE_AMOUNT").ToString
                    .Row = .MaxRows : .Col = EnumInv.Advance_Amt : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = Val(dtRec.Rows(i)("ADVANCE").ToString)
                    '--IGST
                    .Row = .MaxRows : .Col = EnumInv.IGST_Tax_type : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 15 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("IGST_TAX_TYPE").ToString <> "", dtRec.Rows(i)("IGST_TAX_TYPE").ToString, String.Empty)
                    .Row = .MaxRows : .Col = EnumInv.IGST_Tax_Per : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = dtRec.Rows(i)("IGST_TAX_PER").ToString
                    .Row = .MaxRows : .Col = EnumInv.IGST_Tax_Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = dtRec.Rows(i)("IGST_AMT").ToString
                    '--CGST
                    .Row = .MaxRows : .Col = EnumInv.CGST_Tax_type : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 15 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("CGSTTXRT_TYPE").ToString <> "", dtRec.Rows(i)("CGSTTXRT_TYPE").ToString, String.Empty)
                    .Row = .MaxRows : .Col = EnumInv.CGST_Tax_Per : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = IIf(dtRec.Rows(i)("CGST_PERCENT").ToString <> "", dtRec.Rows(i)("CGST_PERCENT").ToString, 0.0)
                    .Row = .MaxRows : .Col = EnumInv.CGST_Tax_Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = dtRec.Rows(i)("CGST_AMT").ToString
                    '--SGST
                    .Row = .MaxRows : .Col = EnumInv.SGST_Tax_type : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 15 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("SGSTTXRT_TYPE").ToString <> "", dtRec.Rows(i)("SGSTTXRT_TYPE").ToString, String.Empty)
                    .Row = .MaxRows : .Col = EnumInv.SGST_Tax_Per : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = IIf(dtRec.Rows(i)("SGST_PERCENT").ToString <> "", dtRec.Rows(i)("SGST_PERCENT").ToString, 0.0)
                    .Row = .MaxRows : .Col = EnumInv.SGST_Tax_Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = IIf(dtRec.Rows(i)("SGST_AMT").ToString <> "", dtRec.Rows(i)("SGST_AMT").ToString, 0.0)
                    '--UTGST
                    .Row = .MaxRows : .Col = EnumInv.UTGST_Tax_type : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 15 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("UTGSTTXRT_TYPE").ToString <> "", dtRec.Rows(i)("UTGSTTXRT_TYPE").ToString, String.Empty)
                    .Row = .MaxRows : .Col = EnumInv.UTGST_Tax_Per : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = dtRec.Rows(i)("UTGST_PERCENT").ToString
                    .Row = .MaxRows : .Col = EnumInv.UTGST_Tax_Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = dtRec.Rows(i)("UTGST_AMT").ToString

                    '--UTGST
                    .Row = .MaxRows : .Col = EnumInv.CSESS_Tax_type : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 15 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("CESS_TYPE").ToString <> "", dtRec.Rows(i)("CESS_TYPE").ToString, String.Empty)
                    .Row = .MaxRows : .Col = EnumInv.CSESS_Tax_Per : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = dtRec.Rows(i)("CESS_PERCENT").ToString
                    .Row = .MaxRows : .Col = EnumInv.CSESS_Tax_Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = dtRec.Rows(i)("CESS_AMT").ToString

                    '--- ITEMTOTAL
                    .Row = .MaxRows : .Col = EnumInv.ItemTotal : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Text = IIf(dtRec.Rows(i)("ITEM_VALUE").ToString <> "", dtRec.Rows(i)("ITEM_VALUE").ToString, 0)
                    .Row = .MaxRows : .Col = EnumInv.Remarks : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 200 : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .Text = IIf(dtRec.Rows(i)("REMARKS").ToString <> "", dtRec.Rows(i)("REMARKS").ToString, String.Empty)
                    '----INTERNAL ITEM DESCRIPTION
                    .Row = .MaxRows : .Col = EnumInv.Internal_Item_Desc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("DESCRIPTION").ToString <> "", dtRec.Rows(i)("DESCRIPTION").ToString, String.Empty)
                    '--- CUSTDRGNO DESC
                    .Row = .MaxRows : .Col = EnumInv.Cust_DrgNo_Desc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("CUST_ITEM_DESC").ToString <> "", dtRec.Rows(i)("CUST_ITEM_DESC").ToString, String.Empty)
                    '--- CUST DRGNO
                    .Row = .MaxRows : .Col = EnumInv.Cust_Drgno : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .Lock = True : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .Text = IIf(dtRec.Rows(i)("CUST_ITEM_CODE").ToString <> "", dtRec.Rows(i)("CUST_ITEM_CODE").ToString, String.Empty)
                    .set_RowHeight(.MaxRows, 300)
                End With
            Next
            Call Calculate_GridTotal()
        End If
    End Sub

#Region "CALCULATE TAXATION AMOUNT"

    Private Sub Calculate_Grid_Values(ByVal intRow As Integer)
        Dim PrevQty_SO As Object
        Dim DiscountPer As Object
        Dim Advance_Amt As Object
        Dim BasicVal As Object
        Dim DiscountVal As Object
        Dim ItemTotal As Object
        Dim AccessibleValue As Object

        Dim IGST_Tax_Per As Object
        Dim CGST_Tax_Per As Object
        Dim SGST_Tax_Per As Object
        Dim UTGST_Tax_Per As Object
        Dim CSESS_Tax_Per As Object

        Dim IGST_Tax_Val As Object
        Dim CGST_Tax_Val As Object
        Dim SGST_Tax_Val As Object
        Dim UTGST_Tax_Val As Object
        Dim CSESS_Tax_Val As Object
        Dim Quantity As Object
        Dim Rate As Object
        Dim Basic_value As Object



        Try
            With sspr
                .Row = intRow
                PrevQty_SO = Nothing
                Quantity = Nothing
                Rate = Nothing
                Basic_value = Nothing
                DiscountPer = Nothing
                Advance_Amt = Nothing
                BasicVal = Nothing
                DiscountVal = Nothing
                ItemTotal = Nothing
                AccessibleValue = Nothing

                IGST_Tax_Per = Nothing
                CGST_Tax_Per = Nothing
                SGST_Tax_Per = Nothing
                UTGST_Tax_Per = Nothing
                CSESS_Tax_Per = Nothing

                IGST_Tax_Val = Nothing
                CGST_Tax_Val = Nothing
                SGST_Tax_Val = Nothing
                UTGST_Tax_Val = Nothing
                CSESS_Tax_Val = Nothing

                '--Prev SO Qty
                .Col = EnumInv.PrevQty
                PrevQty_SO = Val(.Text)
                '--- Quantity-
                .Col = EnumInv.Quantity
                Quantity = Val(.Text)

                If Val(Quantity) > Val(PrevQty_SO) Then
                    MessageBox.Show("Quatity can't be more than SO Quantity.", "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    .Col = EnumInv.Quantity
                    .Text = PrevQty_SO
                    Calculate_Grid_Values(.Row)     '--- Call Recersively when validation occurs to reset grid row again
                    Exit Sub
                End If
                '--- Quantity-
                .Col = EnumInv.Rate
                Rate = CDec(.Text)

                '--Basic Value
                .Col = EnumInv.Basic_value
                .Text = CDec(Quantity) * CDec(Rate)
                .GetText(EnumInv.Basic_value, .Row, BasicVal)
                If IsNothing(BasicVal) = True Then BasicVal = 0



                '-Dis %
                .Col = EnumInv.DiscountPer
                .GetText(EnumInv.DiscountPer, .Row, DiscountPer)
                If IsNothing(DiscountPer) = True Then DiscountPer = 0

                '-Dis Value  BASIC VAL * (DIS%\100)
                .Col = EnumInv.DiscountVal
                DiscountVal = CDec(BasicVal) * CDec(DiscountPer / 100)
                .SetText(EnumInv.DiscountVal, .Row, DiscountVal)
                .Text = DiscountVal
                '-Assable Value  ***  Assable_Value= VAL -DIS VALUE
                .Col = EnumInv.Assable_Value
                AccessibleValue = CDec(BasicVal) - CDec(DiscountVal)
                If Val(AccessibleValue) < 0 Then
                    MessageBox.Show("Discount Percentage can't be more than 100%", "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    .Col = EnumInv.DiscountPer
                    .Text = 0
                    Calculate_Grid_Values(.Row)     '--- Call Recersively when validation occurs to reset grid row again
                    Exit Sub
                End If
                .Text = AccessibleValue
                '-Advance Amount %
                .Col = EnumInv.Advance_Amt
                .GetText(EnumInv.Advance_Amt, .Row, Advance_Amt)
                If IsNothing(Advance_Amt) = True Then Advance_Amt = 0

                If Val(Advance_Amt) > Val(BasicVal) Then
                    MessageBox.Show("Advance Amount Can not be more than Basic Value.", "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    .Col = EnumInv.Advance_Amt
                    .Text = 0
                    .Col = EnumInv.ItemTotal
                    .Text = 0
                    .SetText(EnumInv.ItemTotal, .Row, 0)
                    Calculate_Grid_Values(.Row)     '--- Call Recersively when validation occurs to reset grid row again
                    Exit Sub
                End If
                '-IGST TAX % ***  BASIC VAL * (DIS%\100)
                .Col = EnumInv.IGST_Tax_Per
                .GetText(EnumInv.IGST_Tax_Per, .Row, IGST_Tax_Per)
                If IsNothing(IGST_Tax_Per) = True Then IGST_Tax_Per = 0
                .Text = IGST_Tax_Per
                '-IGST TAX Val ***  BASIC VAL * (DIS%\100)
                .Col = EnumInv.IGST_Tax_Value
                IGST_Tax_Val = CDec(Advance_Amt) * CDec(IGST_Tax_Per / 100)
                If IsNothing(IGST_Tax_Val) = True Then IGST_Tax_Val = 0
                .Text = IGST_Tax_Val

                '-CGST TAX % ***  BASIC VAL * (DIS%\100)
                .Col = EnumInv.CGST_Tax_Per
                .GetText(EnumInv.CGST_Tax_Per, .Row, CGST_Tax_Per)
                If IsNothing(CGST_Tax_Per) = True Then CGST_Tax_Per = 0

                '-CGST TAX Val ***  BASIC VAL * (DIS%\100)
                .Col = EnumInv.CGST_Tax_Value
                CGST_Tax_Val = CDec(Advance_Amt) * CDec(CGST_Tax_Per / 100)
                If IsNothing(CGST_Tax_Val) = True Then CGST_Tax_Val = 0
                .Text = CGST_Tax_Val

                '-SGST TAX % ***  BASIC VAL * (DIS%\100)
                .Col = EnumInv.SGST_Tax_Per
                .GetText(EnumInv.SGST_Tax_Per, .Row, SGST_Tax_Per)
                If IsNothing(SGST_Tax_Per) = True Then SGST_Tax_Per = 0

                '-SGST TAX Val ***  BASIC VAL * (DIS%\100)
                .Col = EnumInv.SGST_Tax_Value
                SGST_Tax_Val = CDec(Advance_Amt) * CDec(SGST_Tax_Per / 100)
                If IsNothing(SGST_Tax_Val) = True Then SGST_Tax_Val = 0
                .Text = SGST_Tax_Val

                '-UTGST TAX % ***  BASIC VAL * (DIS%\100)
                .Col = EnumInv.UTGST_Tax_Per
                .GetText(EnumInv.UTGST_Tax_Per, .Row, UTGST_Tax_Per)
                If IsNothing(UTGST_Tax_Per) = True Then UTGST_Tax_Per = 0

                '-UTGST TAX Val ***  BASIC VAL * (DIS%\100)
                .Col = EnumInv.UTGST_Tax_Value
                UTGST_Tax_Val = CDec(Advance_Amt) * CDec(UTGST_Tax_Per / 100)
                If IsNothing(UTGST_Tax_Val) = True Then UTGST_Tax_Val = 0
                .Text = UTGST_Tax_Val


                '-CSESS TAX % ***  BASIC VAL * (DIS%\100)
                .Col = EnumInv.CSESS_Tax_Per
                .GetText(EnumInv.CSESS_Tax_Per, .Row, CSESS_Tax_Per)
                If IsNothing(CSESS_Tax_Per) = True Then CSESS_Tax_Per = 0

                '-CSESS TAX Val ***  BASIC VAL * (DIS%\100)
                .Col = EnumInv.CSESS_Tax_Value
                CSESS_Tax_Val = CDec(Advance_Amt) * CDec(CSESS_Tax_Per / 100)
                If IsNothing(CSESS_Tax_Val) = True Then CSESS_Tax_Val = 0
                .Text = CSESS_Tax_Val

                '-Item Total ***  BASIC VAL * (DIS%\100)
                ItemTotal = CDec(Advance_Amt) + CDec(IGST_Tax_Val) + CDec(CGST_Tax_Val) + CDec(SGST_Tax_Val) + CDec(UTGST_Tax_Val) + CDec(CSESS_Tax_Val)
                .Col = EnumInv.ItemTotal
                If IsNothing(ItemTotal) = True Then ItemTotal = 0
                .Text = CDec(ItemTotal)
                'ItemTotal = CDec(AccessibleValue) + CDec(IGST_Tax_Val) + CDec(CGST_Tax_Val) + CDec(SGST_Tax_Val) + CDec(UTGST_Tax_Val) + CDec(CSESS_Tax_Val)
               
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
#End Region

    Private Sub sspr_EditChange(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles sspr.EditChange
        Try
            If (e.col = EnumInv.DiscountPer) Or (e.col = EnumInv.Advance_Amt) Or (e.col = EnumInv.Quantity) Then
                Calculate_Grid_Values(e.row)
            End If
            Call Calculate_GridTotal()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub Calculate_GridTotal()
        Try
            Dim ItemTotal As Decimal = 0
            Dim NetAdvAmt As Object = Nothing
            For i As Integer = 1 To sspr.MaxRows
                With sspr
                    .Row = i
                    .Col = EnumInv.ItemTotal
                    ItemTotal = ItemTotal + Val(.Text)

                    .Col = EnumInv.Advance_Amt
                    NetAdvAmt = NetAdvAmt + Val(.Text)
                End With
            Next
            LblNetInvoiceValue.Text = CDec(ItemTotal) - CDec(NetAdvAmt)  'CDec(ItemTotal + Val(ctlInsurance.Text) + Val(txtFreight.Text))
            lblRoundOff.Text = Math.Round(CDec(ItemTotal))
            lblNetAdvAmt.Text = CDec(NetAdvAmt)

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub sspr_ClickEvent(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles sspr.ClickEvent
        With sspr
            .Row = e.row
            .Col = EnumInv.Internal_Item_Desc
            lblInternalPartDesc.Text = .Text

            .Col = EnumInv.Cust_DrgNo_Desc
            lblCustPartDesc.Text = .Text

        End With
    End Sub
    Private Function SaveData() As Boolean
        Dim DocumentNo As String = String.Empty
        Dim cmd As SqlCommand = Nothing

        Try
            SaveData = False


            If Me.txtCustCode.Text.Trim = String.Empty Then
                MsgBox("Please Select The Customer.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            ElseIf Me.txtRefNo.Text.Trim = String.Empty Then
                MsgBox("Please Select The SO Number.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            ElseIf Me.sspr.MaxRows = 0 Then
                MsgBox("Please Select Item To Be Despatched.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            ElseIf Val(LblNetInvoiceValue.Text) = 0 Then
                MsgBox("Please Check The Invoice It's Value Can't Be Zero.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            End If

            sspr.Row = 1
            sspr.Col = EnumInv.ItemTotal
            If Val(sspr.Text) = 0 Then
                MsgBox("Item Quantity In Invoice Can't Be Zero.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            End If
            Dim ValidateItem_Code As Object
            Dim ValidateItem_total As Object
            For intCount As Integer = 1 To sspr.MaxRows
                With sspr
                    .Col = EnumInv.ItemCode
                    ValidateItem_Code = .Text

                    .Col = EnumInv.Quantity
                    If Val(.Text) = 0 Then
                        MessageBox.Show("Quantity Can not be Zero for Item-" & ValidateItem_Code & " at Row No-" & intCount, "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Function
                    End If
                    .Col = EnumInv.ItemTotal
                    ValidateItem_total = Val(.Text)
                    If Val(ValidateItem_total) = 0 Then
                        MessageBox.Show("Item Total Can not be Zero for Item-" & ValidateItem_Code & " at Row No-" & intCount, "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Function
                    End If
                End With
            Next


            'Sqlcmd.CommandText = "SELECT INVOICE_TYPE,SUB_TYPE FROM SALECONF WHERE UNIT_CODE='" + gstrUNITID + "' AND  DESCRIPTION ='" & Trim(CmbInvType.Text) & "'AND SUB_TYPE_DESCRIPTION ='" & Trim(CmbInvSubType.Text) & "' AND (FIN_START_DATE <= GETDATE() AND FIN_END_DATE >= GETDATE())"
            'DataRd = Sqlcmd.ExecuteReader()
            'If DataRd.HasRows Then
            '    DataRd.Read()
            '    strInvType = DataRd("INVOICE_TYPE")
            '    strInvSubType = DataRd("SUB_TYPE")
            'End If
            'If DataRd.IsClosed = False Then DataRd.Close()
            Dim SONoSPlit() As String = txtRefNo.Text.Split(",")
            SqlConnectionclass.CloseGlobalConnection()
            SqlConnectionclass.OpenGlobalConnection()

            cmd = New System.Data.SqlClient.SqlCommand()
            cmd.Connection = SqlConnectionclass.GetConnection
            cmd.Transaction = cmd.Connection.BeginTransaction(System.Data.IsolationLevel.Serializable)

            Dim QryExecuted As Boolean = False
            If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                Try
                    Dim CGST_VAL As Object
                    Dim IGST_VAL As Object
                    Dim SGST_VAL As Object
                    Dim UTGST_VAL As Object
                    Dim CSESS_VAL As Object
                    Dim ITEM_TOTAL As Object
                    Dim ADVANCE_AMT As Object
                    Dim Error_Msg As Object = String.Empty

                    For i As Integer = 1 To sspr.MaxRows
                        With sspr
                            .Row = i
                            .Col = EnumInv.CGST_Tax_Value
                            CGST_VAL = Val(CGST_VAL) + Val(.Text)

                            .Col = EnumInv.IGST_Tax_Value
                            IGST_VAL = Val(IGST_VAL) + Val(.Text)

                            .Col = EnumInv.SGST_Tax_Value
                            SGST_VAL = Val(SGST_VAL) + Val(.Text)

                            .Col = EnumInv.UTGST_Tax_Value
                            UTGST_VAL = Val(UTGST_VAL) + Val(.Text)

                            .Col = EnumInv.CSESS_Tax_Value
                            CSESS_VAL = Val(CSESS_VAL) + Val(.Text)

                            .Col = EnumInv.ItemTotal
                            ITEM_TOTAL = Val(ITEM_TOTAL) + Val(.Text)

                            .Col = EnumInv.Advance_Amt
                            ADVANCE_AMT = Val(ADVANCE_AMT) + Val(.Text)
                        End With
                    Next

                    With cmd
                        QryExecuted = False
                        .CommandText = String.Empty
                        .Parameters.Clear()
                        .CommandTimeout = 0
                        .CommandType = CommandType.StoredProcedure
                        .CommandText = "Proc_GST_Performa_Invoice"

                        .Parameters.Add("@Cust_Code", SqlDbType.VarChar, 30).Value = txtCustCode.Text.Trim.ToString
                        .Parameters.Add("@SO_No", SqlDbType.VarChar, 100).Value = SONoSPlit(0).ToString.Trim
                        .Parameters.Add("@Amendment_No", SqlDbType.VarChar, 200).Value = SONoSPlit(1).ToString.Trim
                        .Parameters.Add("@Insurance", SqlDbType.VarChar, 100).Value = Format(Double.Parse(ctlInsurance.Text), "#.00") ' Val(ctlInsurance.Text.ToString)
                        .Parameters.Add("@Freight", SqlDbType.VarChar, 100).Value = Format(Double.Parse(txtFreight.Text), "#.00") ' txtFreight.Text.ToString
                        .Parameters.Add("@Cust_Name", SqlDbType.VarChar, 200).Value = lblCustCodeDes.Text.ToString
                        .Parameters.Add("@Round_Off", SqlDbType.VarChar, 40).Value = Format(Double.Parse(lblRoundOff.Text), "#.00") 'Val(.Text.ToString)
                        .Parameters.Add("@User_Id", SqlDbType.VarChar, 100).Value = mP_User
                        .Parameters.Add("@Unit_Code", SqlDbType.VarChar, 10).Value = gstrUNITID

                        .Parameters.Add("@CGST_Tax_Value", SqlDbType.VarChar, 100).Value = Val(CGST_VAL.ToString)
                        .Parameters.Add("@SGST_Tax_Value", SqlDbType.VarChar, 100).Value = Val(SGST_VAL.ToString)
                        .Parameters.Add("@IGST_Tax_Value", SqlDbType.VarChar, 100).Value = Val(IGST_VAL.ToString)
                        .Parameters.Add("@UTGST_Tax_Value", SqlDbType.VarChar, 100).Value = Val(UTGST_VAL.ToString)
                        .Parameters.Add("@CSESS_Tax_Value", SqlDbType.VarChar, 100).Value = Val(CSESS_VAL.ToString)
                        .Parameters.Add("@Advance_Amount", SqlDbType.VarChar, 100).Value = Val(ADVANCE_AMT.ToString.Trim)
                        .Parameters.Add("@Para", SqlDbType.VarChar, 200).Value = "SAVESALES_DTL"

                        Dim retMessage As SqlParameter = New SqlParameter("@PI_No", SqlDbType.VarChar, 10)
                        retMessage.Direction = ParameterDirection.Output
                        .Parameters.Add(retMessage)


                        Dim retErrorMessage As SqlParameter = New SqlParameter("@Error_Msg", SqlDbType.VarChar, 8000)
                        retErrorMessage.Direction = ParameterDirection.Output
                        .Parameters.Add(retErrorMessage)

                        cmd.ExecuteNonQuery()
                        DocumentNo = .Parameters("@PI_No").Value.ToString()
                        Error_Msg = .Parameters("@Error_Msg").Value.ToString()

                        If Not String.IsNullOrEmpty(Error_Msg) Then
                            MessageBox.Show(Error_Msg.ToString, "Empro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            QryExecuted = False
                            cmd.Transaction.Rollback()
                            Exit Function
                        Else
                            If Not String.IsNullOrEmpty(DocumentNo.ToString) Then
                                QryExecuted = True
                            Else
                                cmd.Transaction.Rollback()
                                SaveData = False
                                QryExecuted = False
                                MessageBox.Show("Document Number not generated.", "Empro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                Exit Function
                            End If
                        End If

                        If Not String.IsNullOrEmpty(DocumentNo) Then
                            Dim Item_code As Object
                            Dim Itemtotal As Object
                            Dim Rate As Object
                            Dim Cust_Drgno As Object
                            Dim Cust_DrgNo_Desc As Object
                            Dim Basic_Value As Object
                            Dim Assessible_Value As Object
                            Dim Discount_Value As Object
                            Dim Discount_Per As Object
                            Dim Remarks As Object
                            Dim HSN_SAC_NO As Object
                            Dim HSN_SAC_TYPE As Object
                            Dim CGST_Tax_Type As Object
                            Dim CGST_Tax_Per As Object
                            Dim CGST_Tax_Value As Object
                            Dim SGST_Tax_Type As Object
                            Dim SGST_Tax_Per As Object
                            Dim SGST_Tax_Value As Object
                            Dim UTGST_Tax_Type As Object
                            Dim UTGST_Tax_Per As Object
                            Dim UTGST_Tax_Value As Object
                            Dim IGST_Tax_Type As Object
                            Dim IGST_Tax_Per As Object
                            Dim IGST_Tax_Value As Object
                            Dim CSESS_Tax_Type As Object
                            Dim CSESS_Tax_Per As Object
                            Dim CSESS_Tax_Value As Object
                            Dim Advance_Amount As Object
                            Dim Quantity As Object

                            For intcounter As Integer = 1 To sspr.MaxRows
                                Item_code = Nothing
                                Itemtotal = Nothing
                                Rate = Nothing
                                Cust_Drgno = Nothing
                                Cust_DrgNo_Desc = Nothing
                                Basic_Value = Nothing
                                Assessible_Value = Nothing
                                Discount_Value = Nothing
                                Discount_Per = Nothing
                                Remarks = Nothing
                                HSN_SAC_TYPE = Nothing
                                Quantity = Nothing
                                With sspr
                                    .Row = intcounter

                                    .Col = EnumInv.Quantity
                                    Quantity = Val(.Text)

                                    .Col = EnumInv.ItemCode
                                    Item_code = .Text

                                    .Col = EnumInv.ItemTotal
                                    Itemtotal = Val(.Text)

                                    .Col = EnumInv.Rate
                                    Rate = Val(.Text)

                                    .Col = EnumInv.Cust_Drgno
                                    Cust_Drgno = .Text

                                    .Col = EnumInv.Cust_DrgNo_Desc
                                    Cust_DrgNo_Desc = .Text

                                    .Col = EnumInv.Basic_value
                                    Basic_Value = Val(.Text)

                                    .Col = EnumInv.Assable_Value
                                    Assessible_Value = Val(.Text)

                                    .Col = EnumInv.DiscountVal
                                    Discount_Value = Val(.Text)

                                    .Col = EnumInv.DiscountPer
                                    Discount_Per = Val(.Text)

                                    .Col = EnumInv.Remarks
                                    Remarks = .Text

                                    .Col = EnumInv.HSN_SAC_No
                                    HSN_SAC_NO = .Text

                                    .Col = EnumInv.HSN_SAC_Type
                                    HSN_SAC_TYPE = .Text

                                    .Col = EnumInv.CGST_Tax_type
                                    CGST_Tax_Type = .Text

                                    .Col = EnumInv.CGST_Tax_Per
                                    CGST_Tax_Per = Val(.Text)

                                    .Col = EnumInv.CGST_Tax_Value
                                    CGST_Tax_Value = Val(.Text)

                                    .Col = EnumInv.IGST_Tax_type
                                    IGST_Tax_Type = .Text

                                    .Col = EnumInv.IGST_Tax_Per
                                    IGST_Tax_Per = Val(.Text)

                                    .Col = EnumInv.IGST_Tax_Value
                                    IGST_Tax_Value = Val(.Text)

                                    .Col = EnumInv.SGST_Tax_type
                                    SGST_Tax_Type = Val(.Text)

                                    .Col = EnumInv.SGST_Tax_Per
                                    SGST_Tax_Per = Val(.Text)

                                    .Col = EnumInv.SGST_Tax_Value
                                    SGST_Tax_Value = Val(.Text)

                                    .Col = EnumInv.UTGST_Tax_type
                                    UTGST_Tax_Type = .Text

                                    .Col = EnumInv.UTGST_Tax_Per
                                    UTGST_Tax_Per = Val(.Text)

                                    .Col = EnumInv.UTGST_Tax_Value
                                    UTGST_Tax_Value = Val(.Text)

                                    .Col = EnumInv.CSESS_Tax_type
                                    CSESS_Tax_Type = .Text

                                    .Col = EnumInv.CSESS_Tax_Per
                                    CSESS_Tax_Per = Val(.Text)

                                    .Col = EnumInv.CSESS_Tax_Value
                                    CSESS_Tax_Value = Val(.Text)

                                    .Col = EnumInv.Advance_Amt
                                    Advance_Amount = Val(.Text)




                                    '----- DTL  LEVEL i.e. Item Level Entry
                                    With cmd

                                        cmd.CommandText = String.Empty
                                        cmd.Parameters.Clear()
                                        .CommandTimeout = 0
                                        .CommandType = CommandType.StoredProcedure
                                        .CommandText = "Proc_GST_Performa_Invoice"

                                        .Parameters.Add("@Doc_No", SqlDbType.VarChar, 10).Value = DocumentNo.ToString.Trim
                                        .Parameters.Add("@Item_Code", SqlDbType.VarChar, 30).Value = Item_code.ToString


                                        .Parameters.Add("@Quantity", SqlDbType.VarChar, 100).Value = Val(Quantity.ToString)
                                        .Parameters.Add("@Rate", SqlDbType.VarChar, 100).Value = Rate.ToString.Trim
                                        .Parameters.Add("@Cust_Drgno", SqlDbType.VarChar, 100).Value = Cust_Drgno.ToString.ToString
                                        .Parameters.Add("@Cust_DrgNo_Desc", SqlDbType.VarChar, 200).Value = Cust_DrgNo_Desc.ToString.ToString
                                        .Parameters.Add("@Basic_Value", SqlDbType.VarChar, 100).Value = Basic_Value.ToString.Trim
                                        .Parameters.Add("@Assessible_Value", SqlDbType.VarChar, 100).Value = Assessible_Value.ToString
                                        .Parameters.Add("@Discount_Value", SqlDbType.VarChar, 100).Value = Discount_Value.ToString
                                        .Parameters.Add("@Discount_Per", SqlDbType.VarChar, 100).Value = Discount_Per.ToString


                                        .Parameters.Add("@SO_No", SqlDbType.VarChar, 100).Value = SONoSPlit(0).ToString.Trim
                                        .Parameters.Add("@Amendment_No", SqlDbType.VarChar, 200).Value = SONoSPlit(1).ToString.Trim
                                        .Parameters.Add("@Remarks", SqlDbType.VarChar, 200).Value = Remarks.ToString.Trim.Replace("'", "''")


                                        .Parameters.Add("@HSN_SAC_NO", SqlDbType.VarChar, 100).Value = HSN_SAC_NO.ToString
                                        .Parameters.Add("@HSN_SAC_Type", SqlDbType.VarChar, 100).Value = HSN_SAC_TYPE.ToString

                                        .Parameters.Add("@CGST_Tax_Type", SqlDbType.VarChar, 200).Value = CGST_Tax_Type.ToString
                                        .Parameters.Add("@CGST_Tax_Per", SqlDbType.VarChar, 100).Value = CGST_Tax_Per.ToString
                                        .Parameters.Add("@CGST_Tax_Value", SqlDbType.VarChar, 100).Value = CGST_Tax_Value.ToString


                                        .Parameters.Add("@SGST_Tax_Type", SqlDbType.VarChar, 200).Value = SGST_Tax_Type.ToString
                                        .Parameters.Add("@SGST_Tax_Per", SqlDbType.VarChar, 100).Value = SGST_Tax_Per.ToString
                                        .Parameters.Add("@SGST_Tax_Value", SqlDbType.VarChar, 100).Value = SGST_Tax_Value.ToString


                                        .Parameters.Add("@UTGST_Tax_Type", SqlDbType.VarChar, 200).Value = UTGST_Tax_Type.ToString
                                        .Parameters.Add("@UTGST_Tax_Per", SqlDbType.VarChar, 100).Value = UTGST_Tax_Per.ToString
                                        .Parameters.Add("@UTGST_Tax_Value", SqlDbType.VarChar, 100).Value = UTGST_Tax_Value.ToString



                                        .Parameters.Add("@IGST_Tax_Type", SqlDbType.VarChar, 200).Value = IGST_Tax_Type.ToString
                                        .Parameters.Add("@IGST_Tax_Per", SqlDbType.VarChar, 100).Value = IGST_Tax_Per.ToString
                                        .Parameters.Add("@IGST_Tax_Value", SqlDbType.VarChar, 100).Value = IGST_Tax_Value.ToString


                                        .Parameters.Add("@CSESS_Tax_Type", SqlDbType.VarChar, 200).Value = CSESS_Tax_Type.ToString
                                        .Parameters.Add("@CSESS_Tax_Per", SqlDbType.VarChar, 100).Value = CSESS_Tax_Per.ToString
                                        .Parameters.Add("@CSESS_Tax_Value", SqlDbType.VarChar, 100).Value = CSESS_Tax_Value.ToString

                                        .Parameters.Add("@Total_Amount", SqlDbType.VarChar, 100).Value = Itemtotal.ToString

                                        .Parameters.Add("@Advance_Amount", SqlDbType.VarChar, 100).Value = Advance_Amount.ToString

                                        .Parameters.Add("@User_Id", SqlDbType.VarChar, 100).Value = mP_User
                                        .Parameters.Add("@Unit_Code", SqlDbType.VarChar, 10).Value = gstrUNITID

                                        .Parameters.Add("@Para", SqlDbType.VarChar, 200).Value = "SAVECHALLAN_DTL"

                                        Dim retError_Message As SqlParameter = New SqlParameter("@Error_Msg", SqlDbType.VarChar, 8000)
                                        retError_Message.Direction = ParameterDirection.Output
                                        .Parameters.Add(retError_Message)

                                        cmd.ExecuteNonQuery()
                                        Error_Msg = .Parameters("@Error_Msg").Value.ToString()
                                        If Not String.IsNullOrEmpty(Error_Msg) Then
                                            cmd.Transaction.Rollback()
                                            MessageBox.Show("4427-Transaction Fail.", "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                            QryExecuted = False
                                            SaveData = False
                                            Exit Function
                                        Else
                                            QryExecuted = True
                                        End If
                                        '----- For Sql Cmd..
                                    End With
                                    '----- For Loop For FarSpread
                                End With
                            Next
                            If QryExecuted Then
                                cmd.Transaction.Commit()
                                cmd.Dispose()
                                SaveData = True
                                txtChallanNo.Text = DocumentNo.ToString.Trim

                            End If
                        End If

                    End With

                Catch ex As Exception
                    cmd.Transaction.Dispose()
                    RaiseException(ex)
                    Exit Function
                End Try
                '------------------END------


            End If
            '---------------------------STOCK UPDATION----------------------------------

        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            If IsNothing(cmd) = False Then
                cmd.Dispose()
            End If

        End Try
    End Function


    Private Sub CmdChallanNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdChallanNo.Click
        Try
            Dim strQselect As String = String.Empty
            Dim PI_detail As String = String.Empty
            strQselect = " SELECT DISTINCT  A.DOC_NO AS PI_NO,A.INVOICE_DATE AS PI_DATE  " & _
                          " , B.Cust_ref +','+ B.Amendment_No AS SO_No,A.ACCOUNT_CODE AS CUST_CODE ,A.CUST_NAME  , CASE WHEN A.CANCEL_FLAG=0 THEN 'OPEN' WHEN A.CANCEL_FLAG=1 THEN 'CLOSE' END AS STATUS,CASE WHEN (SELECT Top 1 1 FROM FIN_BV_PI_MAPPING WHERE UNIT_CODE=A.UNIT_CODE and PI_NO=A.Doc_No  )=1 THEN 'No' Else 'Yes' END AS Cancellation_Allowed " & _
                          " FROM PI_SALESCHALLAN_DTL(NOLOCK) AS A" & _
                          " INNER JOIN PI_SALES_DTL (NOLOCK) AS B" & _
                          " ON A.UNIT_CODE=B.UNIT_CODE AND A.CUST_REF=B.CUST_REF AND A.AMENDMENT_NO=B.AMENDMENT_NO AND A.DOC_NO=B.DOC_NO " & _
                          " WHERE A.UNIT_CODE='" & gstrUNITID & "' ORDER BY A.INVOICE_DATE DESC"


            
            PI_detail = GetDocumentNo(strQselect, "PI_No")
            If PI_detail = Nothing Then
                MessageBox.Show("Operation Cancelled", "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                Exit Sub
            End If
            Dim Split() As String = PI_detail.Split("~")
            If Split(0).Contains("No record found.") Then
                MessageBox.Show("No Record Found", "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                Exit Sub
            End If
            If Not String.IsNullOrEmpty(Split(0)) Then
                txtChallanNo.Text = Split(0).ToString
                If IsDate(Split(1).ToString) Then
                    dtpDateDesc.Text = Split(1).ToString
                Else
                    dtpDateDesc.Text = Now.Date.ToString
                End If
                If Not String.IsNullOrEmpty(Split(3)) Then
                    txtCustCode.Text = Split(3).ToString
                End If
                If Not String.IsNullOrEmpty(Split(4)) Then
                    lblCustCodeDes.Text = Split(4).ToString
                End If
                Dim SONoSPlit() As String = Split(2).ToString.Split(",")
                If Not String.IsNullOrEmpty(SONoSPlit(0)) Then
                    txtRefNo.Text = SONoSPlit(0).ToString
                End If
                Call Fill_Grid(SONoSPlit(0).ToString, SONoSPlit(1).ToString, "SHOW")  '--- FILL GRID
                If IsDate(Split(1).ToString) Then
                    dtpDateSO.Text = Split(1).ToString
                Else
                    dtpDateSO.Text = Now.Date.ToString
                End If
                Me.Group4.Enabled = True
                Me.Group2.Enabled = True
                Me.Cmditems.Enabled = True
                If Not String.IsNullOrEmpty(Split(5)) Then
                    If Split(5).ToString = "CLOSE" Then
                        lblCancelledInvoice.Visible = True
                    ElseIf Split(5).ToString = "OPEN" Then
                        lblCancelledInvoice.Visible = False
                    End If

                End If
                If Not String.IsNullOrEmpty(Split(6)) Then
                    If Split(6).ToString = "Yes" And Split(5).ToString = "CLOSE" Then
                        Me.Cmditems.Visible = False
                    ElseIf Split(6).ToString = "Yes" And Split(5).ToString = "OPEN" Then
                        Me.Cmditems.Visible = True
                    ElseIf Split(6).ToString = "No" Then
                        Me.Cmditems.Visible = False
                    End If

                End If

            Else
                txtChallanNo.Text = String.Empty
                Call ConfirmWindow(10225, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            End If





        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub



    Private Sub Cmditems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmditems.Click
        Dim cmd As SqlCommand = Nothing
        Try
            If MsgBox("Do you want to cancel this Invoice? ", vbYesNo, "eMPro") = vbYes Then
            Else
                Exit Sub
            End If
            If txtChallanNo.Text.ToString.Trim.Length > 0 Then


                Dim Error_Msg As String = String.Empty
                SqlConnectionclass.CloseGlobalConnection()
                SqlConnectionclass.OpenGlobalConnection()

                cmd = New System.Data.SqlClient.SqlCommand()
                cmd.Connection = SqlConnectionclass.GetConnection
                cmd.Transaction = cmd.Connection.BeginTransaction(System.Data.IsolationLevel.Serializable)
                With cmd

                    .CommandText = String.Empty
                    .Parameters.Clear()
                    .CommandTimeout = 0
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "Proc_GST_Performa_Invoice"

                    .Parameters.Add("@Doc_No", SqlDbType.VarChar, 100).Value = txtChallanNo.Text.ToString.Trim
                    .Parameters.Add("@User_Id", SqlDbType.VarChar, 100).Value = mP_User
                    .Parameters.Add("@Unit_Code", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@Remarks_Credit", SqlDbType.VarChar, 200).Value = txtRemarks.Text.ToString.Trim
                    .Parameters.Add("@Para", SqlDbType.VarChar, 200).Value = "CancelInvoice"

                    Dim retErrorMessage As SqlParameter = New SqlParameter("@Error_Msg", SqlDbType.VarChar, 8000)
                    retErrorMessage.Direction = ParameterDirection.Output
                    .Parameters.Add(retErrorMessage)

                    cmd.ExecuteNonQuery()

                    Error_Msg = .Parameters("@Error_Msg").Value.ToString()

                    If Not String.IsNullOrEmpty(Error_Msg) Then
                        MessageBox.Show(Error_Msg.ToString, "Empro", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        cmd.Transaction.Rollback()
                        Cmditems.Visible = True
                        Exit Sub
                    Else
                        cmd.Transaction.Commit()
                        MessageBox.Show("Invoice Cancelled Successfully.".ToString, "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                        lblCancelledInvoice.Enabled = True
                        lblCancelledInvoice.Visible = True
                        Cmditems.Visible = False
                    End If
                End With

            End If
        Catch ex As Exception
            RaiseException(ex)

        End Try
    End Sub
End Class