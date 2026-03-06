Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.Vbe.Interop
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Data.SqlClient
Friend Class frmMKTTRN0018
    Inherits System.Windows.Forms.Form
    '============================================================================================================
    ' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
    ' File Name         :   FRMMKTTRN0018.frm
    ' Function          :   Used to ADD/EDIT/DELETE Supplementary Invoice
    ' Created By        :   Nisha
    ' Created On        :   03 Nov 2003
    ' Revision History  :   1. 17 nov 2003
    '                   :   2. 13 Oct 2004
    '                   :   3. 18 Jan 2005
    ' Revision History  :   1. changed for adding Print Preview of Custannex_dtl
    '                   :   2. Changes done by Sourabh on 13 oct 2004 for add ECSS Value.
    '                   :   3. Changed the display of Previous Invoices and included a new DLL prj_InvoiceCalc
    '                   :      The Dll will do the calculation of the formulas that were feeded in the grid
    '                   :      The save data string for Sales Dtl Table, Supplementary_Dtl table and Credit
    '                   :      Advice Dtl Table
    '===========================================================================================================
    ' Revision History  :
    '                   :   1. Data stored into the supplementary table according to sales parameter table (roundoff Fields)
    ' Revisised By      :   Parveen Kumar on 07 Mar 2006
    'Issue Id           :   17290 : Supplementary invoice posting problem (Supp. Invoice - Db. Cr. Mismatch).
    '--------------------------------------------------------------------------------
    ' Revised by          - Davinder Singh
    ' Revision Date       - 26 Apr 2007
    ' Revision History    - 1) To include the Secondary Ecess(SECESS) in Supplementary Invoice
    '                       2) DURING PICKING ECESS,SALESTAX,SURCHARGE_SALES_TAX
    '                          IT SHOULD BE CALCULATED
    '                          INSTEAD OF PICKING IT DIRECTLY FROM HDR TABLE AS
    '                          HDR TABLE CONTAINS SUM FOR ALL ITEMS IN INVOICE
    ' Issue ID            - 19786
    '--------------------------------------------------------------------------------
    ' Revised by          - Davinder Singh
    ' Revision Date       - 01 May 2007
    ' Revision History    - 1) While making Supp inv. agst invoices made with MRP Excise,Accessible,
    '                          Ecess,Shecess etc. calculated -ve bcoz all calculations were made with
    '                          Rate instead of MRP. Now the concept of MRP introduced.
    '                       2) Calculation of invoice is also not correct.
    '                          (eg. We are taking Packing % as Packing amount. ,
    '                           Cust. Supp. Material is added in Total invoice value)
    '                       3) During locking of invoice we are not Posting PKG
    ' Issue ID            - 19958
    '--------------------------------------------------------------------------------
    '                               THINGS NEED TO BE DONE (Davinder)
    ' 1) Multilevel Supp Inv concept is not properly implemented
    '    (Inv - Supp - Supp - .....)
    ' 2) While considering supplemetary invoices in function SUPPLEMENTARYDATA
    '    Bill flag was not considered (so user can make supplementary invoice
    '    upon supplementary and then edit the any invoice below the tree or can delete it)
    ' 3) All taxes are not handled in supp inv.
    ' 4) EOU unit concept is not implemented properly its totally wrong (CVD etc.)
    '--------------------------------------------------------------------------------
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : 21551
    'Revision Date   : 22-Nov-2007
    'History         : Add New Tax VAT with Sale Tax help
    '--------------------------------------------------------------------------------
    'Revised By      : Vinod Singh Kemwal
    'Revision Date   : 09/06/2011
    'History         : Multi Unit changes
    '--------------------------------------------------------------------------------
    'Revised By      : Vinod Singh
    'Issue ID        : 10137894 
    'Revision Date   : 16/09/2011
    'Desc           : Object Referece error while selecting invoice nos
    '----------------------------------------------------------------
    ' Revised By     :   Pankaj Kumar
    ' Revision Date  :   13 Oct 2011
    ' Description    :   Modified for MultiUnit Change Management

    'Revised By         : Ashish Sharma
    'Issue ID           : 101188073
    'Revision Date      : 18 JUL 2017
    'Desc               : GST Changes
    '***********************************************************************************
    Private Enum enumPreInvoiceDetails
        Select_Invoice = 1
        Invoice_No = 2
        Invoice_Date = 3
        LastSupplementary = 4
        SupplementaryDate = 5
        Quantity = 6
        Rate = 7
        New_Rate = 8
        Rate_diff = 9
        TotalPacking = 10
        NewPacking = 11
        NewTotalPacking = 12
        Basic = 13
        NewBasic = 14
        BasicDiff = 15
        TotalCustSuppMaterial = 16
        NewCustSuppMaterial = 17
        NewTotalCustSuppMaterial = 18
        CustSuppMaterial_diff = 19
        ToolCost = 20
        newToolCost = 21
        NewTotalToolCost = 22
        ToolCost_diff = 23
        AccessableValue = 24
        NewAccessableValue = 25
        AccessableValue_Diff = 26
        TotalExciseValue = 27
        NewExciseValue = 28
        NewCVDValue = 29
        NewSADValue = 30
        NewTotalExciseValue = 31
        TotalExciseValueDiff = 32
        TotalEcessValue = 33
        NewEcessValue = 34
        TotalEcssDiff = 35
        TotalsEcessValue = 36
        NewsEcessValue = 37
        TotalsEcssDiff = 38
        SalesTaxValue = 39
        NewSalesTaxValue = 40
        SalesTaxValueDiff = 41
        SSTVlaue = 42
        NewSSTValue = 43
        SSTVlaueDiff = 44
        '101188073
        CGST_AMT = 45
        NEW_CGST_AMT = 46
        DIFF_CGST_AMT = 47
        SGST_AMT = 48
        NEW_SGST_AMT = 49
        DIFF_SGST_AMT = 50
        UTGST_AMT = 51
        NEW_UTGST_AMT = 52
        DIFF_UTGST_AMT = 53
        IGST_AMT = 54
        NEW_IGST_AMT = 55
        DIFF_IGST_AMT = 56
        CCESS_AMT = 57
        NEW_CCESS_AMT = 58
        DIFF_CCESS_AMT = 59
        '101188073
        TotalCurrInvValue = 60
        TotalInvoiceValue = 61
        flag = 62
    End Enum
    Private Enum enumInvoiceSummery
        Rate = 1
        BasicValue = 2
        CustSuppMat = 3
        ToolCost = 4
        AccessableValue = 5
        ExciseValue = 6
        EcssValue = 7
        sEcssValue = 8
        PackingValue = 9
        SalesTaxType = 10
        SSTType = 11
        '101188073
        CGST = 12
        SGST = 13
        UTGST = 14
        IGST = 15
        CCESS = 16
        '101188073
        SummeryInvoiceValue = 17
    End Enum
    Dim mintIndex As Short
    Dim Financial_Start_Date As Date
    Dim Financial_End_Date As Date
    Dim blnEOU_FLAG As Boolean
    Dim strDefaultLocation As String
    Private strSelInvoices As String 'to store the selected invoice numbers
    Private strNotSelInvoices As String ' to store the deselected invoice numbers
    Private bCheck As Boolean 'boolean to specify the class function if invoice inclusion check is required
    Private blnInclude As Boolean ' retval of get invoice summary function from the dll
    Dim strInvoiceNo As String
    Dim cColor As System.Drawing.Color
    Dim mstrCustRefNo As String
    Dim mstrSOType As New VB6.FixedLengthString(1)
    Dim bool_check As Boolean = False
    Dim bool_valid As Boolean = False
    Dim frmRpt As eMProCrystalReportViewer
    Dim CRDoc As ReportDocument
    Private objInvoiceCls As New prj_InvoiceCalc.clsInvoiceCalculation(gstrUnitId)
    Dim _cgstApplicable As Boolean = True
    Dim _sgstApplicable As Boolean = True
    Dim _utgstApplicable As Boolean = True
    Dim _igstApplicable As Boolean = True
    Dim _ccessApplicable As Boolean = True

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
                If ToCheckForNegativeValuesNoMessages(intLoopCounter) = True Then
                    With lstInv
                        bool_check = True
                        .Items.Item(intLoopCounter).Checked = True
                        bool_check = False
                    End With
                Else
                    With lstInv
                        bool_check = True
                        .Items.Item(intLoopCounter).Checked = False
                        bool_check = False
                    End With
                End If
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
    Private Sub CmdChallanNo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdChallanNo.Click
        On Error GoTo ErrHandler
        Dim strHelpString As String
        Dim strChallanNo() As String
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Trim(txtLocationCode.Text) = "" Then
                    Call ConfirmWindow(10239, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO, 100)
                    If txtLocationCode.Enabled Then txtLocationCode.Focus()
                    Exit Sub
                End If
                strHelpString = "Select DISTINCT Doc_No,cast(Location_Code as varchar(10)) as location_code from SupplementaryInv_hdr where Unit_code='" & gstrUNITID & "' and Location_code ='" & txtLocationCode.Text & "'"
                strHelpString = Trim(strHelpString) & " and Cancel_flag = 0"
        End Select
        strChallanNo = ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelpString, "Supplementary Invoice No")
        If UBound(strChallanNo) < 0 Then Exit Sub
        If strChallanNo(0) = "0" Then
            MsgBox("No Supplementary Invoice Available To Display", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100)) : txtChallanNo.Text = "" : txtChallanNo.Focus() : Exit Sub
        Else
            If strChallanNo(0) <> "" Then
                txtChallanNo.Text = strChallanNo(0)
            End If
        End If
        txtChallanNo.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        fraInvoice.Visible = False
    End Sub
    Private Sub cmdCustAmend_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCustAmend.Click
        On Error GoTo ErrHandler
        Dim strrefno() As String
        Dim strRefSql As String
        If Len(Trim(txtRefNo.Text)) = 0 Then
            MsgBox("First select Customer Referance", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            txtRefNo.Focus()
            Exit Sub
        End If
        strRefSql = "Select b.Cust_Ref,b.Amendment_No from Cust_Ord_hdr a,Cust_Ord_Dtl b"
        strRefSql = strRefSql & " where a.unit_code=b.unit_code and a.Unit_code='" & gstrUNITID & "' and b.Account_Code='" & Trim(txtCustCode.Text) & "' and b.Active_flag ='A' and "
        strRefSql = strRefSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and PO_type <> 'E'and "
        strRefSql = strRefSql & " a.Amendment_No = b.amendment_No  AND a.Authorized_Flag = 1 and b.Cust_drgNo = '"
        strRefSql = strRefSql & Trim(txtCustPartCode.Text) & "' and ITem_code = '" & Trim(lblItemCode.Text) & "'"
        strRefSql = strRefSql & " and a.Valid_date >'" & Format(GetServerDate, "dd MMM yyyy") & "' and effect_Date <= '"
        strRefSql = strRefSql & Format(GetServerDate, "dd MMM yyyy") & "' and b.Cust_ref = '" & Trim(txtRefNo.Text)
        strRefSql = strRefSql & "' order by b.Cust_Ref,b.Amendment_No"
        strrefno = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strRefSql, "Amendment Details ")
        If UBound(strrefno) < 0 Then Exit Sub
        If strrefno(0) = "0" Then
            MsgBox("No Refrence available to Display", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100)) : txtRefNo.Text = "" : txtRefNo.Focus() : Exit Sub
        Else
            If strrefno(0) <> "" Then
                txtRefNo.Text = strrefno(0)
                txtAmendment.Text = strrefno(1)
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub CmdCustCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdCustCodeHelp.Click
        Dim strCustMst As String
        Dim strCustHelp As String
        Dim rsCustMst As ClsResultSetDB
        Dim strCustomer() As String
        On Error GoTo ErrHandler
        If Len(Trim(txtLocationCode.Text)) = 0 Then
            Call ConfirmWindow(10116, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            If txtLocationCode.Enabled Then txtLocationCode.Focus()
            Exit Sub
        End If
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                strCustHelp = "Select DISTINCT Account_code,Cust_Name from VW_SUPP_INVOICE_CUSTHELP Where Unit_code='" & gstrUNITID & "' and Location_code = '" & Trim(txtLocationCode.Text) & "'"
                strCustHelp = Trim(strCustHelp) & " and invoice_date >= '" & Format(dtpDateFrom.Value, "dd MMM yyyy") & "' and "
                strCustHelp = Trim(strCustHelp) & " Invoice_date <= '" & Format(dtpDateTo.Value, "dd MMM yyyy") & "' "
        End Select
        strCustomer = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strCustHelp, "Customer List")
        If UBound(strCustomer) < 0 Then Exit Sub
        If strCustomer(0) = "0" Then
            MsgBox("No Customer Available to Display", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100)) : txtCustCode.Text = "" : txtCustCode.Focus() : Exit Sub
        Else
            If strCustomer(0) <> "" Then
                txtCustCode.Text = strCustomer(0)
                lblCustCodeDes.Text = strCustomer(1)
            End If
        End If
        If Len(Trim(txtCustCode.Text)) > 0 Then
            rsCustMst = New ClsResultSetDB
            strCustMst = "SELECT Bill_Address1 + ', '  +  Bill_Address2 + ', ' + Bill_City + ' - ' + Bill_Pin as  invoiceAddress from Customer_Mst where Unit_code='" & gstrUNITID & "' and Customer_code ='" & txtCustCode.Text & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
            rsCustMst.GetResult(strCustMst)
            If rsCustMst.GetNoRows > 0 Then
                lblAddressDes.Text = rsCustMst.GetValue("InvoiceAddress")
            End If
            rsCustMst.ResultSetClose()
            rsCustMst = Nothing
        End If
        Call txtCustCode_Leave(txtCustCode, New System.EventArgs())
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdcustPartCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdcustPartCode.Click
        Dim strCustItemMst As String
        Dim strCustItemHelp As String
        Dim rsCustItemMst As ClsResultSetDB
        Dim strCustItem() As String
        On Error GoTo ErrHandler
        If Len(Trim(txtLocationCode.Text)) = 0 Then
            Call ConfirmWindow(10116, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            If txtLocationCode.Enabled Then txtLocationCode.Focus()
            Exit Sub
        End If
        If Len(Trim(txtCustCode.Text)) = 0 Then
            MsgBox("First Select Customer Code.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            If txtCustCode.Enabled Then txtCustCode.Focus()
            Exit Sub
        End If
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                strCustItem = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select DISTINCT Cust_Item_code,Item_code,cast(Cust_Item_desc as varchar(50)) as Cust_Item_desc  from Saleschallan_dtl a,sales_dtl b where a.Unit_code = b.Unit_code and a.Unit_code='" & gstrUNITID & "' and  Account_code = '" & txtCustCode.Text & "' and invoice_date > = '" & Format(dtpDateFrom.Value, "dd MMM yyyy") & "' and invoice_date < = '" & Format(dtpDateTo.Value, "dd MMM yyyy") & "' and a.Doc_no = b.Doc_no and bill_flag = 1 and cancel_flag =0", "Customer Part Code")
        End Select
        If UBound(strCustItem) < 0 Then Exit Sub
        If strCustItem(0) = "0" Then
            MsgBox("No Customer Part Code for selected Customer & Date Range is Availabe To Dispaly", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100)) : txtCustCode.Text = "" : txtCustCode.Focus() : Exit Sub
        Else
            If strCustItem(0) <> "" Then
                txtCustPartCode.Text = strCustItem(0)
                lblItemCode.Text = strCustItem(1)
                lblCustItemDesc.Text = strCustItem(2)
            End If
        End If
        If Len(Trim(txtCustCode.Text)) > 0 Then
            rsCustItemMst = New ClsResultSetDB
            strCustItemMst = "SELECT Bill_Address1 + ', '  +  Bill_Address2 + ', ' + Bill_City + ' - ' + Bill_Pin as  invoiceAddress from Customer_Mst where Unit_code='" & gstrUNITID & "' and Customer_code ='" & txtCustCode.Text & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
            rsCustItemMst.GetResult(strCustItemMst)
            If rsCustItemMst.GetNoRows > 0 Then
                lblAddressDes.Text = rsCustItemMst.GetValue("InvoiceAddress")
            End If
            rsCustItemMst.ResultSetClose()
            rsCustItemMst = Nothing
        End If
        Call txtCustPartCode_Leave(txtCustPartCode, New System.EventArgs())
        txtRefNo.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdCVDCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCVDCodeHelp.Click
        Dim strHelp As String
        Dim strSTaxHelp() As String
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strHelp = "Select TxRT_Rate_No,TxRt_Percentage from Gen_taxRate where Unit_code='" & gstrUNITID & "' and Tx_TaxeID ='CVD' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                strSTaxHelp = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, "CVD Tax Help")
                If UBound(strSTaxHelp) < 0 Then Exit Sub
                If strSTaxHelp(0) = "0" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtCVDCode.Text = "" : txtCVDCode.Focus() : Exit Sub
                Else
                    If strSTaxHelp(0) <> "" Then
                        txtCVDCode.Text = strSTaxHelp(0)
                        lblCVD_Per.Text = strSTaxHelp(1)
                        txtCVDCode.Focus()
                    End If
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdEcssCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEcssCodeHelp.Click
        On Error GoTo Errorhandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        Dim strHelp As String
        Dim strSSTaxHelp() As String
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strHelp = "Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where Unit_code='" & gstrUNITID & "' and tx_TaxeID ='ECS' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                strSSTaxHelp = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, "E CESS Tax Help")
                If UBound(strSSTaxHelp) < 0 Then Exit Sub
                If strSSTaxHelp(0) = "0" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : Me.txtEcssCode.Text = "" : Me.txtEcssCode.Focus() : Exit Sub
                Else
                    If strSSTaxHelp(0) <> "" Then
                        txtEcssCode.Text = strSSTaxHelp(0)
                        lblEcssCode.Text = strSSTaxHelp(1)
                        Call SetFormulaofColumns(0, 4)
                        Call ToShowDatainSummeryGrid()
                        txtEcssCode.Focus()
                    End If
                End If
        End Select
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub CmdExciseTaxType_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdExciseTaxType.Click
        Dim strHelp As String
        Dim strSTaxHelp() As String
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strHelp = "Select Distinct TxRT_Rate_No,TxRt_Percentage from Gen_taxRate where Unit_code='" & gstrUNITID & "' and Tx_TaxeID ='EXC' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                strSTaxHelp = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, "Exc.Tax Help")
                If UBound(strSTaxHelp) < 0 Then Exit Sub
                If strSTaxHelp(0) = "0" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtExciseTaxType.Text = "" : txtExciseTaxType.Focus() : Exit Sub
                Else
                    If strSTaxHelp(0) <> "" Then
                        txtExciseTaxType.Text = strSTaxHelp(0)
                        lblExctax_Per.Text = strSTaxHelp(1)
                        Call SetFormulaofColumns(0, 4)
                        Call ToShowDatainSummeryGrid()
                        txtExciseTaxType.Focus()
                    End If
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOk.Click
        Dim Intcounter As Short
        strSelInvoices = ""
        For Intcounter = 0 To lstInv.Items.Count - 1
            If lstInv.Items.Item(Intcounter).Checked = True Then
                strSelInvoices = strSelInvoices & "|" & lstInv.Items.Item(Intcounter).Text
            Else
                strNotSelInvoices = strNotSelInvoices & "|" & lstInv.Items.Item(Intcounter).Text
            End If
        Next Intcounter
        If VB.Left(strSelInvoices, 1) = "|" Then
            strSelInvoices = VB.Right(strSelInvoices, Len(strSelInvoices) - 1)
        End If
        If VB.Left(strNotSelInvoices, 1) = "|" Then
            strNotSelInvoices = VB.Right(strNotSelInvoices, Len(strNotSelInvoices) - 1)
        End If
        If strSelInvoices = "" Then
            Call ConfirmWindow(60676, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
        Else
            fraInvoice.Visible = False
            Call SetFormulaofColumns(0, 4)
            Call ToShowDatainSummeryGrid()
        End If
    End Sub
    Private Sub cmdPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrint.Click
        On Error GoTo ErrHandler
        Dim strSQL As String
        Dim strSelectionFormula As String
        Dim strponumber As String
        Dim strLocationList As String
        Dim strdate As String
        Dim strTypeOfIssue As String
        Dim lngCount As Integer
        Dim stritemslist As String
        Dim strsupplierlist As String
        Dim lngponumber As Integer
        Dim lngCheckeditems, lngCheckedlocations, lngCheckedsuppliers As Integer
        Dim straddress As String
        Dim strQSNo As String
        Dim datInvoice_Date As Date
        'Update Registry Settings for DSN to support two different databases from the same machine with same DSN Name
        Call UpdateRegistryDSNProperties(gstrCONNECTIONDSN, gstrCONNECTIONDATABASE, gstrCONNECTIONSERVER)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        frmRpt = New eMProCrystalReportViewer
        CRDoc = frmRpt.GetReportDocument
        CRDoc.Load(My.Application.Info.DirectoryPath & "\Reports\rptInvoiceDetails.rpt")
        frmRpt.ShowPrintButton = True
        frmRpt.ShowTextSearchButton = True
        frmRpt.ShowCloseButton = True
        'To Initialize the Report Window Title
        frmRpt.ReportHeader = Me.ctlFormHeader1.HeaderString()
        cmdOK_Click(cmdOk, New System.EventArgs())
        If strSelInvoices <> "" Then
            strLocationList = Replace(strSelInvoices, "|", ",")
            strLocationList = " {Sales_dtl.Unit_code}='" & gstrUNITID & "' and {Sales_dtl.Doc_No} IN [" & strLocationList & "] and "
        Else
            MsgBox("Select at least One Invoice.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            Me.lstInv.Focus()
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
            Exit Sub
        End If
        strsupplierlist = txtCustCode.Text
        strsupplierlist = "{saleschallan_dtl.account_code} IN ['" & strsupplierlist & "'] and "
        strLocationList = strLocationList & strsupplierlist
        strdate = Year(Me.dtpDateFrom.Value) & "," & Month(Me.dtpDateFrom.Value) & "," & VB.Day(Me.dtpDateFrom.Value)
        strLocationList = strLocationList & "{saleschallan_dtl.invoice_date} >= date(" & strdate & ")  and "
        strdate = Year(Me.dtpDateTo.Value) & "," & Month(Me.dtpDateTo.Value) & "," & VB.Day(Me.dtpDateTo.Value)
        strLocationList = strLocationList & "{saleschallan_dtl.invoice_date}  <= date(" & strdate & ")"
        straddress = gstr_WRK_ADDRESS1 & " " & gstr_WRK_ADDRESS2
        CRDoc.DataDefinition.FormulaFields("from").Text = "'" & Me.dtpDateFrom.Text & "'"
        CRDoc.DataDefinition.FormulaFields("to").Text = "'" & Me.dtpDateTo.Text & "'"
        CRDoc.DataDefinition.FormulaFields("comp_name").Text = "'" & gstrCOMPANY & "'"
        CRDoc.DataDefinition.FormulaFields("comp_address").Text = "'" & straddress & "'"
        CRDoc.RecordSelectionFormula = strLocationList
        frmRpt.Zoom = 150
        frmRpt.Show()
        frmRpt = Nothing
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        If Err.Number = 20545 Then
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            Resume Next
        Else
            frmRpt = Nothing
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End If
    End Sub
    Private Sub CmdRefNoHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdRefNoHelp.Click
        On Error GoTo ErrHandler
        Dim strrefno() As String
        Dim strRefSql As String
        strRefSql = "Select DISTINCT b.Cust_Ref,b.Amendment_No,a.po_type from Cust_Ord_hdr a,Cust_Ord_Dtl b"
        strRefSql = strRefSql & " where a.Unit_code=b.Unit_code and a.Unit_code='" & gstrUNITID & "' and b.Account_Code='" & Trim(txtCustCode.Text) & "' and b.Active_flag ='A' and "
        strRefSql = strRefSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and PO_type <> 'E'and "
        strRefSql = strRefSql & " a.Amendment_No = b.amendment_No  AND a.Authorized_Flag = 1 and b.Cust_drgNo = '"
        strRefSql = strRefSql & Trim(txtCustPartCode.Text) & "' and ITem_code = '" & Trim(lblItemCode.Text) & "'"
        strRefSql = strRefSql & " and a.Valid_date >'" & Format(GetServerDate, "dd MMM yyyy") & "' and effect_Date <= '"
        strRefSql = strRefSql & Format(GetServerDate, "dd MMM yyyy") & "' order by b.Cust_Ref,b.Amendment_No"
        strrefno = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strRefSql, "Refrence Details")
        If UBound(strrefno) < 0 Then Exit Sub
        If strrefno(0) = "0" Then
            MsgBox("No Reference available to Display", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100)) : txtRefNo.Text = "" : txtRefNo.Focus() : Exit Sub
        Else
            If strrefno(0) <> "" Then
                txtRefNo.Text = Trim(strrefno(0))
                txtAmendment.Text = Trim(strrefno(1))
                mstrSOType.Value = Trim(strrefno(2))
                Call ChangeInvTypeCaption(mstrSOType.Value)
                Call txtRefNo_Leave(txtRefNo, New System.EventArgs())
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdSADCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSADCodeHelp.Click
        Dim strHelp As String
        Dim strSTaxHelp() As String
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strHelp = "Select TxRT_Rate_No,TxRt_Percentage from Gen_taxRate where Unit_code='" & gstrUNITID & "' and Tx_TaxeID ='SAD' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                strSTaxHelp = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, "SAD Tax Help")
                If UBound(strSTaxHelp) < 0 Then Exit Sub
                If strSTaxHelp(0) = "0" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtSADCode.Text = "" : txtSADCode.Focus() : Exit Sub
                Else
                    If strSTaxHelp(0) <> "" Then
                        txtSADCode.Text = strSTaxHelp(0)
                        lblSAD_per.Text = strSTaxHelp(1)
                        txtSADCode.Focus()
                    End If
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub CmdSaleTaxType_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdSaleTaxType.Click
        Dim strHelp As String
        Dim strSTaxHelp() As String
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strHelp = "Select TxRT_Rate_No,TxRt_Percentage from Gen_taxRate where Unit_code='" & gstrUNITID & "' and (Tx_TaxeID ='CST' OR Tx_TaxeID ='LST' OR Tx_TaxeID ='VAT') and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                strSTaxHelp = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, "S.Tax Help")
                If UBound(strSTaxHelp) < 0 Then Exit Sub
                If strSTaxHelp(0) = "0" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtSaleTaxType.Text = "" : txtSaleTaxType.Focus() : Exit Sub
                Else
                    If strSTaxHelp(0) <> "" Then
                        txtSaleTaxType.Text = strSTaxHelp(0)
                        lblSaltax_Per.Text = strSTaxHelp(1)
                        txtSaleTaxType.Focus()
                    End If
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdSEcssCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSEcssCodeHelp.Click
        On Error GoTo Errorhandler
        Dim strHelp As String
        Dim strSSTaxHelp() As String
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        Select Case CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strHelp = "Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where Unit_code='" & gstrUNITID & "' and tx_TaxeID ='ECSSH' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                strSSTaxHelp = ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, "S. ECESS Tax Help")
                If UBound(strSSTaxHelp) < 0 Then Exit Sub
                If strSSTaxHelp(0) = "0" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    txtSEcssCode.Text = ""
                    txtSEcssCode.Focus()
                    Exit Sub
                Else
                    If strSSTaxHelp(0) <> "" Then
                        txtSEcssCode.Text = strSSTaxHelp(0)
                        lblSEcssCode.Text = strSSTaxHelp(1)
                        Call SetFormulaofColumns(0, 4)
                        Call ToShowDatainSummeryGrid()
                        txtSEcssCode.Focus()
                    End If
                End If
        End Select
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdSelectInvoice_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSelectInvoice.Click
        Dim objTempRs As New ADODB.Recordset
        If Len(Trim(txtCustPartCode.Text)) = 0 Then
            MsgBox("Please select the customer part code.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            Exit Sub
        End If
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            If FillInvoiceNumber() = True Then
                fraInvoice.Left = VB6.TwipsToPixelsX(2625)
                fraInvoice.Width = VB6.TwipsToPixelsX(4215)
                fraInvoice.Visible = True
                fraInvoice.BringToFront()
            Else
                MsgBox("No invoice for the selected Customer part.", MsgBoxStyle.Information, "Empower")
                Exit Sub
            End If
        ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            If strSelInvoices = "" Then
                strSelInvoices = "Select RefDoc_No from SupplementaryInv_Dtl Where Unit_code='" & gstrUNITID & "' and Doc_No = '" & txtChallanNo.Text & "'"
                objTempRs.Open(strSelInvoices, mP_Connection)
                strSelInvoices = objTempRs.GetString(ADODB.StringFormatEnum.adClipString, , , "|")
                If VB.Right(strSelInvoices, 1) = "|" Then
                    strSelInvoices = VB.Left(strSelInvoices, Len(strSelInvoices) - 1)
                End If
            End If
            If FillInvoiceNumber() = True Then
                fraInvoice.Left = VB6.TwipsToPixelsX(2625)
                fraInvoice.Width = VB6.TwipsToPixelsX(4215)
                fraInvoice.Visible = True
                fraInvoice.BringToFront()
            Else
                MsgBox("No invoice for the selected Customer part.", MsgBoxStyle.Information, "Empower")
                Exit Sub
            End If
        ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If strSelInvoices = "" Then
                strSelInvoices = "Select RefDoc_No from SupplementaryInv_Dtl Where Unit_code='" & gstrUNITID & "' and Doc_No = '" & txtChallanNo.Text & "'"
                objTempRs.Open(strSelInvoices, mP_Connection)
                strSelInvoices = objTempRs.GetString(ADODB.StringFormatEnum.adClipString, , , "|")
                If VB.Right(strSelInvoices, 1) = "|" Then
                    strSelInvoices = VB.Left(strSelInvoices, Len(strSelInvoices) - 1)
                End If
            End If
            If FillInvoiceNumber() = True Then
                fraInvoice.Left = VB6.TwipsToPixelsX(2625)
                fraInvoice.Width = VB6.TwipsToPixelsX(4215)
                fraInvoice.Visible = True
                fraInvoice.BringToFront()
                cmdClose.Enabled = True
            End If
        End If
    End Sub
    Private Sub cmdSurchargeTaxCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSurchargeTaxCode.Click
        On Error GoTo ErrHandler
        Dim strHelp As String
        Dim strSSTaxHelp() As String
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strHelp = "Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where Unit_code='" & gstrUNITID & "' and tx_TaxeID ='SST' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                strSSTaxHelp = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, "S.Sales Tax Help")
                If UBound(strSSTaxHelp) < 0 Then Exit Sub
                If strSSTaxHelp(0) = "0" Then
                    MsgBox("This Surcharge Tax Type is not correct . For help press F1.", MsgBoxStyle.Information, "eMpro")
                Else
                    If strSSTaxHelp(0) <> "" Then
                        txtSurchargeTaxType.Text = strSSTaxHelp(0)
                        lblSurcharge_Per.Text = strSSTaxHelp(1)
                        txtSurchargeTaxType.Focus()
                    End If
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub spdInvDetails_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles spdInvDetails.Enter
        On Error GoTo ErrHandler
        If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            Call ToShowDatainSummeryGrid()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub sstbInvoiceDtl_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles sstbInvoiceDtl.SelectedIndexChanged
        Static PreviousTab As Short = sstbInvoiceDtl.SelectedIndex()
        If sstbInvoiceDtl.SelectedIndex = 1 Then
            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.WaitCursor)
            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                bCheck = True
            End If
            Call ToShowDatainSummeryGrid()
            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
        End If
        PreviousTab = sstbInvoiceDtl.SelectedIndex()
    End Sub
    Private Sub txtAmendment_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAmendment.TextChanged
        On Error GoTo ErrHandler
        If Len(Trim(txtAmendment.Text)) = 0 Then
            txtRefNo.Text = ""
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtAmendment_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAmendment.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            cmdCustAmend.PerformClick()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtAmendment_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAmendment.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
            Case System.Windows.Forms.Keys.Return
                txtCustRefRemarks.Focus()
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtAmendment_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAmendment.Leave
        On Error GoTo ErrHandler
        Dim clsrrefno As ClsResultSetDB
        Dim strRefSql As String
        If Len(Trim(txtAmendment.Text)) = 0 Then Exit Sub
        If Len(Trim(txtRefNo.Text)) = 0 Then
            MsgBox("First select Customer Referance", MsgBoxStyle.Information, "empower")
            txtAmendment.Text = ""
            txtRefNo.Focus()
            Exit Sub
        End If
        strRefSql = "Select b.Cust_Ref,b.Amendment_No,a.po_type from Cust_Ord_hdr a,Cust_Ord_Dtl b"
        strRefSql = strRefSql & " where a.Unit_code=b.unit_code and a.Unit_code='" & gstrUNITID & "' and b.Account_Code='" & Trim(txtCustCode.Text) & "' and b.Active_flag ='A' and "
        strRefSql = strRefSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and PO_type <> 'E'and "
        strRefSql = strRefSql & " a.Amendment_No = b.amendment_No  AND a.Authorized_Flag = 1 and b.Cust_drgNo = '"
        strRefSql = strRefSql & Trim(txtCustPartCode.Text) & "' and ITem_code = '" & Trim(lblItemCode.Text) & "'"
        strRefSql = strRefSql & " and a.Valid_date >'" & Format(GetServerDate, "dd MMM yyyy") & "' and effect_Date <= '"
        strRefSql = strRefSql & Format(GetServerDate, "dd MMM yyyy") & "' and b.Cust_ref = '" & Trim(txtRefNo.Text) & "' and b.Amendment_no = '" & Trim(txtAmendment.Text) & "'"
        strRefSql = strRefSql & " order by b.Cust_Ref,b.Amendment_No"
        clsrrefno = New ClsResultSetDB
        clsrrefno.GetResult(strRefSql)
        If clsrrefno.GetNoRows = 0 Then
            MsgBox("No Amendment available to Display", MsgBoxStyle.Information) : txtRefNo.Text = "" : txtRefNo.Focus() : Exit Sub
        Else
            txtRefNo.Text = clsrrefno.GetValue("Cust_ref")
            txtAmendment.Text = clsrrefno.GetValue("amendment_no")
            mstrSOType.Value = Trim(clsrrefno.GetValue("po_type"))
            clsrrefno.ResultSetClose()
            clsrrefno = Nothing
            Call ChangeInvTypeCaption(mstrSOType.Value)
            Call DisplayNewSOHdrDetails()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtChallanNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtChallanNo.TextChanged
        On Error GoTo ErrHandler
        If Len(Trim(txtChallanNo.Text)) = 0 Then
            txtCustCode.Text = "" : txtCustPartCode.Text = "" : txtRefNo.Text = "" : txtAmendment.Text = ""
            txtCustRefRemarks.Text = "" : txtRemarks.Text = "" : txtCVDCode.Text = "" : txtSADCode.Text = "" : txtExciseTaxType.Text = ""
            txtSaleTaxType.Text = "" : txtSurchargeTaxType.Text = ""
            '101188073
            txtCGSTType.Text = "" : txtSGSTType.Text = "" : txtUTGSTType.Text = ""
            txtIGSTType.Text = "" : txtCompCessType.Text = "" : lblHSNSAC.Text = "" : lblHSNSACCODE.Text = ""
            '101188073
            spdPrevInv.MaxRows = 0
            spdInvDetails.MaxRows = 0
            spdPrevInv.MaxRows = 1
            spdInvDetails.MaxRows = 1
            lblCurrencyDes.Text = ""
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtCustCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustCode.TextChanged
        On Error GoTo ErrHandler
        If Len(Trim(txtCustCode.Text)) = 0 Then
            lblCustCodeDes.Text = "" : lblAddressDes.Text = ""
            txtCustPartCode.Text = "" : txtRefNo.Text = "" : txtAmendment.Text = ""
            txtCustRefRemarks.Text = "" : txtRemarks.Text = "" : txtCVDCode.Text = "" : txtSADCode.Text = "" : txtExciseTaxType.Text = ""
            txtSaleTaxType.Text = "" : txtSurchargeTaxType.Text = ""
            '101188073
            txtCGSTType.Text = "" : txtSGSTType.Text = "" : txtUTGSTType.Text = ""
            txtIGSTType.Text = "" : txtCompCessType.Text = "" : lblHSNSAC.Text = "" : lblHSNSACCODE.Text = ""
            '101188073
            spdPrevInv.MaxRows = 1
            spdPrevInv.MaxRows = 1
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtCustPartCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustPartCode.TextChanged
        On Error GoTo ErrHandler
        If Len(Trim(txtCustPartCode.Text)) = 0 Then
            lblItemCode.Text = "" : txtRefNo.Text = "" : txtAmendment.Text = ""
            txtCustRefRemarks.Text = "" : txtRemarks.Text = "" : txtCVDCode.Text = "" : txtSADCode.Text = "" : txtExciseTaxType.Text = ""
            txtSaleTaxType.Text = "" : txtSurchargeTaxType.Text = ""
            '101188073
            txtCGSTType.Text = "" : txtSGSTType.Text = "" : txtUTGSTType.Text = ""
            txtIGSTType.Text = "" : txtCompCessType.Text = "" : lblHSNSAC.Text = "" : lblHSNSACCODE.Text = ""
            '101188073
            spdPrevInv.MaxRows = 1
            spdPrevInv.MaxRows = 1
        End If
        strSelInvoices = ""
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtCustPartCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustPartCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call cmdcustPartCode_Click(cmdcustPartCode, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtCustPartCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustPartCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
            Case System.Windows.Forms.Keys.Return
                Select Case CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        txtRefNo.Focus()
                End Select
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCustPartCode_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustPartCode.Leave
        Dim strCustItemMst As String
        Dim strCustItemHelp As String
        Dim intNoOfRecords As Short
        Dim rsCustItemMst As ClsResultSetDB
        Dim strCustItem() As String
        Dim objTempRs As New ADODB.Recordset
        If bool_valid = True Then
            Exit Sub
        End If
        On Error GoTo ErrHandler
        If Len(Trim(txtCustPartCode.Text)) = 0 Then Exit Sub
        If Len(Trim(txtLocationCode.Text)) = 0 Then
            Call ConfirmWindow(10116, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            If txtLocationCode.Enabled Then txtLocationCode.Focus()
            Exit Sub
        End If
        If Len(Trim(txtCustCode.Text)) = 0 Then
            MsgBox("First Select Customer Code.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            If txtCustCode.Enabled Then txtCustCode.Focus()
            Exit Sub
        End If
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                strCustItemHelp = "select distinct Cust_Item_code,Item_code,Cust_Item_desc from Saleschallan_dtl a,sales_dtl b "
                strCustItemHelp = strCustItemHelp & " where a.Unit_code = B.Unit_code and a.Unit_code='" & gstrUNITID & "' and Account_code = '" & txtCustCode.Text & "' and invoice_date > = '" & Format(dtpDateFrom.Value, "dd MMM yyyy")
                strCustItemHelp = strCustItemHelp & "' and invoice_date < = '" & Format(dtpDateTo.Value, "dd MMM yyyy") & "' and a.Doc_no = b.Doc_no and bill_flag = 1 and cancel_flag =0"
                strCustItemHelp = strCustItemHelp & " and cust_item_code = '" & Trim(txtCustPartCode.Text) & "'"
                rsCustItemMst = New ClsResultSetDB
                rsCustItemMst.GetResult(strCustItemHelp)
                intNoOfRecords = rsCustItemMst.GetNoRows
                If intNoOfRecords > 0 Then
                    txtEcssCode.Text = "EC2"
                    bool_valid = True
                    txtEcssCode_Validating(txtEcssCode, New System.ComponentModel.CancelEventArgs(False))
                    bool_valid = False
                    If intNoOfRecords = 1 Then
                        lblItemCode.Text = rsCustItemMst.GetValue("Item_Code")
                        lblCustItemDesc.Text = rsCustItemMst.GetValue("Cust_item_desc")
                        Call SetCellTypeofGrids()
                        Call SetMaxLengthofGrid(4)
                        Call DisplayDetailsInAddMode()
                        fraSelectUnSelect.Enabled = True : chkSelectAll.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
                        chkUnCheckall.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
                    ElseIf intNoOfRecords > 1 Then
                        MsgBox("There are more then one records for selected Customer Part Code,You can select from List,Press F1.")
                        rsCustItemMst.MoveFirst()
                        lblCustItemDesc.Text = rsCustItemMst.GetValue("Cust_item_desc")
                        lblItemCode.Text = rsCustItemMst.GetValue("Item_Code")
                        Call SetCellTypeofGrids()
                        Call SetMaxLengthofGrid(4)
                        Call DisplayDetailsInAddMode()
                        fraSelectUnSelect.Enabled = True : chkSelectAll.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
                        chkUnCheckall.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
                        rsCustItemMst.ResultSetClose()
                        Exit Sub
                    End If
                    txtCustPartCode_Validating(txtCustPartCode, New System.ComponentModel.CancelEventArgs(False))
                    txtRefNo.Focus()
                Else
                    MsgBox("Invalid Customer Part code.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    txtCustPartCode.Text = ""
                    txtCustPartCode.Focus()
                    rsCustItemMst.ResultSetClose()
                    Exit Sub
                End If
                rsCustItemMst.ResultSetClose()
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub CmdGrpChEnt_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles CmdGrpChEnt.ButtonClick
        On Error GoTo ErrHandler
        Dim strSQL As String
        Dim strToCheckPrevInvoices As String
        Dim rsToCheckPrevInvoices As ClsResultSetDB
        Dim intCase As String
        Dim strMessageString As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim strInputBox As String
        Select Case eventArgs.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                Call SetFinancialYearDates()
                Call EnableControls(True, Me, True)
                fraSelectUnSelect.Enabled = False : chkSelectAll.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                chkUnCheckall.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                If blnEOU_FLAG = True Then
                    txtCVDCode.Enabled = True : txtCVDCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdCVDCodeHelp.Enabled = True
                    txtSADCode.Enabled = True : txtSADCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdSADCodeHelp.Enabled = True
                Else
                    txtCVDCode.Enabled = False : txtCVDCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdCVDCodeHelp.Enabled = False
                    txtSADCode.Enabled = False : txtSADCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdSADCodeHelp.Enabled = False
                End If
                '101188073
                If gblnGSTUnit Then
                    txtExciseTaxType.Enabled = False : txtExciseTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : CmdExciseTaxType.Enabled = False
                    lblExctax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtSaleTaxType.Enabled = False : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : CmdSaleTaxType.Enabled = False
                    lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtSurchargeTaxType.Enabled = False : txtSurchargeTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdSurchargeTaxCode.Enabled = False
                    lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtEcssCode.Enabled = False : txtEcssCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdEcssCodeHelp.Enabled = False
                    lblEcssCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtSEcssCode.Enabled = False : txtSEcssCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdSEcssCodeHelp.Enabled = False
                    lblSEcssCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtCVDCode.Enabled = False : txtCVDCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdCVDCodeHelp.Enabled = False
                    lblCVD_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtSADCode.Enabled = False : txtSADCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdSADCodeHelp.Enabled = False
                    lblSAD_per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                Else
                    txtCGSTType.Enabled = False : txtCGSTType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdCGSTType.Enabled = False
                    lblCGSTPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtSGSTType.Enabled = False : txtSGSTType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdSGSTType.Enabled = False
                    lblSGSTPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtUTGSTType.Enabled = False : txtUTGSTType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdUTGSTType.Enabled = False
                    lblUTGSTPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtIGSTType.Enabled = False : txtIGSTType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdIGSTType.Enabled = False
                    lblIGSTPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtCompCessType.Enabled = False : txtCompCessType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdCompCessType.Enabled = False
                    lblCompCessPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                End If
                '101188073
                Call SelectChallanNoFromSupplementatryInvHdr()
                txtChallanNo.Enabled = False : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : CmdChallanNo.Enabled = False
                txtChallanNo.Enabled = False : lblCustCodeDes.Text = ""
                With dtpDateFrom
                    .Value = GetServerDate()
                    .Visible = True 'Show DatePicker
                End With
                With dtpDateTo
                    .Value = GetServerDate()
                    .Visible = True 'Show DatePicker
                End With
                frareportName.Enabled = False : optsupp.Enabled = False : optcredit.Enabled = False
                txtLocationCode.Text = strDefaultLocation
                Me.chkShowDetailAnexture.Enabled = True
                Me.chkShowDetailAnexture.CheckState = System.Windows.Forms.CheckState.Checked
                Call HideMRPlabel(True)
                txtMRP.Text = "0.0000"
                Call ChangeInvTypeCaption("")
                dtpDateFrom.Focus()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                Call frmMKTTRN0018_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
                chkShowDetailAnexture.Enabled = True
                chkShowDetailAnexture.CheckState = System.Windows.Forms.CheckState.Checked
                Call ChangeInvTypeCaption("")
                Call HideMRPlabel(False)
                mstrCustRefNo = ""
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                If ValidatebeforeSave() = False Then Exit Sub
                If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    If Len(Trim(txtRefNo.Text)) > 0 Then
                        strToCheckPrevInvoices = "select * from ReturnPrevSupplementaryData('" & gstrUNITID & "','"
                        strToCheckPrevInvoices = strToCheckPrevInvoices & Format(dtpDateFrom.Value, "dd MMM yyyy") & "','"
                        strToCheckPrevInvoices = strToCheckPrevInvoices & Format(dtpDateTo.Value, "dd MMM yyyy") & "','"
                        strToCheckPrevInvoices = strToCheckPrevInvoices & Trim(txtCustCode.Text) & "','" & Trim(txtCustPartCode.Text)
                        strToCheckPrevInvoices = strToCheckPrevInvoices & "','" & txtRefNo.Text & "')"
                    Else
                        strToCheckPrevInvoices = "select * from ReturnPrevSupplementaryData('" & gstrUNITID & "','"
                        strToCheckPrevInvoices = strToCheckPrevInvoices & Format(dtpDateFrom.Value, "dd MMM yyyy") & "','"
                        strToCheckPrevInvoices = strToCheckPrevInvoices & Format(dtpDateTo.Value, "dd MMM yyyy") & "','"
                        strToCheckPrevInvoices = strToCheckPrevInvoices & Trim(txtCustCode.Text) & "','" & Trim(txtCustPartCode.Text)
                        strToCheckPrevInvoices = strToCheckPrevInvoices & "','')"
                    End If
                    rsToCheckPrevInvoices = New ClsResultSetDB
                    rsToCheckPrevInvoices.GetResult(strToCheckPrevInvoices)
                    If rsToCheckPrevInvoices.GetNoRows > 0 Then
                        rsToCheckPrevInvoices.MoveFirst()
                        intCase = rsToCheckPrevInvoices.GetValue("CaseNo")
                        Select Case intCase
                            Case CStr(1)
                                strMessageString = "Following Invoice already Generated With the Same Range,Customer,Customer Part Code,Referance No Would You Like to Proceed ?" & vbCrLf
                            Case CStr(2)
                                strMessageString = "Following Invoice already Generated With the Same Range,Customer,Customer Part Code Would You Like to Proceed ?" & vbCrLf
                            Case CStr(3)
                                strMessageString = "Following Invoice already Generated With the Same Range,Customer,Referance No. Would You Like to Proceed ?" & vbCrLf
                            Case CStr(4)
                                strMessageString = "Following Invoice already Generated With the Same Customer,Customer Part Code,Referance No Would You Like to Proceed ?" & vbCrLf
                            Case CStr(5)
                                strMessageString = "Following Invoice already Generated With the Same Range,Customer,Customer Part Code Would You Like to Proceed ?" & vbCrLf
                        End Select
                        strMessageString = strMessageString & "Invoice No       Invoice Date" & vbCrLf
                        intMaxLoop = rsToCheckPrevInvoices.GetNoRows
                        rsToCheckPrevInvoices.MoveFirst()
                        For intLoopCounter = 1 To intMaxLoop
                            strMessageString = strMessageString & rsToCheckPrevInvoices.GetValue("Doc_no") & "         "
                            strMessageString = strMessageString & VB6.Format(rsToCheckPrevInvoices.GetValue("Invoice_date")) & vbCrLf
                            rsToCheckPrevInvoices.MoveNext()
                        Next
                        If MsgBox(strMessageString, MsgBoxStyle.YesNo, "empower") = MsgBoxResult.No Then
                            rsToCheckPrevInvoices.ResultSetClose()
                            Exit Sub
                        End If
                    End If
                    rsToCheckPrevInvoices.ResultSetClose()
                End If
                Call SaveData()
                Me.CmdGrpChEnt.Revert()
                gblnCancelUnload = False : gblnFormAddEdit = False
                Call EnableControls(False, Me)
                cmdSelectInvoice.Enabled = True 'to enable the select invoice button in view mode
                lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                lblCVD_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                lblSAD.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                lblExctax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                spdPrevInv.Enabled = True
                spdPrevInv.MaxRows = 1
                spdPrevInv.Row = 1
                spdPrevInv.Row2 = spdPrevInv.MaxRows
                spdPrevInv.Col = 0
                spdPrevInv.Col2 = 12
                spdPrevInv.BlockMode = True
                spdPrevInv.Lock = True
                spdPrevInv.BlockMode = False
                txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                CmdLocCodeHelp.Enabled = True : CmdChallanNo.Enabled = True 'Me.SpChEntry.Enabled = False
                If txtLocationCode.Enabled Then
                    txtLocationCode.Focus()
                End If
                Me.chkShowDetailAnexture.Enabled = True
                Me.chkShowDetailAnexture.CheckState = System.Windows.Forms.CheckState.Checked
                Call HideMRPlabel(False)
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                Call EnableControls(True, Me, False)
                fraSelectUnSelect.Enabled = False : chkSelectAll.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                chkUnCheckall.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                If blnEOU_FLAG = True Then
                    txtCVDCode.Enabled = True : txtCVDCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdCVDCodeHelp.Enabled = True
                    txtSADCode.Enabled = True : txtSADCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdSADCodeHelp.Enabled = True
                Else
                    txtCVDCode.Enabled = False : txtCVDCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdCVDCodeHelp.Enabled = False
                    txtSADCode.Enabled = False : txtSADCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdSADCodeHelp.Enabled = False
                End If
                '101188073
                If gblnGSTUnit Then
                    txtExciseTaxType.Enabled = False : txtExciseTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : CmdExciseTaxType.Enabled = False
                    lblExctax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtSaleTaxType.Enabled = False : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : CmdSaleTaxType.Enabled = False
                    lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtSurchargeTaxType.Enabled = False : txtSurchargeTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdSurchargeTaxCode.Enabled = False
                    lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtEcssCode.Enabled = False : txtEcssCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdEcssCodeHelp.Enabled = False
                    lblEcssCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtSEcssCode.Enabled = False : txtSEcssCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdSEcssCodeHelp.Enabled = False
                    lblSEcssCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtCVDCode.Enabled = False : txtCVDCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdCVDCodeHelp.Enabled = False
                    lblCVD_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtSADCode.Enabled = False : txtSADCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdSADCodeHelp.Enabled = False
                    lblSAD_per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                Else
                    txtCGSTType.Enabled = False : txtCGSTType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdCGSTType.Enabled = False
                    lblCGSTPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtSGSTType.Enabled = False : txtSGSTType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdSGSTType.Enabled = False
                    lblSGSTPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtUTGSTType.Enabled = False : txtUTGSTType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdUTGSTType.Enabled = False
                    lblUTGSTPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtIGSTType.Enabled = False : txtIGSTType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdIGSTType.Enabled = False
                    lblIGSTPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtCompCessType.Enabled = False : txtCompCessType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdCompCessType.Enabled = False
                    lblCompCessPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                End If
                '101188073
                DisplayDetailsinEditMode()
                fraSelectUnSelect.Enabled = True : chkSelectAll.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
                chkUnCheckall.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
                dtpDateFrom.Enabled = False : dtpDateTo.Enabled = False : txtLocationCode.Enabled = False
                txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtChallanNo.Enabled = False : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtCustCode.Enabled = False : txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtCustPartCode.Enabled = False
                txtCustPartCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                Call ToLockCellofGrid()
                frareportName.Enabled = False : optsupp.Enabled = False : optcredit.Enabled = False
                Me.chkShowDetailAnexture.Enabled = True
                Me.chkShowDetailAnexture.CheckState = System.Windows.Forms.CheckState.Checked
                Call HideMRPlabel(False)
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
                If ConfirmWindow(10054, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    Call DeleteRecords()
                    Call EnableControls(False, Me, True)
                    lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    lblCVD_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    lblSAD.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    lblExctax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    spdPrevInv.Enabled = True
                    spdPrevInv.MaxRows = 0
                    spdPrevInv.MaxRows = 1
                    spdPrevInv.Row = 1
                    spdPrevInv.Row2 = spdPrevInv.MaxRows
                    spdPrevInv.Col = 0
                    spdPrevInv.Col2 = 12
                    spdPrevInv.BlockMode = True
                    spdPrevInv.Lock = True
                    spdPrevInv.BlockMode = False
                    txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    CmdLocCodeHelp.Enabled = True : CmdChallanNo.Enabled = True
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    CmdGrpChEnt.Enabled(2) = False
                    If txtLocationCode.Enabled Then
                        txtLocationCode.Focus()
                    End If
                End If
                chkShowDetailAnexture.Enabled = True
                chkShowDetailAnexture.CheckState = System.Windows.Forms.CheckState.Checked
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
                If txtChallanNo.Text = "" Then
                    MessageBox.Show("Please enter challan no.", "eMPro", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    txtChallanNo.Focus()
                    Exit Sub
                End If
                strSQL = "{SupplementaryInv_hdr.Unit_Code}='" & gstrUNITID & "' and {SupplementaryInv_hdr.Location_Code}='" & Trim(txtLocationCode.Text) & "' and {SupplementaryInv_hdr.Doc_No} =" & Trim(txtChallanNo.Text)
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                frmRpt = New eMProCrystalReportViewer
                CRDoc = frmRpt.GetReportDocument
                Dim strRptfilename As String
                If optsupp.Checked = True Then
                    If chkShowDetailAnexture.CheckState = System.Windows.Forms.CheckState.Checked Then
                        strRptfilename = My.Application.Info.DirectoryPath & "\Reports\rptSuppInvAnnexure.rpt"
                    Else
                        strRptfilename = My.Application.Info.DirectoryPath & "\Reports\rptSuppInvAnnexureSummary.rpt"
                    End If
                Else
                    strRptfilename = My.Application.Info.DirectoryPath & "\Reports\rptSuppCrAdvise.rpt"
                End If
                CRDoc.Load(strRptfilename)
                frmRpt.ShowPrintButton = True
                frmRpt.ShowTextSearchButton = True
                CRDoc.RecordSelectionFormula = strSQL
                frmRpt.Show()
                frmRpt = Nothing
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub CmdLocCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdLocCodeHelp.Click
        Dim strLocationCode() As String
        On Error GoTo ErrHandler
        Select Case CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                strLocationCode = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select DISTINCT cast(s.Location_Code as varchar(10)) as Location_Code ,l.Description from Location_Mst l,SaleConf s where s.Unit_code=l.Unit_code and s.Unit_code='" & gstrUNITID & "' and s.Location_code = l.Location_code and s.Location_code like'" & txtLocationCode.Text & "%'", "Accounting Locations")
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                strLocationCode = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select DISTINCT cast(s.Location_Code as varchar(10)) as Location_Code,l.Description from Location_Mst l,SupplementaryInv_hdr s where s.Unit_code = l.Unit_code and s.Unit_code='" & gstrUNITID & "' and s.Location_code = l.Location_code and s.Location_code like'" & txtLocationCode.Text & "%'", "Accounting Locations")
        End Select
        If UBound(strLocationCode) < 0 Then Exit Sub
        If strLocationCode(0) = "0" Then
            MsgBox("No Accounting Location Available to Display.") : txtLocationCode.Text = "" : txtLocationCode.Focus() : Exit Sub
        Else
            If strLocationCode(0) <> "" Then
                txtLocationCode.Text = strLocationCode(0)
                strDefaultLocation = Trim(txtLocationCode.Text)
            End If
        End If
        Call txtLocationCode_Leave(txtLocationCode, New System.EventArgs())
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0018_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0018_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0018_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Err_Handler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlFormHeader1_Click(ctlFormHeader1, New System.EventArgs()) : Exit Sub
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0018_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                If Me.CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    If fraInvoice.Visible = True Then
                        strSelInvoices = ""
                        fraInvoice.Visible = False
                    End If
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call Me.CmdGrpChEnt.Revert()
                        Call EnableControls(False, Me, True)
                        gblnCancelUnload = False
                        gblnFormAddEdit = False
                        spdPrevInv.MaxRows = 0
                        spdInvDetails.MaxRows = 0
                        spdPrevInv.MaxRows = 1
                        spdInvDetails.MaxRows = 1
                        dtpDateFrom.Value = GetServerDate() : dtpDateTo.Value = GetServerDate()
                        txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdLocCodeHelp.Enabled = True
                        txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdChallanNo.Enabled = True
                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                        CmdGrpChEnt.Enabled(2) = False
                        lblCurrencyDes.Text = "" : lblCustItemDesc.Text = ""
                        txtLocationCode.Text = strDefaultLocation
                        txtLocationCode.Focus()
                    Else
                        Me.ActiveControl.Focus()
                    End If
                End If
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub frmMKTTRN0018_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        Dim rsCompanyMst As ClsResultSetDB
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FillLabelFromResFile(Me) 'Fill Labels >From Resource File
        Call FitToClient(Me, FraChEnt, ctlFormHeader1, CmdGrpChEnt, 500)
        Call EnableControls(False, Me, True)
        sstbInvoiceDtl.Enabled = True : spdPrevInv.Enabled = True : spdInvDetails.Enabled = True
        dtpDateFrom.Format = DateTimePickerFormat.Custom
        dtpDateTo.Format = DateTimePickerFormat.Custom
        dtpDateFrom.CustomFormat = gstrDateFormat
        dtpDateTo.CustomFormat = gstrDateFormat
        sstbInvoiceDtl.SelectedIndex = 0
        If CheckFinancialYearDates() = True Then
            Call SetFinancialYearDates()
        End If
        txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdLocCodeHelp.Enabled = True
        txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdChallanNo.Enabled = True
        CmdLocCodeHelp.Image = My.Resources.ico111.ToBitmap
        CmdChallanNo.Image = My.Resources.ico111.ToBitmap
        CmdCustCodeHelp.Image = My.Resources.ico111.ToBitmap
        cmdcustPartCode.Image = My.Resources.ico111.ToBitmap
        CmdRefNoHelp.Image = My.Resources.ico111.ToBitmap
        cmdCustAmend.Image = My.Resources.ico111.ToBitmap
        cmdCVDCodeHelp.Image = My.Resources.ico111.ToBitmap
        cmdSADCodeHelp.Image = My.Resources.ico111.ToBitmap
        CmdExciseTaxType.Image = My.Resources.ico111.ToBitmap
        CmdSaleTaxType.Image = My.Resources.ico111.ToBitmap
        cmdSurchargeTaxCode.Image = My.Resources.ico111.ToBitmap
        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        CmdGrpChEnt.Enabled(2) = False
        spdPrevInv.SetRefStyle(2)
        Call AddHeadersOfGrids()
        Call SetWidthofColumnsinGrid()
        spdPrevInv.MaxRows = 0
        spdInvDetails.MaxRows = 0
        spdPrevInv.MaxRows = 1
        spdInvDetails.MaxRows = 1
        rsCompanyMst = New ClsResultSetDB
        rsCompanyMst.GetResult("Select EOU_Flag from Company_Mst where Unit_code = '" & gstrUNITID & "' ")
        If rsCompanyMst.GetNoRows > 0 Then
            blnEOU_FLAG = rsCompanyMst.GetValue("EOU_Flag")
            rsCompanyMst.ResultSetClose()
            If blnEOU_FLAG = True Then
                txtCVDCode.Enabled = True : txtCVDCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdCVDCodeHelp.Enabled = True
                txtSADCode.Enabled = True : txtSADCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdSADCodeHelp.Enabled = True
                With spdPrevInv
                    .Col = enumPreInvoiceDetails.NewCVDValue
                    .Col2 = enumPreInvoiceDetails.NewCVDValue
                    .Row = 1
                    .Row2 = .MaxRows
                    .BlockMode = True
                    .ColHidden = False
                    .BlockMode = False
                    .Col = enumPreInvoiceDetails.NewSADValue
                    .Col2 = enumPreInvoiceDetails.NewSADValue
                    .Row = 1
                    .Row2 = .MaxRows
                    .BlockMode = True
                    .ColHidden = False
                    .BlockMode = False
                End With
            Else
                txtCVDCode.Enabled = False : txtCVDCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdCVDCodeHelp.Enabled = False
                txtSADCode.Enabled = False : txtSADCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdSADCodeHelp.Enabled = False
                With spdPrevInv
                    .Col = enumPreInvoiceDetails.NewCVDValue
                    .Col2 = enumPreInvoiceDetails.NewCVDValue
                    .Row = 1
                    .Row2 = .MaxRows
                    .BlockMode = True
                    .ColHidden = True
                    .BlockMode = False
                    .Col = enumPreInvoiceDetails.NewSADValue
                    .Col2 = enumPreInvoiceDetails.NewSADValue
                    .Row = 1
                    .Row2 = .MaxRows
                    .BlockMode = True
                    .ColHidden = True
                    .BlockMode = False
                End With
            End If
        End If
        '101188073
        If gblnGSTUnit Then
            lblHSNSACCODE.Visible = True
            lblHSN.Visible = True
            lblHSNSAC.Visible = True
            lblHSNSACCODE.Text = ""
            lblHSNSAC.Text = ""
            With spdPrevInv
                .BlockMode = True
                .Col = enumPreInvoiceDetails.CGST_AMT
                .Col2 = enumPreInvoiceDetails.DIFF_CCESS_AMT
                .Row = 1
                .Row2 = .MaxRows
                .ColHidden = False
                .BlockMode = False
            End With
            With spdInvDetails
                .BlockMode = True
                .Col = enumInvoiceSummery.CGST
                .Col2 = enumInvoiceSummery.CCESS
                .Row = 1
                .Row2 = .MaxRows
                .ColHidden = False
                .BlockMode = False
            End With
        Else
            lblHSNSACCODE.Visible = False
            lblHSN.Visible = False
            lblHSNSAC.Visible = False
            With spdPrevInv
                .BlockMode = True
                .Col = enumPreInvoiceDetails.CGST_AMT
                .Col2 = enumPreInvoiceDetails.DIFF_CCESS_AMT
                .Row = 1
                .Row2 = .MaxRows
                .ColHidden = True
                .BlockMode = False
            End With
            With spdInvDetails
                .BlockMode = True
                .Col = enumInvoiceSummery.CGST
                .Col2 = enumInvoiceSummery.CCESS
                .Row = 1
                .Row2 = .MaxRows
                .ColHidden = True
                .BlockMode = False
            End With
        End If
        '101188073
        Call SetCellTypeofGrids()
        Call SetMaxLengthofGrid(4)
        Call SetFormulaofColumns(0, 4)
        bCheck = False
        ''--- ToShowDatainSummeryGrid
        Call HideMRPlabel(False)
        strDefaultLocation = ""
        chkShowDetailAnexture.Enabled = True
        chkShowDetailAnexture.CheckState = System.Windows.Forms.CheckState.Checked
        SetGridDateFormat(spdPrevInv, enumPreInvoiceDetails.Invoice_Date)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0018_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        On Error GoTo ErrHandler
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
            If Me.CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                enmValue = ConfirmWindow(10055, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Or enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call CmdGrpChEnt_ButtonClick(CmdGrpChEnt, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE))
                    ElseIf enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Then
                        gblnCancelUnload = False
                        gblnFormAddEdit = False
                    End If
                Else
                    gblnCancelUnload = True
                    gblnFormAddEdit = True
                End If
            Else
                Me.Dispose()
                Exit Sub
            End If
        End If
        If gblnCancelUnload = True Then eventArgs.Cancel = 1
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTTRN0018_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Me.Dispose()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SelectDescriptionForField(ByRef pstrFieldName1 As String, ByRef pstrFieldName2 As String, ByRef pstrTableName As String, ByRef pContrName As System.Windows.Forms.Control, ByRef pstrControlText As String)
        On Error GoTo ErrHandler
        Dim strDesSql As String 'Declared to make Select Query
        Dim rsDescription As ClsResultSetDB
        strDesSql = "Select " & Trim(pstrFieldName1) & " from " & Trim(pstrTableName) & " where " & Trim(pstrFieldName2) & "='" & Trim(pstrControlText) & "' and Unit_code='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
        rsDescription = New ClsResultSetDB
        rsDescription.GetResult(strDesSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsDescription.GetNoRows > 0 Then
            pContrName.Text = rsDescription.GetValue(Trim(pstrFieldName1))
        End If
        rsDescription.ResultSetClose()
        rsDescription = Nothing
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtChallanNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtChallanNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            Select Case CmdGrpChEnt.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    CmdGrpChEnt.Focus()
            End Select
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtChallanNo_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtChallanNo.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call CmdChallanNo_Click(CmdChallanNo, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtChallanNo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtChallanNo.Leave
        Dim strInvoiceType As String
        Dim strCondition As String
        Dim rsSalesChallan As ClsResultSetDB
        Dim objTempRs As New ADODB.Recordset
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                If Len(txtChallanNo.Text) > 0 Then
                    cmdSelectInvoice.Enabled = True
                    strCondition = " Location_code = '" & Trim(txtLocationCode.Text) & "' and Unit_code = '" & gstrUNITID & "' "
                    If CheckExistanceOfFieldData((txtChallanNo.Text), "Doc_No", "supplementaryInv_hdr", strCondition) Then
                        'If Challan No. Exists
                        'Get Data From Challan_Dtl,Cust_Ord_Dtl,Sales_Dtl
                        If Len(txtLocationCode.Text) > 0 Then
                            If DisplayDetailsinViewMode() Then 'if record found
                                rsSalesChallan = New ClsResultSetDB
                                rsSalesChallan.GetResult("Select Bill_Flag from SupplementaryInv_hdr where Unit_code = '" & gstrUNITID & "' and Location_Code = '" & txtLocationCode.Text & "' and Doc_No = " & txtChallanNo.Text)
                                If rsSalesChallan.GetNoRows > 0 Then
                                    frareportName.Enabled = True : optsupp.Enabled = True : optcredit.Enabled = True : optsupp.Checked = True
                                    If rsSalesChallan.GetValue("Bill_Flag") = True Then
                                        CmdGrpChEnt.Enabled(1) = False
                                        CmdGrpChEnt.Enabled(2) = False
                                    Else
                                        CmdGrpChEnt.Enabled(1) = True
                                        CmdGrpChEnt.Enabled(2) = True
                                        CmdGrpChEnt.Enabled(5) = True
                                    End If
                                End If
                                rsSalesChallan.ResultSetClose()
                                strSelInvoices = "Select RefDoc_No from SupplementaryInv_Dtl Where Unit_code = '" & gstrUNITID & "' and Doc_No = '" & txtChallanNo.Text & "' order By RefDoc_No"
                                objTempRs.Open(strSelInvoices, mP_Connection)
                                strSelInvoices = objTempRs.GetString(ADODB.StringFormatEnum.adClipString, , , "|")
                                If VB.Right(strSelInvoices, 1) = "|" Then
                                    strSelInvoices = VB.Left(strSelInvoices, Len(strSelInvoices) - 1)
                                End If
                            Else 'if no record found then display message
                                Call ConfirmWindow(10414, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                                txtChallanNo.Focus()
                                Exit Sub
                            End If
                        Else 'if location code field is blank
                            Call ConfirmWindow(10239, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            txtLocationCode.Focus()
                            Exit Sub
                        End If
                    Else 'If Doc_No Is Invalid
                        Call ConfirmWindow(10404, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        txtChallanNo.Text = "" : txtChallanNo.Focus() : Exit Sub
                    End If
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If FillInvoiceNumber() = True Then
                    fraInvoice.Left = VB6.TwipsToPixelsX(2625)
                    fraInvoice.Width = VB6.TwipsToPixelsX(4215)
                    fraInvoice.Visible = True
                    fraInvoice.BringToFront()
                Else
                    MsgBox("No invoice for the selected Customer part.", MsgBoxStyle.Information, "Empower")
                    Exit Sub
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtCustCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            Select Case CmdGrpChEnt.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    txtCustPartCode.Focus()
            End Select
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtcustcode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call CmdCustCodeHelp_Click(CmdCustCodeHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtCustCode_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustCode.Leave
        Dim rsCustMst As ClsResultSetDB
        Dim strCustMst As String
        Dim strCondition As String
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If Len(txtCustCode.Text) > 0 Then
                    strCondition = "Unit_Code ='" & gstrUNITID & "' and Location_code = '" & Trim(txtLocationCode.Text) & "' and invoice_date >= '" & getDateForDB(dtpDateFrom.Value) & "'"
                    strCondition = strCondition & " and Invoice_date <= '" & getDateForDB(dtpDateTo.Value) & "' and bill_flag = 1 and Cancel_flag = 0"
                    If CheckExistanceOfFieldData((txtCustCode.Text), "Account_Code", "SalesChallan_dtl", strCondition) Then
                        Call SelectDescriptionForField("Cust_Name", "Customer_Code", "Customer_Mst", lblCustCodeDes, Trim(txtCustCode.Text))
                        txtCustPartCode.Focus()
                    Else
                        MsgBox("No Customer available To Display.", MsgBoxStyle.Information, "empower")
                        txtCustCode.Text = ""
                        txtCustCode.Focus()
                    End If
                    If Len(Trim(txtCustCode.Text)) > 0 Then
                        rsCustMst = New ClsResultSetDB
                        strCustMst = "SELECT Bill_Address1 + ', '  +  Bill_Address2 + ', ' + Bill_City + ' - ' + Bill_Pin as  invoiceAddress from Customer_Mst where Unit_Code ='" & gstrUNITID & "' and Customer_code ='" & txtCustCode.Text & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                        rsCustMst.GetResult(strCustMst)
                        If rsCustMst.GetNoRows > 0 Then
                            lblAddressDes.Text = rsCustMst.GetValue("InvoiceAddress")
                        End If
                        rsCustMst.ResultSetClose()
                        rsCustMst = Nothing
                    End If
                    EnableDisableGST()
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtCustRefRemarks_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustRefRemarks.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
            Case System.Windows.Forms.Keys.Return
                txtRemarks.Focus()
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCVDCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCVDCode.TextChanged
        On Error GoTo ErrHandler
        If Len(txtCVDCode.Text) = 0 Then
            lblCVD_Per.Text = "0.00"
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtCVDCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCVDCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If Len(txtCVDCode.Text) > 0 Then
                            Call txtCVDCode_Validating(txtCVDCode, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            txtSADCode.Focus()
                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCVDCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCVDCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If cmdCVDCodeHelp.Enabled Then Call cmdCVDCodeHelp_Click(cmdCVDCodeHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtCVDCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCVDCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        If Len(txtCVDCode.Text) > 0 Then
            If CheckExistanceOfFieldData((txtCVDCode.Text), "TxRt_Rate_No", "Gen_TaxRate", "Unit_code='" & gstrUNITID & "' and Tx_TaxeID='CVD' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))") Then
                lblCVD_Per.Text = CStr(GetTaxRate((txtCVDCode.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Unit_code= '" & gstrUNITID & "' and Tx_TaxeID='CVD') and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"))
                If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    Call SetFormulaofColumns(0, 4)
                    ToShowDatainSummeryGrid()
                End If
                If txtSADCode.Enabled Then txtSADCode.Focus()
            Else
                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtCVDCode.Text = ""
                If txtCVDCode.Enabled Then txtCVDCode.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtEcssCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEcssCode.TextChanged
        On Error GoTo ErrHandler
        If Trim(Me.txtEcssCode.Text) = "" Then
            Me.lblEcssCode.Text = "0.00"
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtEcssCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtEcssCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call cmdEcssCodeHelp_Click(cmdEcssCodeHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtEcssCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtEcssCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        Call txtEcssCode_Validating(txtEcssCode, New System.ComponentModel.CancelEventArgs(False))
                        With spdPrevInv
                            If txtSEcssCode.Enabled Then txtSEcssCode.Focus()
                        End With
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtEcssCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtEcssCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        If Trim(txtEcssCode.Text) <> "" Then
            If CheckExistanceOfFieldData(Trim(txtEcssCode.Text), "TxRt_Rate_No", "Gen_TaxRate", "Unit_code= '" & gstrUNITID & "' and Tx_TaxeID='ECS' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))") Then
                lblEcssCode.Text = CStr(GetTaxRate((Me.txtEcssCode.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Unit_code='" & gstrUNITID & "' and Tx_TaxeID='ECS' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"))
                If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    Call SetFormulaofColumns(0, 4)
                    ToShowDatainSummeryGrid()
                End If
                With Me.spdPrevInv
                    If txtSEcssCode.Enabled Then txtSEcssCode.Focus()
                End With
            Else
                MsgBox("Invalid Ecess Tax Code!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                Cancel = True
                txtEcssCode.Text = ""
                If txtEcssCode.Enabled Then txtEcssCode.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtExciseTaxType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtExciseTaxType.TextChanged
        On Error GoTo ErrHandler
        If Len(txtExciseTaxType.Text) = 0 Then
            lblExctax_Per.Text = "0.00"
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtExciseTaxType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtExciseTaxType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If Len(txtExciseTaxType.Text) > 0 Then
                            Call txtExciseTaxType_Validating(txtExciseTaxType, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            txtSaleTaxType.Focus()
                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtExciseTaxType_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtExciseTaxType.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdExciseTaxType.Enabled Then Call CmdExciseTaxType_Click(CmdExciseTaxType, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtExciseTaxType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtExciseTaxType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        If Len(txtExciseTaxType.Text) > 0 Then
            If CheckExistanceOfFieldData((txtExciseTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "Unit_code='" & gstrUNITID & "' and Tx_TaxeID='EXC' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))") Then
                lblExctax_Per.Text = CStr(GetTaxRate((txtExciseTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", "(Unit_code='" & gstrUNITID & "' and Tx_TaxeID='EXC') and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"))
                If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    Call SetFormulaofColumns(0, 4)
                    ToShowDatainSummeryGrid()
                End If
                If txtSaleTaxType.Enabled Then txtSaleTaxType.Focus()
            Else
                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtExciseTaxType.Text = ""
                If txtExciseTaxType.Enabled Then txtExciseTaxType.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtLocationCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLocationCode.TextChanged
        On Error GoTo ErrHandler
        If Len(Trim(txtLocationCode.Text)) = 0 Then
            If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                txtChallanNo.Text = ""
            End If
            txtCustCode.Text = "" : txtCustPartCode.Text = "" : txtRefNo.Text = "" : txtAmendment.Text = ""
            txtCustRefRemarks.Text = "" : txtRemarks.Text = "" : txtCVDCode.Text = "" : txtSADCode.Text = "" : txtExciseTaxType.Text = ""
            txtSaleTaxType.Text = "" : txtSurchargeTaxType.Text = ""
            '101188073
            txtCGSTType.Text = "" : txtSGSTType.Text = "" : txtUTGSTType.Text = ""
            txtIGSTType.Text = "" : txtCompCessType.Text = "" : lblHSNSAC.Text = "" : lblHSNSACCODE.Text = ""
            '101188073
            spdPrevInv.MaxRows = 0
            spdPrevInv.MaxRows = 1
            spdInvDetails.MaxRows = 0
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLocationCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtLocationCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call CmdLocCodeHelp_Click(CmdLocCodeHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLocationCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLocationCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            Select Case CmdGrpChEnt.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    txtChallanNo.Focus()
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    txtCustCode.Focus()
            End Select
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtLocationCode_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLocationCode.Leave
        On Error GoTo ErrHandler
        txtLocationCode.Text = UCase(txtLocationCode.Text)
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Len(txtLocationCode.Text) > 0 Then
                    If CheckExistanceOfFieldData((txtLocationCode.Text), "Location_Code", "SupplementaryInv_hdr", "Unit_code= '" & gstrUNITID & "'") Then
                        strDefaultLocation = Trim(txtLocationCode.Text)
                        If txtChallanNo.Enabled Then
                            txtChallanNo.Focus()
                        Else
                            txtCustCode.Focus()
                        End If
                    Else
                        Call ConfirmWindow(10411, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        txtLocationCode.Text = ""
                        txtLocationCode.Focus()
                    End If
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If Len(txtLocationCode.Text) > 0 Then
                    If CheckExistanceOfFieldData((txtLocationCode.Text), "Location_Code", "SaleConf", "UNIT_CODE = '" & gstrUNITID & "'") Then
                        strDefaultLocation = Trim(txtLocationCode.Text)
                        If txtChallanNo.Enabled Then
                            txtChallanNo.Focus()
                        Else
                            txtCustCode.Focus()
                        End If
                    Else
                        Call ConfirmWindow(10411, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        txtLocationCode.Text = ""
                        txtLocationCode.Focus()
                    End If
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function CheckExistanceOfFieldData(ByRef pstrFieldText As String, ByRef pstrColumnName As String, ByRef pstrTableName As String, Optional ByRef pstrCondition As String = "") As Boolean
        On Error GoTo ErrHandler
        CheckExistanceOfFieldData = False
        Dim strTableSql As String 'Declared To Make Select Query
        Dim rsExistData As ClsResultSetDB
        If Len(Trim(pstrCondition)) > 0 Then
            strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and " & pstrCondition
        Else
            strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "'"
        End If
        rsExistData = New ClsResultSetDB
        rsExistData.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsExistData.GetNoRows > 0 Then
            CheckExistanceOfFieldData = True
        Else
            CheckExistanceOfFieldData = False
        End If
        rsExistData.ResultSetClose()
        rsExistData = Nothing
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
            strTableSql = "select " & Trim(pstrFieldName_WhichValueRequire) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and " & pstrCondition
        Else
            strTableSql = "select " & Trim(pstrFieldName_WhichValueRequire) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "'"
        End If
        rsExistData = New ClsResultSetDB
        rsExistData.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsExistData.GetNoRows > 0 Then
            GetTaxRate = rsExistData.GetValue(Trim(pstrFieldName_WhichValueRequire))
        Else
            GetTaxRate = 0
        End If
        rsExistData.ResultSetClose()
        rsExistData = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function SetFinancialYearDates() As Object
        Dim rsBusinessPeriodDates As ClsResultSetDB
        On Error GoTo ErrHandler
        rsBusinessPeriodDates = New ClsResultSetDB
        rsBusinessPeriodDates.GetResult("Select Per_PeriodStart,Per_PeriodEnd from Gen_BusienssPeriod where Unit_code= '" & gstrUNITID & "' ")
        If rsBusinessPeriodDates.GetNoRows > 0 Then
            dtpDateFrom.Value = getDateForDB(VB6.Format(rsBusinessPeriodDates.GetValue("Per_PeriodStart"), gstrDateFormat))
            dtpDateTo.Value = getDateForDB(VB6.Format(rsBusinessPeriodDates.GetValue("Per_PeriodEnd"), gstrDateFormat))
            SetFinancialYearDates = True
        End If
        rsBusinessPeriodDates.ResultSetClose()
        rsBusinessPeriodDates = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function CheckFinancialYearDates() As Boolean
        Dim rsBusinessPeriodDates As ClsResultSetDB
        On Error GoTo ErrHandler
        CheckFinancialYearDates = False
        rsBusinessPeriodDates = New ClsResultSetDB
        rsBusinessPeriodDates.GetResult("Select Per_PeriodStart,Per_PeriodEnd from Gen_BusienssPeriod where Unit_code='" & gstrUNITID & "'")
        If rsBusinessPeriodDates.GetNoRows > 0 Then
            Financial_Start_Date = CDate(getDateForDB(VB6.Format(rsBusinessPeriodDates.GetValue("Per_PeriodStart"), gstrDateFormat)))
            Financial_End_Date = CDate(getDateForDB(VB6.Format(rsBusinessPeriodDates.GetValue("Per_PeriodEnd"), gstrDateFormat)))
            CheckFinancialYearDates = True
        Else
            MsgBox("No Business Period Defined in Gen_business Period.", MsgBoxStyle.Information, "empower")
            CheckFinancialYearDates = False
            rsBusinessPeriodDates.ResultSetClose()
            rsBusinessPeriodDates = Nothing
            Exit Function
        End If
        rsBusinessPeriodDates.ResultSetClose()
        rsBusinessPeriodDates = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function DisplayDetailsinViewMode() As Boolean
        Dim strsqlHdr As String = ""
        Dim strSqlDtl As String = ""
        Dim intMaxLoop As Short
        Dim intLoopCounter As Short
        Dim rsSupplementaryHdr As ClsResultSetDB
        Dim rsSupplementaryDtl As ClsResultSetDB
        On Error GoTo ErrHandler
        rsSupplementaryHdr = New ClsResultSetDB
        strsqlHdr = "select Location_Code,Account_Code,Cust_name,Cust_Ref,Amendment_No,Doc_No,Invoice_DateFrom,Invoice_DateTo,Item_Code,"
        strsqlHdr = strsqlHdr & "Cust_Item_Code,Currency_Code,Rate,Basic_Amount,Accessible_amount,Excise_type,CVD_type,SAD_type,Excise_per,"
        strsqlHdr = strsqlHdr & "CVD_per,SVD_per,TotalExciseAmount,CustMtrl_Amount,ToolCost_amount,SalesTax_Type,SalesTax_Per,Sales_Tax_Amount,"
        strsqlHdr = strsqlHdr & "Surcharge_salesTaxType,Surcharge_SalesTax_Per,Surcharge_Sales_Tax_Amount,Total_amount,SuppInv_Remarks,"
        strsqlHdr = strsqlHdr & "Remarks "
        strsqlHdr = strsqlHdr & ",Ecess_Type,Ecess_per,Ecess_Amount"
        strsqlHdr = strsqlHdr & ",SEcess_Type,SEcess_per,SEcess_Amount,MRP,isnull(Packing_Amount,0) as Packing_Amount"
        strsqlHdr = strsqlHdr & ",ISNULL(HSN_SAC_CODE,'') HSN_SAC_CODE,ISNULL(ISHSNORSAC,'') ISHSNORSAC"
        strsqlHdr = strsqlHdr & ",ISNULL(CGSTTXRT_TYPE,'') CGSTTXRT_TYPE,ISNULL(CGST_PERCENT,0) CGST_PERCENT"
        strsqlHdr = strsqlHdr & ",ISNULL(SGSTTXRT_TYPE,'') SGSTTXRT_TYPE,ISNULL(SGST_PERCENT,0) SGST_PERCENT"
        strsqlHdr = strsqlHdr & ",ISNULL(UTGSTTXRT_TYPE,'') UTGSTTXRT_TYPE,ISNULL(UTGST_PERCENT,0) UTGST_PERCENT"
        strsqlHdr = strsqlHdr & ",ISNULL(IGSTTXRT_TYPE,'') IGSTTXRT_TYPE,ISNULL(IGST_PERCENT,0) IGST_PERCENT"
        strsqlHdr = strsqlHdr & ",ISNULL(COMPENSATION_CESS_TYPE,'') COMPENSATION_CESS_TYPE,ISNULL(COMPENSATION_CESS_PERCENT,0) COMPENSATION_CESS_PERCENT,ISNULL(CGST_AMT,0) CGST_AMT,ISNULL(SGST_AMT,0) SGST_AMT,ISNULL(UTGST_AMT,0) UTGST_AMT,ISNULL(IGST_AMT,0) IGST_AMT,ISNULL(CCESS_AMT,0) CCESS_AMT"
        strsqlHdr = strsqlHdr & " from supplementaryINV_hdr where Unit_code='" & gstrUNITID & "' and Location_code = '" & Trim(txtLocationCode.Text) & "' and Doc_no = "
        strsqlHdr = strsqlHdr & Trim(txtChallanNo.Text)
        rsSupplementaryHdr.GetResult(strsqlHdr)
        If rsSupplementaryHdr.GetNoRows > 0 Then
            spdInvDetails.MaxRows = 1
            spdPrevInv.MaxRows = 1
            Call SetCellTypeofGrids()
            dtpDateFrom.Value = getDateForDB(VB6.Format(rsSupplementaryHdr.GetValue("Invoice_DateFrom"), gstrDateFormat))
            dtpDateTo.Value = getDateForDB(VB6.Format(rsSupplementaryHdr.GetValue("Invoice_DateTo"), gstrDateFormat))
            txtCustCode.Text = rsSupplementaryHdr.GetValue("Account_code")
            txtCustPartCode.Text = rsSupplementaryHdr.GetValue("Cust_item_code")
            lblItemCode.Text = rsSupplementaryHdr.GetValue("item_code")
            txtRefNo.Text = rsSupplementaryHdr.GetValue("Cust_ref")
            mstrCustRefNo = rsSupplementaryHdr.GetValue("Cust_ref")
            txtAmendment.Text = rsSupplementaryHdr.GetValue("Amendment_no")
            lblCurrencyDes.Text = rsSupplementaryHdr.GetValue("Currency_code")
            txtCustRefRemarks.Text = rsSupplementaryHdr.GetValue("SuppInv_Remarks")
            txtRemarks.Text = rsSupplementaryHdr.GetValue("Remarks")
            txtCVDCode.Text = rsSupplementaryHdr.GetValue("CVD_Type")
            lblCVD_Per.Text = rsSupplementaryHdr.GetValue("CVD_Per")
            txtSADCode.Text = rsSupplementaryHdr.GetValue("SAD_Type")
            lblSAD_per.Text = rsSupplementaryHdr.GetValue("SVD_Per")
            txtExciseTaxType.Text = rsSupplementaryHdr.GetValue("Excise_Type")
            lblExctax_Per.Text = rsSupplementaryHdr.GetValue("Excise_Per")
            txtEcssCode.Text = rsSupplementaryHdr.GetValue("Ecess_Type")
            lblEcssCode.Text = rsSupplementaryHdr.GetValue("Ecess_Per")
            txtSEcssCode.Text = rsSupplementaryHdr.GetValue("SEcess_Type")
            lblSEcssCode.Text = rsSupplementaryHdr.GetValue("SEcess_Per")
            txtMRP.Text = rsSupplementaryHdr.GetValue("MRP")
            txtSaleTaxType.Text = rsSupplementaryHdr.GetValue("SalesTax_type")
            lblSaltax_Per.Text = rsSupplementaryHdr.GetValue("SalesTax_Per")
            txtSurchargeTaxType.Text = rsSupplementaryHdr.GetValue("Surcharge_SalesTaxtype")
            lblSurcharge_Per.Text = rsSupplementaryHdr.GetValue("Surcharge_SalesTax_Per")
            '101188073
            If gblnGSTUnit Then
                lblHSNSAC.Text = rsSupplementaryHdr.GetValue("ISHSNORSAC")
                lblHSNSACCODE.Text = rsSupplementaryHdr.GetValue("HSN_SAC_CODE")
                txtCGSTType.Text = rsSupplementaryHdr.GetValue("CGSTTXRT_TYPE")
                lblCGSTPercent.Text = rsSupplementaryHdr.GetValue("CGST_PERCENT")
                txtSGSTType.Text = rsSupplementaryHdr.GetValue("SGSTTXRT_TYPE")
                lblSGSTPercent.Text = rsSupplementaryHdr.GetValue("SGST_PERCENT")
                txtUTGSTType.Text = rsSupplementaryHdr.GetValue("UTGSTTXRT_TYPE")
                lblUTGSTPercent.Text = rsSupplementaryHdr.GetValue("UTGST_PERCENT")
                txtIGSTType.Text = rsSupplementaryHdr.GetValue("IGSTTXRT_TYPE")
                lblIGSTPercent.Text = rsSupplementaryHdr.GetValue("IGST_PERCENT")
                txtCompCessType.Text = rsSupplementaryHdr.GetValue("COMPENSATION_CESS_TYPE")
                lblCompCessPercent.Text = rsSupplementaryHdr.GetValue("COMPENSATION_CESS_PERCENT")
            End If
            '101188073
            With spdInvDetails
                Call .SetText(enumInvoiceSummery.Rate, 1, rsSupplementaryHdr.GetValue("Rate"))
                Call .SetText(enumInvoiceSummery.BasicValue, 1, rsSupplementaryHdr.GetValue("Basic_amount"))
                Call .SetText(enumInvoiceSummery.CustSuppMat, 1, rsSupplementaryHdr.GetValue("CustMtrl_amount"))
                Call .SetText(enumInvoiceSummery.ToolCost, 1, rsSupplementaryHdr.GetValue("ToolCost_amount"))
                Call .SetText(enumInvoiceSummery.AccessableValue, 1, rsSupplementaryHdr.GetValue("Accessible_amount"))
                Call .SetText(enumInvoiceSummery.ExciseValue, 1, rsSupplementaryHdr.GetValue("TotalExciseAmount"))
                Call .SetText(enumInvoiceSummery.EcssValue, 1, rsSupplementaryHdr.GetValue("Ecess_amount"))
                Call .SetText(enumInvoiceSummery.sEcssValue, 1, rsSupplementaryHdr.GetValue("SEcess_amount"))
                Call .SetText(enumInvoiceSummery.PackingValue, 1, rsSupplementaryHdr.GetValue("Packing_Amount"))
                Call .SetText(enumInvoiceSummery.SalesTaxType, 1, rsSupplementaryHdr.GetValue("Sales_tax_Amount"))
                Call .SetText(enumInvoiceSummery.SSTType, 1, rsSupplementaryHdr.GetValue("Surcharge_Sales_tax_amount"))
                Call .SetText(enumInvoiceSummery.CGST, 1, rsSupplementaryHdr.GetValue("CGST_AMT"))
                Call .SetText(enumInvoiceSummery.SGST, 1, rsSupplementaryHdr.GetValue("SGST_AMT"))
                Call .SetText(enumInvoiceSummery.UTGST, 1, rsSupplementaryHdr.GetValue("UTGST_AMT"))
                Call .SetText(enumInvoiceSummery.IGST, 1, rsSupplementaryHdr.GetValue("IGST_AMT"))
                Call .SetText(enumInvoiceSummery.CCESS, 1, rsSupplementaryHdr.GetValue("CCESS_AMT"))
                Call .SetText(enumInvoiceSummery.SummeryInvoiceValue, 1, rsSupplementaryHdr.GetValue("Total_amount"))
                .Enabled = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = 1
                .MaxRows = .MaxRows
                .BlockMode = True
                .Lock = True
                .BlockMode = False
            End With
        Else
            DisplayDetailsinViewMode = False
        End If
        rsSupplementaryHdr.ResultSetClose()
        strSqlDtl = "Select SelectInvoice = 1,RefDoc_No,RefDoc_Date,LastSupplementary,SuppInvdate,Item_code,Cust_Item_Code,PrevRate,Rate,Rate_diff,Quantity,PrevPacking_Amount,Packing_Per,"
        strSqlDtl = strSqlDtl & "Packing_amountDiff,PrevBasic_Amount,Basic_Amount,Basic_AmountDiff,PrevAccessible_amount,Accessible_amount,"
        strSqlDtl = strSqlDtl & "Accessible_amountDiff,PrevTotalExciseAmount,CVD_Amount,SAD_amount,Excise_amount,TotalExciseAmount,TotalExciseAmountDiff,PrevCustMtrl_Amount,"
        strSqlDtl = strSqlDtl & "CustMtrl_Amount,TotalCustMtrl_Amount,CustMtrl_AmountDiff,PrevToolCost_amount,ToolCost_amount,"
        strSqlDtl = strSqlDtl & "TotalToolCost_amount,ToolCost_amountDiff,PrevSales_Tax_Amount,Sales_Tax_Amount,Sales_Tax_AmountDiff,"
        strSqlDtl = strSqlDtl & "PrevSSTAmount,SST_Amount,SST_AmountDiff,total_amount,total_amountDiff "
        strSqlDtl = strSqlDtl & ",ECESS_Amount,PrevECESS_Amount,ECESS_Amount_Diff"
        strSqlDtl = strSqlDtl & ",SECESS_Amount,PrevSECESS_Amount,SECESS_Amount_Diff"
        strSqlDtl = strSqlDtl & ",ISNULL(PREV_CGST_AMT,0) PREV_CGST_AMT ,ISNULL(NEW_CGST_AMT,0) NEW_CGST_AMT ,ISNULL(DIFF_CGST_AMT,0) DIFF_CGST_AMT"
        strSqlDtl = strSqlDtl & ",ISNULL(PREV_SGST_AMT,0) PREV_SGST_AMT ,ISNULL(NEW_SGST_AMT,0) NEW_SGST_AMT ,ISNULL(DIFF_SGST_AMT,0) DIFF_SGST_AMT"
        strSqlDtl = strSqlDtl & ",ISNULL(PREV_UTGST_AMT,0) PREV_UTGST_AMT ,ISNULL(NEW_UTGST_AMT,0) NEW_UTGST_AMT ,ISNULL(DIFF_UTGST_AMT,0) DIFF_UTGST_AMT"
        strSqlDtl = strSqlDtl & ",ISNULL(PREV_IGST_AMT,0) PREV_IGST_AMT ,ISNULL(NEW_IGST_AMT,0) NEW_IGST_AMT ,ISNULL(DIFF_IGST_AMT,0) DIFF_IGST_AMT"
        strSqlDtl = strSqlDtl & ",ISNULL(PREV_CCESS_AMT,0) PREV_CCESS_AMT ,ISNULL(NEW_CCESS_AMT,0) NEW_CCESS_AMT ,ISNULL(DIFF_CCESS_AMT,0) DIFF_CCESS_AMT"
        strSqlDtl = strSqlDtl & " from SupplementaryInv_dtl where unit_code='" & gstrUnitId & "' and Location_code ='"
        strSqlDtl = strSqlDtl & Trim(txtLocationCode.Text) & "' and Doc_no = " & Trim(txtChallanNo.Text) & vbCrLf
        strSqlDtl = strSqlDtl & "UNION Select SelectInvoice = 0,RefDoc_No,RefDoc_Date,LastSupplementary,SuppInvdate,Item_code,Cust_Item_Code,PrevRate,Rate,Rate_diff,Quantity,PrevPacking_Amount,Packing_Per,"
        strSqlDtl = strSqlDtl & "Packing_amountDiff,PrevBasic_Amount,Basic_Amount,Basic_AmountDiff,PrevAccessible_amount,Accessible_amount,"
        strSqlDtl = strSqlDtl & "Accessible_amountDiff,PrevTotalExciseAmount,CVD_Amount,SAD_amount,Excise_amount,TotalExciseAmount,TotalExciseAmountDiff,PrevCustMtrl_Amount,"
        strSqlDtl = strSqlDtl & "CustMtrl_Amount,TotalCustMtrl_Amount,CustMtrl_AmountDiff,PrevToolCost_amount,ToolCost_amount,"
        strSqlDtl = strSqlDtl & "TotalToolCost_amount,ToolCost_amountDiff,PrevSales_Tax_Amount,Sales_Tax_Amount,Sales_Tax_AmountDiff,"
        strSqlDtl = strSqlDtl & "PrevSSTAmount,SST_Amount,SST_AmountDiff,Total_amount,total_amountDiff "
        strSqlDtl = strSqlDtl & ",ECESS_Amount,PrevECESS_Amount,ECESS_Amount_Diff"
        strSqlDtl = strSqlDtl & ",SECESS_Amount,PrevSECESS_Amount,SECESS_Amount_Diff"
        strSqlDtl = strSqlDtl & ",ISNULL(PREV_CGST_AMT,0) PREV_CGST_AMT ,ISNULL(NEW_CGST_AMT,0) NEW_CGST_AMT ,ISNULL(DIFF_CGST_AMT,0) DIFF_CGST_AMT"
        strSqlDtl = strSqlDtl & ",ISNULL(PREV_SGST_AMT,0) PREV_SGST_AMT ,ISNULL(NEW_SGST_AMT,0) NEW_SGST_AMT ,ISNULL(DIFF_SGST_AMT,0) DIFF_SGST_AMT"
        strSqlDtl = strSqlDtl & ",ISNULL(PREV_UTGST_AMT,0) PREV_UTGST_AMT ,ISNULL(NEW_UTGST_AMT,0) NEW_UTGST_AMT ,ISNULL(DIFF_UTGST_AMT,0) DIFF_UTGST_AMT"
        strSqlDtl = strSqlDtl & ",ISNULL(PREV_IGST_AMT,0) PREV_IGST_AMT ,ISNULL(NEW_IGST_AMT,0) NEW_IGST_AMT ,ISNULL(DIFF_IGST_AMT,0) DIFF_IGST_AMT"
        strSqlDtl = strSqlDtl & ",ISNULL(PREV_CCESS_AMT,0) PREV_CCESS_AMT ,ISNULL(NEW_CCESS_AMT,0) NEW_CCESS_AMT ,ISNULL(DIFF_CCESS_AMT,0) DIFF_CCESS_AMT"
        strSqlDtl = strSqlDtl & " from SuppCreditAdvise_Dtl where Unit_code='" & gstrUnitId & "' and Location_code ='"
        strSqlDtl = strSqlDtl & Trim(txtLocationCode.Text) & "' and Doc_no = " & Trim(txtChallanNo.Text)
        strSqlDtl = strSqlDtl & " order by SelectInvoice DESC "
        rsSupplementaryDtl = New ClsResultSetDB
        rsSupplementaryDtl.GetResult(strSqlDtl)
        GetInvoiceNumbers(rsSupplementaryDtl)
        cmdSelectInvoice.Enabled = True
        If rsSupplementaryDtl.GetNoRows > 0 Then
            intMaxLoop = rsSupplementaryDtl.GetNoRows
            rsSupplementaryDtl.MoveFirst()
            With spdPrevInv
                spdPrevInv.MaxRows = 0
                spdPrevInv.MaxRows = 1
                Call SetCellTypeofGrids()
                Call SetMaxLengthofGrid(4)
                Call SetWidthofColumnsinGrid()
                For intLoopCounter = 1 To 1
                    If rsSupplementaryDtl.GetValue("SelectInvoice") = 1 Then
                        .Col = enumPreInvoiceDetails.Select_Invoice
                        .Row = intLoopCounter
                        .Value = System.Windows.Forms.CheckState.Checked
                    Else
                        .Col = enumPreInvoiceDetails.Select_Invoice
                        .Row = intLoopCounter
                        .Value = System.Windows.Forms.CheckState.Unchecked
                    End If
                    Call .SetText(enumPreInvoiceDetails.Invoice_No, intLoopCounter, rsSupplementaryDtl.GetValue("RefDoc_no"))
                    Call .SetText(enumPreInvoiceDetails.Invoice_Date, intLoopCounter, VB6.Format(rsSupplementaryDtl.GetValue("RefDoc_Date"), gstrDateFormat))
                    Call .SetText(enumPreInvoiceDetails.LastSupplementary, intLoopCounter, rsSupplementaryDtl.GetValue("LastSupplementary"))
                    Call .SetText(enumPreInvoiceDetails.SupplementaryDate, intLoopCounter, VB6.Format(rsSupplementaryDtl.GetValue("SuppInvdate"), gstrDateFormat))
                    Call .SetText(enumPreInvoiceDetails.Quantity, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Quantity")))
                    Call .SetText(enumPreInvoiceDetails.Rate, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevRate")))
                    Call .SetText(enumPreInvoiceDetails.New_Rate, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Rate")))
                    Call .SetText(enumPreInvoiceDetails.Rate_diff, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Rate_diff")))
                    Call .SetText(enumPreInvoiceDetails.TotalPacking, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevPacking_amount")))
                    Call .SetText(enumPreInvoiceDetails.NewPacking, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Packing_Per")))
                    Call .SetText(enumPreInvoiceDetails.NewTotalPacking, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Packing_amountdiff")))
                    Call .SetText(enumPreInvoiceDetails.Basic, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevBasic_amount")))
                    Call .SetText(enumPreInvoiceDetails.NewBasic, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Basic_amount")))
                    Call .SetText(enumPreInvoiceDetails.BasicDiff, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Basic_amountDiff")))
                    Call .SetText(enumPreInvoiceDetails.TotalCustSuppMaterial, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevCustMtrl_amount")))
                    Call .SetText(enumPreInvoiceDetails.NewCustSuppMaterial, intLoopCounter, Val(rsSupplementaryDtl.GetValue("CustMtrl_amount")))
                    Call .SetText(enumPreInvoiceDetails.NewTotalCustSuppMaterial, intLoopCounter, Val(rsSupplementaryDtl.GetValue("TotalCustMtrl_amount")))
                    Call .SetText(enumPreInvoiceDetails.CustSuppMaterial_diff, intLoopCounter, Val(rsSupplementaryDtl.GetValue("CustMtrl_amountDiff")))
                    Call .SetText(enumPreInvoiceDetails.ToolCost, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevToolCost_amount")))
                    Call .SetText(enumPreInvoiceDetails.newToolCost, intLoopCounter, Val(rsSupplementaryDtl.GetValue("ToolCost_amount")))
                    Call .SetText(enumPreInvoiceDetails.NewTotalToolCost, intLoopCounter, Val(rsSupplementaryDtl.GetValue("TotalToolCost_amount")))
                    Call .SetText(enumPreInvoiceDetails.ToolCost_diff, intLoopCounter, Val(rsSupplementaryDtl.GetValue("ToolCost_amountDiff")))
                    Call .SetText(enumPreInvoiceDetails.AccessableValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevAccessible_amount")))
                    Call .SetText(enumPreInvoiceDetails.NewAccessableValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Accessible_amount")))
                    Call .SetText(enumPreInvoiceDetails.AccessableValue_Diff, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Accessible_amountDiff")))
                    Call .SetText(enumPreInvoiceDetails.TotalExciseValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevTotalExciseAmount")))
                    Call .SetText(enumPreInvoiceDetails.NewCVDValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("CVD_Amount")))
                    Call .SetText(enumPreInvoiceDetails.NewSADValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("SAD_Amount")))
                    Call .SetText(enumPreInvoiceDetails.NewExciseValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Excise_amount")))
                    Call .SetText(enumPreInvoiceDetails.NewTotalExciseValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("TotalExciseamount")))
                    Call .SetText(enumPreInvoiceDetails.TotalExciseValueDiff, intLoopCounter, Val(rsSupplementaryDtl.GetValue("TotalExciseamountDiff")))
                    Call .SetText(enumPreInvoiceDetails.TotalEcessValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevECESS_Amount")))
                    Call .SetText(enumPreInvoiceDetails.NewEcessValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("ECESS_Amount")))
                    Call .SetText(enumPreInvoiceDetails.TotalEcssDiff, intLoopCounter, Val(rsSupplementaryDtl.GetValue("ECESS_Amount_Diff")))
                    Call .SetText(enumPreInvoiceDetails.TotalsEcessValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevSECESS_Amount")))
                    Call .SetText(enumPreInvoiceDetails.NewsEcessValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("SECESS_Amount")))
                    Call .SetText(enumPreInvoiceDetails.TotalsEcssDiff, intLoopCounter, Val(rsSupplementaryDtl.GetValue("SECESS_Amount_Diff")))
                    Call .SetText(enumPreInvoiceDetails.SalesTaxValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevSales_Tax_amount")))
                    Call .SetText(enumPreInvoiceDetails.NewSalesTaxValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Sales_Tax_amount")))
                    Call .SetText(enumPreInvoiceDetails.SalesTaxValueDiff, intLoopCounter, Val(rsSupplementaryDtl.GetValue("Sales_Tax_amountDiff")))
                    Call .SetText(enumPreInvoiceDetails.SSTVlaue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PrevSSTAmount")))
                    Call .SetText(enumPreInvoiceDetails.NewSSTValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("SST_Amount")))
                    Call .SetText(enumPreInvoiceDetails.SSTVlaueDiff, intLoopCounter, Val(rsSupplementaryDtl.GetValue("SST_AmountDiff")))
                    '101188073
                    If gblnGSTUnit Then
                        Call .SetText(enumPreInvoiceDetails.CGST_AMT, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PREV_CGST_AMT")))
                        Call .SetText(enumPreInvoiceDetails.NEW_CGST_AMT, intLoopCounter, Val(rsSupplementaryDtl.GetValue("NEW_CGST_AMT")))
                        Call .SetText(enumPreInvoiceDetails.DIFF_CGST_AMT, intLoopCounter, Val(rsSupplementaryDtl.GetValue("DIFF_CGST_AMT")))
                        Call .SetText(enumPreInvoiceDetails.SGST_AMT, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PREV_SGST_AMT")))
                        Call .SetText(enumPreInvoiceDetails.NEW_SGST_AMT, intLoopCounter, Val(rsSupplementaryDtl.GetValue("NEW_SGST_AMT")))
                        Call .SetText(enumPreInvoiceDetails.DIFF_SGST_AMT, intLoopCounter, Val(rsSupplementaryDtl.GetValue("DIFF_SGST_AMT")))
                        Call .SetText(enumPreInvoiceDetails.UTGST_AMT, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PREV_UTGST_AMT")))
                        Call .SetText(enumPreInvoiceDetails.NEW_UTGST_AMT, intLoopCounter, Val(rsSupplementaryDtl.GetValue("NEW_UTGST_AMT")))
                        Call .SetText(enumPreInvoiceDetails.DIFF_UTGST_AMT, intLoopCounter, Val(rsSupplementaryDtl.GetValue("DIFF_UTGST_AMT")))
                        Call .SetText(enumPreInvoiceDetails.IGST_AMT, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PREV_IGST_AMT")))
                        Call .SetText(enumPreInvoiceDetails.NEW_IGST_AMT, intLoopCounter, Val(rsSupplementaryDtl.GetValue("NEW_IGST_AMT")))
                        Call .SetText(enumPreInvoiceDetails.DIFF_IGST_AMT, intLoopCounter, Val(rsSupplementaryDtl.GetValue("DIFF_IGST_AMT")))
                        Call .SetText(enumPreInvoiceDetails.CCESS_AMT, intLoopCounter, Val(rsSupplementaryDtl.GetValue("PREV_CCESS_AMT")))
                        Call .SetText(enumPreInvoiceDetails.NEW_CCESS_AMT, intLoopCounter, Val(rsSupplementaryDtl.GetValue("NEW_CCESS_AMT")))
                        Call .SetText(enumPreInvoiceDetails.DIFF_CCESS_AMT, intLoopCounter, Val(rsSupplementaryDtl.GetValue("DIFF_CCESS_AMT")))
                    End If
                    '101188073
                    Call .SetText(enumPreInvoiceDetails.TotalCurrInvValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("total_amount")))
                    Call .SetText(enumPreInvoiceDetails.TotalInvoiceValue, intLoopCounter, Val(rsSupplementaryDtl.GetValue("total_amountDiff")))
                    If rsSupplementaryDtl.GetValue("SelectInvoice") = 1 Then
                        Call .SetText(enumPreInvoiceDetails.flag, intLoopCounter, "S")
                    Else
                        Call .SetText(enumPreInvoiceDetails.flag, intLoopCounter, "")
                    End If
                    rsSupplementaryDtl.MoveNext()
                Next
                .Enabled = True
                .Col = 1
                .Col2 = .MaxCols
                .Row = 1
                .MaxRows = .MaxRows
                .BlockMode = True
                .Lock = True
                .BlockMode = False
            End With
        End If
        rsSupplementaryDtl.ResultSetClose()
        DisplayDetailsinViewMode = True
        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
        chkShowDetailAnexture.Enabled = True
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function DisplayDetailsInAddMode() As Boolean
        Dim strSalesDetailData As String
        Dim intMaxRow As Short
        Dim intLoopCounter As Short
        Dim rsLastSupplementary As ClsResultSetDB
        Dim strSupplementary As String
        Dim strRefDocNo As String
        Dim strCustItemCode As String
        Dim strItemCode As String
        On Error GoTo ErrHandler
        Call SetWidthofColumnsinGrid()
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Arrow)
        spdPrevInv.SetRefStyle(2)
        If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            Call SetFormulaofColumns(0, 4)
            bCheck = False
            ToShowDatainSummeryGrid()
        End If
        strSupplementary = "Select Count(*) from Sales_Dtl Where Unit_code='" & gstrUNITID & "' and SupplementaryInvoiceFlag = 1 And Location_Code = '" & txtLocationCode.Text & "' And Cust_Item_Code = '" & txtCustPartCode.Text & "' And Item_Code = '" & lblItemCode.Text & "'"
        rsLastSupplementary = New ClsResultSetDB
        rsLastSupplementary.GetResult(strSupplementary)
        If Val(rsLastSupplementary.GetValueByNo(0)) > 0 Then
            rsLastSupplementary.ResultSetClose()
            rsLastSupplementary = New ClsResultSetDB
            strSupplementary = "Select Distinct LastSupplementary, rate, packing_per, custmtrl_amount, toolCost_amount from SupplementaryInv_Dtl Where Unit_code='" & gstrUNITID & "' and Item_Code = '" & lblItemCode.Text & "' And Cust_Item_Code = '" & txtCustPartCode.Text & "' And Location_Code = '" & txtLocationCode.Text & "'"
            rsLastSupplementary.GetResult(strSupplementary)
            rsLastSupplementary.ResultSetClose()
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function DisplayDetailsInAddMode_Old() As Boolean
        Dim strSalesDetailData As String
        Dim rsSalesDtl As ClsResultSetDB
        Dim intMaxRow As Short
        Dim intLoopCounter As Short
        Dim rsLastSupplementary As ClsResultSetDB
        Dim strSupplementary As String
        Dim strRefDocNo As String
        Dim strCustItemCode As String
        Dim strItemCode As String
        On Error GoTo ErrHandler
        rsSalesDtl = New ClsResultSetDB
        strSalesDetailData = "select * from SupplementaryData('" & gstrUNITID & "','" & Format(dtpDateFrom.Value, "dd MMM yyyy") & "','"
        strSalesDetailData = strSalesDetailData & Format(dtpDateTo.Value, "dd MMM yyyy") & "','" & txtCustCode.Text & "','" & lblItemCode.Text
        strSalesDetailData = strSalesDetailData & "','" & txtCustPartCode.Text & "')"
        rsSalesDtl.GetResult(strSalesDetailData)
        intMaxRow = rsSalesDtl.GetNoRows
        rsSalesDtl.MoveFirst()
        If intMaxRow > 0 Then
            spdPrevInv.MaxRows = 1
            spdPrevInv.MaxRows = intMaxRow
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            Call SetCellTypeofGrids()
            Call SetMaxLengthofGrid(4)
            With spdPrevInv
                rsLastSupplementary = New ClsResultSetDB
                For intLoopCounter = 1 To intMaxRow
                    Call .SetText(enumPreInvoiceDetails.Invoice_No, intLoopCounter, rsSalesDtl.GetValue("Doc_No"))
                    strRefDocNo = rsSalesDtl.GetValue("Doc_No")
                    .SetText(enumPreInvoiceDetails.Invoice_Date, intLoopCounter, VB6.Format(rsSalesDtl.GetValue("Invoice_date"), gstrDateFormat))
                    .SetText(enumPreInvoiceDetails.Quantity, intLoopCounter, rsSalesDtl.GetValue("Sales_Quantity"))
                    .SetText(enumPreInvoiceDetails.Rate, intLoopCounter, rsSalesDtl.GetValue("Rate"))
                    .SetText(enumPreInvoiceDetails.New_Rate, intLoopCounter, 0)
                    .SetText(enumPreInvoiceDetails.TotalCustSuppMaterial, intLoopCounter, rsSalesDtl.GetValue("CustMtrl_amount"))
                    .SetText(enumPreInvoiceDetails.ToolCost, intLoopCounter, rsSalesDtl.GetValue("ToolCost_amount"))
                    .SetText(enumPreInvoiceDetails.TotalPacking, intLoopCounter, rsSalesDtl.GetValue("Packing"))
                    'Excise Rate
                    .SetText(enumPreInvoiceDetails.TotalExciseValue, intLoopCounter, rsSalesDtl.GetValue("Excise_amount"))
                    .SetText(enumPreInvoiceDetails.TotalEcessValue, intLoopCounter, rsSalesDtl.GetValue("Ecess_Amount"))
                    .SetText(enumPreInvoiceDetails.SalesTaxValue, intLoopCounter, rsSalesDtl.GetValue("SalesTax_Amount"))
                    'SalesTax Value
                    .SetText(enumPreInvoiceDetails.SSTVlaue, intLoopCounter, rsSalesDtl.GetValue("SSalesTax_Amount"))
                    'SSTType
                    .SetText(enumPreInvoiceDetails.Basic, intLoopCounter, rsSalesDtl.GetValue("Basic_amount"))
                    .SetText(enumPreInvoiceDetails.AccessableValue, intLoopCounter, rsSalesDtl.GetValue("Accessible_amount"))
                    If rsSalesDtl.GetValue("SupplemnetryInv") = True Then
                        strSupplementary = "select a.Doc_No,b.Invoice_date,a.RefDoc_No,a.Item_code,a.Cust_Item_Code from"
                        strSupplementary = strSupplementary & " SupplementaryInv_dtl a,SupplementaryInv_hdr b Where "
                        strSupplementary = strSupplementary & " a.Unit_code=b.Unit_code and a.Unit_code='" & gstrUNITID & "' and a.Doc_No = b.Doc_No And a.Location_Code = b.Location_Code"
                        strSupplementary = strSupplementary & " and a.RefDoc_no = '" & strRefDocNo & "' and a.Item_code = '"
                        strSupplementary = strSupplementary & Trim(lblItemCode.Text) & "' and a.Cust_Item_code = '"
                        strSupplementary = strSupplementary & Trim(txtCustPartCode.Text) & "' and b.Invoice_date < getdate()"
                        strSupplementary = strSupplementary & " order by Invoice_date,a.Doc_no "
                        rsLastSupplementary.GetResult(strSupplementary)
                        If rsLastSupplementary.GetNoRows > 0 Then
                            rsLastSupplementary.MoveLast()
                            .SetText(enumPreInvoiceDetails.LastSupplementary, intLoopCounter, rsLastSupplementary.GetValue("Doc_no"))
                            .SetText(enumPreInvoiceDetails.SupplementaryDate, intLoopCounter, VB6.Format(rsLastSupplementary.GetValue("Invoice_date"), gstrDateFormat))
                        End If
                        rsLastSupplementary.ResultSetClose()
                        .Row = intLoopCounter
                        .Row2 = intLoopCounter
                        .Col = 1
                        .Col2 = .MaxCols
                        .BlockMode = True : .ForeColor = System.Drawing.Color.Blue
                        .BlockMode = False
                    Else
                        .Row = intLoopCounter
                        .Row2 = intLoopCounter
                        .Col = 1
                        .Col2 = .MaxCols
                        .BlockMode = True : .ForeColor = System.Drawing.SystemColors.WindowText
                        .BlockMode = False
                    End If
                    rsSalesDtl.MoveNext()
                Next
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Arrow)
                .SetRefStyle(2)
                If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    Call SetFormulaofColumns(0, 4)
                    ToShowDatainSummeryGrid()
                End If
            End With
        Else
            MsgBox("No Invoices Available For Selected Condition.")
        End If
        rsSalesDtl.ResultSetClose()
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub txtMRP_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMRP.TextChanged
        On Error GoTo ErrHandler
        If Val(Trim(txtMRP.Text)) = 0 Then
            Call ChangeInvTypeCaption("")
        Else
            Call ChangeInvTypeCaption("M")
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtMRP_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtMRP.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.Return And Shift = 0 Then
            With spdPrevInv
                .Row = 1
                .Col = 1
                .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
            End With
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtMRP_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMRP.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If (KeyAscii < System.Windows.Forms.Keys.D0 Or KeyAscii > System.Windows.Forms.Keys.D9) And KeyAscii <> System.Windows.Forms.Keys.Back And KeyAscii <> System.Windows.Forms.Keys.Delete And KeyAscii <> System.Windows.Forms.Keys.Return And KeyAscii <> System.Windows.Forms.Keys.Tab Then
            KeyAscii = 0
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtRefNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRefNo.TextChanged
        On Error GoTo ErrHandler
        If Len(Trim(txtRefNo.Text)) = 0 Then
            txtAmendment.Text = ""
            txtMRP.Text = "0.0000"
            Call ChangeInvTypeCaption("")
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtRefNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtRefNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call CmdRefNoHelp_Click(CmdRefNoHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtRefNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRefNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
            Case System.Windows.Forms.Keys.Return
                Select Case CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        txtAmendment.Focus()
                End Select
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtRefNo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRefNo.Leave
        Dim rsCustref As ClsResultSetDB
        Dim intRecordCount As Short
        On Error GoTo ErrHandler
        Dim strRefSql As String
        If Len(Trim(txtRefNo.Text)) = 0 Then Exit Sub
        strRefSql = "Select b.Cust_Ref,b.Amendment_No,a.po_type from Cust_Ord_hdr a,Cust_Ord_Dtl b"
        strRefSql = strRefSql & " where a.Unit_code = b.Unit_code and a.Unit_code='" & gstrUNITID & "' and b.Account_Code='" & Trim(txtCustCode.Text) & "' and b.Active_flag ='A' and "
        strRefSql = strRefSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and PO_type <> 'E'and "
        strRefSql = strRefSql & " a.Amendment_No = b.amendment_No  AND a.Authorized_Flag = 1 and b.Cust_drgNo = '"
        strRefSql = strRefSql & Trim(txtCustPartCode.Text) & "' and ITem_code = '" & Trim(lblItemCode.Text) & "' and "
        strRefSql = strRefSql & " a.Valid_date >'" & Format(GetServerDate, "dd MMM yyyy") & "' and effect_Date <= '"
        strRefSql = strRefSql & Format(GetServerDate, "dd MMM yyyy") & "' and b.Cust_ref = '" & Trim(txtRefNo.Text) & "' order by b.Cust_Ref,b.Amendment_No"
        rsCustref = New ClsResultSetDB
        rsCustref.GetResult(strRefSql)
        intRecordCount = rsCustref.GetNoRows
        If intRecordCount = 0 Then
            MsgBox("No Refrence available to Display", MsgBoxStyle.Information) : txtRefNo.Text = "" : txtRefNo.Focus() : Exit Sub
        ElseIf intRecordCount > 1 Then
            MsgBox("There are more then One Return for entered Customer SO.", MsgBoxStyle.Information) : txtRefNo.Text = "" : txtRefNo.Focus() : Exit Sub
        Else
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            txtRefNo.Text = rsCustref.GetValue("Cust_ref")
            txtAmendment.Text = rsCustref.GetValue("Amendment_no")
            mstrSOType.Value = Trim(rsCustref.GetValue("po_type"))
            rsCustref.ResultSetClose()
            Call ChangeInvTypeCaption(mstrSOType.Value)
            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                Call DisplayNewSOHdrDetails()
            End If
            If Len(mstrCustRefNo) > 0 And CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                If (mstrCustRefNo) <> Trim(txtRefNo.Text) Then
                    Call DisplayNewSOHdrDetails()
                End If
            End If
            If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                Call SetFormulaofColumns(0, 4)
                ToShowDatainSummeryGrid()
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Arrow)
            txtCustRefRemarks.Focus()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub DisplayNewSOHdrDetails()
        Dim rsCustOrdHdr As ClsResultSetDB
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim intLoopCounter As Short
        Dim intMaxLoopCount As Short
        Dim strSQL As String
        On Error GoTo ErrHandler
        rsCustOrdHdr = New ClsResultSetDB
        rsCustOrdDtl = New ClsResultSetDB
        strSQL = "Select Currency_code,SalesTax_Type,Surcharge_code from Cust_ord_hdr where Unit_code='" & gstrUNITID & "' and active_flag= 'A' and Authorized_flag  = 1 and Account_code = '" & Trim(txtCustCode.Text)
        strSQL = strSQL & "' and Cust_ref = '" & Trim(txtRefNo.Text) & "' and amendment_no = '" & Trim(txtAmendment.Text) & "'"
        rsCustOrdHdr.GetResult(strSQL)
        If rsCustOrdHdr.GetNoRows > 0 Then
            lblCurrencyDes.Text = rsCustOrdHdr.GetValue("Currency_code")
            If Not gblnGSTUnit Then '101188073
                txtSaleTaxType.Text = rsCustOrdHdr.GetValue("SalesTax_Type")
                txtSurchargeTaxType.Text = rsCustOrdHdr.GetValue("Surcharge_code")
            End If
            strSQL = "select Rate,Cust_Mtrl,Packing,Excise_Duty,Tool_Cost,isnull(AccessibleRateForMRP,0) as AccessibleRateForMRP,ISNULL(CGSTTXRT_TYPE,'') CGSTTXRT_TYPE,ISNULL(SGSTTXRT_TYPE,'') SGSTTXRT_TYPE,ISNULL(UTGSTTXRT_TYPE,'') UTGSTTXRT_TYPE,ISNULL(IGSTTXRT_TYPE,'') IGSTTXRT_TYPE,ISNULL(COMPENSATION_CESS,'') COMPENSATION_CESS from Cust_ord_dtl where Unit_code='" & gstrUnitId & "' and active_flag= 'A' and Authorized_flag  = 1 "
            strSQL = strSQL & " and Account_code = '" & Trim(txtCustCode.Text) & "' and Cust_ref = '" & Trim(txtRefNo.Text)
            strSQL = strSQL & "' and amendment_no = '" & Trim(txtAmendment.Text) & "' and cust_drgNo = '" & Trim(txtCustPartCode.Text) & "' and "
            strSQL = strSQL & " Item_code = '" & Trim(lblItemCode.Text) & "'"
            rsCustOrdDtl.GetResult(strSQL)
            If rsCustOrdDtl.GetNoRows > 0 Then
                If Not gblnGSTUnit Then '101188073
                    txtExciseTaxType.Text = rsCustOrdDtl.GetValue("Excise_duty")
                Else
                    txtCGSTType.Text = rsCustOrdDtl.GetValue("CGSTTXRT_TYPE")
                    txtSGSTType.Text = rsCustOrdDtl.GetValue("SGSTTXRT_TYPE")
                    txtUTGSTType.Text = rsCustOrdDtl.GetValue("UTGSTTXRT_TYPE")
                    txtIGSTType.Text = rsCustOrdDtl.GetValue("IGSTTXRT_TYPE")
                    txtCompCessType.Text = rsCustOrdDtl.GetValue("COMPENSATION_CESS")
                End If
                intMaxLoopCount = 1
                With spdPrevInv
                    For intLoopCounter = 1 To intMaxLoopCount
                        Call .SetText(enumPreInvoiceDetails.New_Rate, intLoopCounter, rsCustOrdDtl.GetValue("Rate"))
                        Call .SetText(enumPreInvoiceDetails.NewCustSuppMaterial, intLoopCounter, rsCustOrdDtl.GetValue("Cust_mtrl"))
                        Call .SetText(enumPreInvoiceDetails.NewPacking, intLoopCounter, rsCustOrdDtl.GetValue("Packing"))
                        Call .SetText(enumPreInvoiceDetails.newToolCost, intLoopCounter, rsCustOrdDtl.GetValue("Tool_cost"))
                        txtMRP.Text = rsCustOrdDtl.GetValue("AccessibleRateForMRP")
                    Next
                End With
                If Not gblnGSTUnit Then '101188073
                    lblExctax_Per.Text = CStr(GetTaxRate((txtExciseTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Unit_code='" & gstrUnitId & "' and Tx_TaxeID='EXC') and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"))
                    lblSaltax_Per.Text = CStr(GetTaxRate((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Unit_code= '" & gstrUnitId & "' and (Tx_TaxeID='CST' OR Tx_TaxeID='LST' OR Tx_TaxeID='VAT') and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"))
                    lblSurcharge_Per.Text = CStr(GetTaxRate((txtSurchargeTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Unit_code= '" & gstrUnitId & "' and Tx_TaxeID='SST' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"))
                Else
                    lblCGSTPercent.Text = CStr(GetTaxRate((txtCGSTType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Unit_code='" & gstrUnitId & "' and Tx_TaxeID='CGST') and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"))
                    lblSGSTPercent.Text = CStr(GetTaxRate((txtSGSTType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Unit_code='" & gstrUnitId & "' and Tx_TaxeID='SGST') and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"))
                    lblUTGSTPercent.Text = CStr(GetTaxRate((txtUTGSTType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Unit_code='" & gstrUnitId & "' and Tx_TaxeID='UTGST') and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"))
                    lblIGSTPercent.Text = CStr(GetTaxRate((txtIGSTType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Unit_code='" & gstrUnitId & "' and Tx_TaxeID='IGST') and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"))
                    lblCompCessPercent.Text = CStr(GetTaxRate((txtCompCessType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Unit_code='" & gstrUnitId & "' and Tx_TaxeID='GSTEC') and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"))
                End If
            End If
            rsCustOrdDtl.ResultSetClose()
        End If
        rsCustOrdHdr.ResultSetClose()
        rsCustOrdHdr = Nothing
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Sub DisplayNewSOHdrDetails_Old()
        Dim rsCustOrdHdr As ClsResultSetDB
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim intLoopCounter As Short
        Dim intMaxLoopCount As Short
        Dim strSQL As String
        On Error GoTo ErrHandler
        rsCustOrdHdr = New ClsResultSetDB
        rsCustOrdDtl = New ClsResultSetDB
        strSQL = "Select Currency_code,SalesTax_Type,Surcharge_code from Cust_ord_hdr where Unit_code='" & gstrUNITID & "' and active_flag= 'A' and Authorized_flag  = 1 and Account_code = '" & Trim(txtCustCode.Text)
        strSQL = strSQL & "' and Cust_ref = '" & Trim(txtRefNo.Text) & "' and amendment_no = '" & Trim(txtAmendment.Text) & "'"
        rsCustOrdHdr.GetResult(strSQL)
        If rsCustOrdHdr.GetNoRows > 0 Then
            lblCurrencyDes.Text = rsCustOrdHdr.GetValue("Currency_code")
            txtSaleTaxType.Text = rsCustOrdHdr.GetValue("SalesTax_Type")
            txtSurchargeTaxType.Text = rsCustOrdHdr.GetValue("Surcharge_code")
            strSQL = "select Rate,Cust_Mtrl,Packing,Excise_Duty,Tool_Cost from Cust_ord_dtl where Unit_code='" & gstrUNITID & "' and active_flag= 'A' and Authorized_flag  = 1 "
            strSQL = strSQL & " and Account_code = '" & Trim(txtCustCode.Text) & "' and Cust_ref = '" & Trim(txtRefNo.Text)
            strSQL = strSQL & "' and amendment_no = '" & Trim(txtAmendment.Text) & "' and cust_drgNo = '" & Trim(txtCustPartCode.Text) & "' and "
            strSQL = strSQL & " Item_code = '" & Trim(lblItemCode.Text) & "'"
            rsCustOrdDtl.GetResult(strSQL)
            If rsCustOrdDtl.GetNoRows > 0 Then
                txtExciseTaxType.Text = rsCustOrdDtl.GetValue("Excise_duty")
                intMaxLoopCount = spdPrevInv.MaxRows
                With spdPrevInv
                    For intLoopCounter = 1 To intMaxLoopCount
                        Call .SetText(enumPreInvoiceDetails.New_Rate, intLoopCounter, rsCustOrdDtl.GetValue("Rate"))
                        Call .SetText(enumPreInvoiceDetails.NewCustSuppMaterial, intLoopCounter, rsCustOrdDtl.GetValue("Cust_mtrl"))
                        Call .SetText(enumPreInvoiceDetails.NewPacking, intLoopCounter, rsCustOrdDtl.GetValue("Packing"))
                        Call .SetText(enumPreInvoiceDetails.newToolCost, intLoopCounter, rsCustOrdDtl.GetValue("Tool_cost"))
                    Next
                End With
                lblExctax_Per.Text = CStr(GetTaxRate((txtExciseTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Unit_code= '" & gstrUNITID & "' and Tx_TaxeID='EXC' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"))
                lblSaltax_Per.Text = CStr(GetTaxRate((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Unit_code= '" & gstrUNITID & "' and (Tx_TaxeID='CST' OR Tx_TaxeID='LST' OR Tx_TaxeID='VAT') and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"))
                lblSurcharge_Per.Text = CStr(GetTaxRate((txtSurchargeTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Unit_code= '" & gstrUNITID & "' and Tx_TaxeID='SST' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"))
            End If
            rsCustOrdDtl.ResultSetClose()
        End If
        rsCustOrdHdr.ResultSetClose()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtRemarks_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRemarks.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
            Case System.Windows.Forms.Keys.Return
                Me.txtExciseTaxType.Focus()
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Public Sub AddHeadersOfGrids()
        On Error GoTo ErrHandler
        With spdPrevInv
            '101188073
            .MaxCols = enumPreInvoiceDetails.flag
            '101188073
            .Row = 0
            .Col = enumPreInvoiceDetails.Select_Invoice : .Text = "Select Invoice"
            .Row = 0
            .Col = enumPreInvoiceDetails.Invoice_No : .Text = "Invoice No"
            .Row = 0
            .Col = enumPreInvoiceDetails.Invoice_Date : .Text = "Invoice Date"
            .Row = 0
            .Col = enumPreInvoiceDetails.LastSupplementary : .Text = "Supp Inv No"
            .Row = 0
            .Col = enumPreInvoiceDetails.SupplementaryDate : .Text = "Supp Inv Date"
            .Row = 0
            .Col = enumPreInvoiceDetails.Quantity : .Text = "Quantity"
            .Row = 0
            .Col = enumPreInvoiceDetails.Rate : .Text = "Rate"
            .Row = 0
            .Col = enumPreInvoiceDetails.New_Rate : .Text = "New Rate"
            .Row = 0
            .Col = enumPreInvoiceDetails.Rate_diff : .Text = "Rate diff"
            .Row = 0
            .Col = enumPreInvoiceDetails.TotalCustSuppMaterial : .Text = "Cust Supp Material Amount"
            .Row = 0
            .Col = enumPreInvoiceDetails.NewCustSuppMaterial : .Text = "New Cust Supp Material (PerValue)"
            .Row = 0
            .Col = enumPreInvoiceDetails.NewTotalCustSuppMaterial : .Text = "New Total Cust Supp Material"
            .Row = 0
            .Col = enumPreInvoiceDetails.CustSuppMaterial_diff : .Text = "Cust Supp Mat.diff"
            .Row = 0
            .Col = enumPreInvoiceDetails.ToolCost : .Text = "Tool Cost Amount"
            .Row = 0
            .Col = enumPreInvoiceDetails.newToolCost : .Text = "New Tool Cost (Per Value)"
            .Row = 0
            .Col = enumPreInvoiceDetails.NewTotalToolCost : .Text = "New Total Tool Cost"
            .Row = 0
            .Col = enumPreInvoiceDetails.ToolCost_diff : .Text = "Tool Cost diff"
            .Row = 0
            .Col = enumPreInvoiceDetails.TotalPacking : .Text = "Packing"
            .Row = 0
            .Col = enumPreInvoiceDetails.NewPacking : .Text = "New Packing (Per Unit)"
            .Row = 0
            .Col = enumPreInvoiceDetails.NewTotalPacking : .Text = "New Total Packing"
            .Row = 0
            .Col = enumPreInvoiceDetails.NewCVDValue : .Text = "New CVD Value"
            .Row = 0
            .Col = enumPreInvoiceDetails.NewSADValue : .Text = "New SAD Value"
            .Row = 0
            .Col = enumPreInvoiceDetails.NewExciseValue : .Text = "New Excise Value"
            .Row = 0
            .Col = enumPreInvoiceDetails.TotalExciseValue : .Text = "Total Excise Value"
            .Row = 0
            .Col = enumPreInvoiceDetails.NewTotalExciseValue : .Text = "New Total Exc. Value"
            .Row = 0
            .Col = enumPreInvoiceDetails.TotalExciseValueDiff : .Text = "Total Excise Diff"
            .Row = 0
            .Col = enumPreInvoiceDetails.TotalEcessValue : .Text = "Total ECSS Value"
            .Row = 0
            .Col = enumPreInvoiceDetails.NewEcessValue : .Text = "New ECSS Value"
            .Row = 0
            .Col = enumPreInvoiceDetails.TotalEcssDiff : .Text = "Total ECSS Diff"
            .Row = 0
            .Col = enumPreInvoiceDetails.TotalsEcessValue : .Text = "Total SECSS Value"
            .Row = 0
            .Col = enumPreInvoiceDetails.NewsEcessValue : .Text = "New SECSS Value"
            .Row = 0
            .Col = enumPreInvoiceDetails.TotalsEcssDiff : .Text = "Total SECSS Diff"
            .Row = 0
            .Col = enumPreInvoiceDetails.SalesTaxValue : .Text = "Sales Tax Value"
            .Row = 0
            .Col = enumPreInvoiceDetails.NewSalesTaxValue : .Text = "New S.Tax Value"
            .Row = 0
            .Col = enumPreInvoiceDetails.SalesTaxValueDiff : .Text = "S.Tax Diff"
            .Row = 0
            .Col = enumPreInvoiceDetails.SSTVlaue : .Text = "SSTax Value"
            .Row = 0
            .Col = enumPreInvoiceDetails.NewSSTValue : .Text = "New SSTax Value"
            .Row = 0
            .Col = enumPreInvoiceDetails.SSTVlaueDiff : .Text = "SSTax Diff"
            .Row = 0
            .Col = enumPreInvoiceDetails.Basic : .Text = "Basic Value"
            .Row = 0
            .Col = enumPreInvoiceDetails.NewBasic : .Text = "New Basic Value"
            .Row = 0
            .Col = enumPreInvoiceDetails.BasicDiff : .Text = "Basic Diff"
            .Row = 0
            .Col = enumPreInvoiceDetails.AccessableValue : .Text = "Assessable Value"
            .Row = 0
            .Col = enumPreInvoiceDetails.NewAccessableValue : .Text = "NEW Assessable Value"
            .Row = 0
            .Col = enumPreInvoiceDetails.AccessableValue_Diff : .Text = "Assessable Value Diff"
            .Row = 0
            .Col = enumPreInvoiceDetails.TotalCurrInvValue : .Text = "New Total Value"
            .Row = 0
            .Col = enumPreInvoiceDetails.TotalInvoiceValue : .Text = "Total Value Diff"
            .Row = 0
            .Col = enumPreInvoiceDetails.flag : .Text = "FLAG"
            .Row = 0
            .Col = enumPreInvoiceDetails.CGST_AMT : .Text = "CGST AMT."
            .Row = 0
            .Col = enumPreInvoiceDetails.NEW_CGST_AMT : .Text = "NEW CGST AMT."
            .Row = 0
            .Col = enumPreInvoiceDetails.DIFF_CGST_AMT : .Text = "CGST DIFF."
            .Row = 0
            .Col = enumPreInvoiceDetails.SGST_AMT : .Text = "SGST AMT."
            .Row = 0
            .Col = enumPreInvoiceDetails.NEW_SGST_AMT : .Text = "NEW SGST AMT."
            .Row = 0
            .Col = enumPreInvoiceDetails.DIFF_SGST_AMT : .Text = "SGST DIFF."
            .Row = 0
            .Col = enumPreInvoiceDetails.UTGST_AMT : .Text = "UTGST AMT."
            .Row = 0
            .Col = enumPreInvoiceDetails.NEW_UTGST_AMT : .Text = "NEW UTGST AMT."
            .Row = 0
            .Col = enumPreInvoiceDetails.DIFF_UTGST_AMT : .Text = "UTGST DIFF."
            .Row = 0
            .Col = enumPreInvoiceDetails.IGST_AMT : .Text = "IGST AMT."
            .Row = 0
            .Col = enumPreInvoiceDetails.NEW_IGST_AMT : .Text = "NEW IGST AMT."
            .Row = 0
            .Col = enumPreInvoiceDetails.DIFF_IGST_AMT : .Text = "IGST DIFF."
            .Row = 0
            .Col = enumPreInvoiceDetails.CCESS_AMT : .Text = "CCESS AMT."
            .Row = 0
            .Col = enumPreInvoiceDetails.NEW_CCESS_AMT : .Text = "NEW CCESS AMT."
            .Row = 0
            .Col = enumPreInvoiceDetails.DIFF_CCESS_AMT : .Text = "CCESS DIFF."
            .MaxRows = 0
        End With
        With spdInvDetails
            .MaxCols = enumInvoiceSummery.SummeryInvoiceValue
            .Row = 0
            .Col = enumInvoiceSummery.Rate : .Text = "Rate"
            .Row = 0
            .Col = enumInvoiceSummery.BasicValue : .Text = "Basic"
            .Row = 0
            .Col = enumInvoiceSummery.CustSuppMat : .Text = "Cust Supp Material"
            .Row = 0
            .Col = enumInvoiceSummery.ToolCost : .Text = "Tool Cost"
            .Row = 0
            .Col = enumInvoiceSummery.AccessableValue : .Text = "Accessable Value"
            .Row = 0
            .Col = enumInvoiceSummery.ExciseValue : .Text = "Total Excise Value"
            .Row = 0
            .Col = enumInvoiceSummery.EcssValue : .Text = "Total Ecss Value "
            .Row = 0
            .Col = enumInvoiceSummery.sEcssValue : .Text = "Total SEcss Value "
            .Row = 0
            .Col = enumInvoiceSummery.PackingValue : .Text = "Total Packing Value "
            .Row = 0
            .Col = enumInvoiceSummery.SalesTaxType : .Text = "Sales Tax Value"
            .Row = 0
            .Col = enumInvoiceSummery.SSTType : .Text = "SST Value"
            .Row = 0
            .Col = enumInvoiceSummery.CGST : .Text = "CGST"
            .Row = 0
            .Col = enumInvoiceSummery.SGST : .Text = "SGST"
            .Row = 0
            .Col = enumInvoiceSummery.UTGST : .Text = "UTGST"
            .Row = 0
            .Col = enumInvoiceSummery.IGST  : .Text = "IGST"
            .Row = 0
            .Col = enumInvoiceSummery.CCESS : .Text = "CCESS"
            .Row = 0
            .Col = enumInvoiceSummery.SummeryInvoiceValue : .Text = "Total Value"
            .MaxRows = 0
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Sub SetCellTypeofGrids()
        On Error GoTo ErrHandler
        With spdPrevInv
            .Row = 1
            .Row2 = .MaxRows
            .Col = enumPreInvoiceDetails.Select_Invoice
            .Col2 = enumPreInvoiceDetails.Select_Invoice
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
            .TypeCheckCenter = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.Invoice_No
            .Col2 = enumPreInvoiceDetails.Invoice_No
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.Invoice_Date
            .Col2 = enumPreInvoiceDetails.Invoice_Date
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.LastSupplementary
            .Col2 = enumPreInvoiceDetails.LastSupplementary
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.SupplementaryDate
            .Col2 = enumPreInvoiceDetails.SupplementaryDate
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.Quantity
            .Col2 = enumPreInvoiceDetails.Quantity
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.Rate
            .Col2 = enumPreInvoiceDetails.Rate
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.New_Rate
            .Col2 = enumPreInvoiceDetails.New_Rate
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .BlockMode = False
            .Col = enumPreInvoiceDetails.Rate_diff
            .Col2 = enumPreInvoiceDetails.Rate_diff
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.TotalCustSuppMaterial
            .Col2 = enumPreInvoiceDetails.TotalCustSuppMaterial
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewCustSuppMaterial
            .Col2 = enumPreInvoiceDetails.NewCustSuppMaterial
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewTotalCustSuppMaterial
            .Col2 = enumPreInvoiceDetails.NewTotalCustSuppMaterial
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.CustSuppMaterial_diff
            .Col2 = enumPreInvoiceDetails.CustSuppMaterial_diff
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.ToolCost
            .Col2 = enumPreInvoiceDetails.ToolCost
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.newToolCost
            .Col2 = enumPreInvoiceDetails.newToolCost
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewTotalToolCost
            .Col2 = enumPreInvoiceDetails.NewTotalToolCost
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.ToolCost_diff
            .Col2 = enumPreInvoiceDetails.ToolCost_diff
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.TotalPacking
            .Col2 = enumPreInvoiceDetails.TotalPacking
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewPacking
            .Col2 = enumPreInvoiceDetails.NewPacking
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewTotalPacking
            .Col2 = enumPreInvoiceDetails.NewTotalPacking
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewCVDValue
            .Col2 = enumPreInvoiceDetails.NewCVDValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewSADValue
            .Col2 = enumPreInvoiceDetails.NewSADValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewExciseValue
            .Col2 = enumPreInvoiceDetails.NewExciseValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.TotalExciseValue
            .Col2 = enumPreInvoiceDetails.TotalExciseValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewTotalExciseValue
            .Col2 = enumPreInvoiceDetails.NewTotalExciseValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.TotalExciseValueDiff
            .Col2 = enumPreInvoiceDetails.TotalExciseValueDiff
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.TotalEcessValue
            .Col2 = enumPreInvoiceDetails.TotalEcessValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Value = 0.0#
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewEcessValue
            .Col2 = enumPreInvoiceDetails.NewEcessValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Value = 0.0#
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.TotalEcssDiff
            .Col2 = enumPreInvoiceDetails.TotalEcssDiff
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Value = 0.0#
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.TotalsEcessValue
            .Col2 = enumPreInvoiceDetails.TotalsEcessValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Value = 0.0#
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewsEcessValue
            .Col2 = enumPreInvoiceDetails.NewsEcessValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Value = 0.0#
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.TotalsEcssDiff
            .Col2 = enumPreInvoiceDetails.TotalsEcssDiff
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Value = 0.0#
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.SalesTaxValue
            .Col2 = enumPreInvoiceDetails.SalesTaxValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewSalesTaxValue
            .Col2 = enumPreInvoiceDetails.NewSalesTaxValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.SalesTaxValueDiff
            .Col2 = enumPreInvoiceDetails.SalesTaxValueDiff
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.SSTVlaue
            .Col2 = enumPreInvoiceDetails.SSTVlaue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewSSTValue
            .Col2 = enumPreInvoiceDetails.NewSSTValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.SSTVlaueDiff
            .Col2 = enumPreInvoiceDetails.SSTVlaueDiff
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.Basic
            .Col2 = enumPreInvoiceDetails.Basic
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewBasic
            .Col2 = enumPreInvoiceDetails.NewBasic
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.BasicDiff
            .Col2 = enumPreInvoiceDetails.BasicDiff
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.AccessableValue
            .Col2 = enumPreInvoiceDetails.AccessableValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewAccessableValue
            .Col2 = enumPreInvoiceDetails.NewAccessableValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumPreInvoiceDetails.AccessableValue_Diff
            .Col2 = enumPreInvoiceDetails.AccessableValue_Diff
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .BlockMode = False
            '101188073
            .Col = enumPreInvoiceDetails.CGST_AMT
            .Col2 = enumPreInvoiceDetails.DIFF_CCESS_AMT
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            '101188073
            .Col = enumPreInvoiceDetails.TotalCurrInvValue
            .Col2 = enumPreInvoiceDetails.TotalCurrInvValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .BlockMode = False
            .Col = enumPreInvoiceDetails.TotalInvoiceValue
            .Col2 = enumPreInvoiceDetails.TotalInvoiceValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .BlockMode = False
            .Col = enumPreInvoiceDetails.flag
            .Col2 = enumPreInvoiceDetails.flag
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .BlockMode = False
        End With
        With spdInvDetails
            .MaxRows = 1
            .Row = 1
            .Row2 = .MaxRows
            .Col = enumInvoiceSummery.Rate
            .Col2 = enumInvoiceSummery.Rate
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumInvoiceSummery.BasicValue
            .Col2 = enumInvoiceSummery.BasicValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumInvoiceSummery.CustSuppMat
            .Col2 = enumInvoiceSummery.CustSuppMat
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumInvoiceSummery.ToolCost
            .Col2 = enumInvoiceSummery.ToolCost
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumInvoiceSummery.AccessableValue
            .Col2 = enumInvoiceSummery.AccessableValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumInvoiceSummery.ExciseValue
            .Col2 = enumInvoiceSummery.ExciseValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumInvoiceSummery.EcssValue
            .Col2 = enumInvoiceSummery.EcssValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumInvoiceSummery.sEcssValue
            .Col2 = enumInvoiceSummery.sEcssValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumInvoiceSummery.PackingValue
            .Col2 = enumInvoiceSummery.PackingValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumInvoiceSummery.SalesTaxType
            .Col2 = enumInvoiceSummery.SalesTaxType
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            .Col = enumInvoiceSummery.SSTType
            .Col2 = enumInvoiceSummery.SSTType
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            '101188073
            .Col = enumInvoiceSummery.CGST
            .Col2 = enumInvoiceSummery.CCESS
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
            '101188073
            .Col = enumInvoiceSummery.SummeryInvoiceValue
            .Col2 = enumInvoiceSummery.SummeryInvoiceValue
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Lock = True
            .BlockMode = False
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function SetMaxLengthofGrid(ByRef pintMaxDecimalPlaces As Object) As Object
        Dim strMin As String
        Dim strMax As String
        Dim intLoopCounter As Short
        On Error GoTo ErrHandler
        If pintMaxDecimalPlaces < 2 Then
            pintMaxDecimalPlaces = 2
        End If
        strMin = "0." : strMax = "99999999999999."
        For intLoopCounter = 1 To pintMaxDecimalPlaces
            strMin = strMin & "0"
            strMax = strMax & "9"
        Next
        With spdPrevInv
            .Row = 1
            .Row2 = .MaxRows
            .Col = enumPreInvoiceDetails.New_Rate
            .Col2 = enumPreInvoiceDetails.New_Rate
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = pintMaxDecimalPlaces
            .TypeFloatMin = strMin
            .TypeFloatMax = strMax
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewCustSuppMaterial
            .Col2 = enumPreInvoiceDetails.NewCustSuppMaterial
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = pintMaxDecimalPlaces
            .TypeFloatMin = strMin
            .TypeFloatMax = strMax
            .BlockMode = False
            .Col = enumPreInvoiceDetails.newToolCost
            .Col2 = enumPreInvoiceDetails.newToolCost
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = pintMaxDecimalPlaces
            .TypeFloatMin = strMin
            .TypeFloatMax = strMax
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewPacking
            .Col2 = enumPreInvoiceDetails.NewPacking
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = pintMaxDecimalPlaces
            .TypeFloatMin = strMin
            .TypeFloatMax = strMax
            .BlockMode = False
        End With
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Sub SetFormulaofColumns(ByRef pintRow As Integer, ByRef pintDecimal As Short)
        Dim strFormula As String
        Dim strParamQuery As String
        Dim rsParameterData As ClsResultSetDB
        Dim blnISInsExcisable As Boolean
        Dim blnEOUFlag As Boolean
        Dim blnISExciseRoundOff As Boolean
        Dim blnISBasicRoundOff As Boolean
        Dim blnISSalesTaxRoundOff As Boolean
        Dim blnISSurChargeTaxRoundOff As Boolean
        Dim blnECSSRoundoff As Boolean
        Dim dblEcessValue As Double
        Dim blnAddCustMatrl As Boolean
        Dim intGSTRoundOffDecimal As Integer
        Dim blnGSTRoundOff As Boolean
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        On Error GoTo ErrHandler
        rsParameterData = New ClsResultSetDB
        strParamQuery = "SELECT InsExc_Excise=isnull(InsExc_Excise,0),CustSupp_Inc=isnull(CustSupp_Inc,0),EOU_Flag=isnull(EOU_Flag,0),SalesTax_Roundoff=isnull(SalesTax_Roundoff,0),Basic_roundoff=isnull(Basic_roundoff,0),Excise_Roundoff=isnull(Excise_Roundoff,0),SST_Roundoff=isnull(SST_Roundoff,0),ECESS_Roundoff=isnull(ECESS_Roundoff,0),Basic_Roundoff_Decimal=isnull(Basic_Roundoff_Decimal,0),SalesTax_Roundoff_Decimal=isnull(SalesTax_Roundoff_Decimal,0),Excise_Roundoff_Decimal=isnull(Excise_Roundoff_Decimal,0),SST_Roundoff_Decimal=isnull(SST_Roundoff_Decimal,0),TCSTax_Roundoff_Decimal=isnull(TCSTax_Roundoff_Decimal,0),TotalToolCostRoundOff_Decimal=isnull(TotalToolCostRoundOff_Decimal,0),ECESSRoundOff_Decimal=Isnull(ECESSRoundOff_Decimal,0),GSTTAX_ROUNDOFF_DECIMAL=ISNULL(GSTTAX_ROUNDOFF_DECIMAL,0),GSTTAX_ROUNDOFF=ISNULL(GSTTAX_ROUNDOFF,0) FROM Sales_Parameter where Unit_code = '" & gstrUnitId & "' "
        rsParameterData.GetResult(strParamQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsParameterData.GetNoRows > 0 Then
            blnEOUFlag = rsParameterData.GetValue("EOU_Flag")
            blnISExciseRoundOff = rsParameterData.GetValue("Excise_Roundoff")
            blnISBasicRoundOff = rsParameterData.GetValue("Basic_Roundoff")
            blnISSalesTaxRoundOff = rsParameterData.GetValue("SalesTax_Roundoff")
            blnISSurChargeTaxRoundOff = rsParameterData.GetValue("SST_Roundoff")
            blnAddCustMatrl = rsParameterData.GetValue("CustSupp_Inc")
            blnECSSRoundoff = rsParameterData.GetValue("ECESS_Roundoff")
            intGSTRoundOffDecimal = rsParameterData.GetValue("GSTTAX_ROUNDOFF_DECIMAL")
            blnGSTRoundOff = rsParameterData.GetValue("GSTTAX_ROUNDOFF")
        Else
            MsgBox("No data define in Sales_Parameter Table", MsgBoxStyle.Information, "empower")
            rsParameterData.ResultSetClose()
            rsParameterData = Nothing
            Exit Sub
        End If
        With spdPrevInv
            .AutoCalc = True
            '********To Set Value For Loop
            If pintRow = 0 Then
                pintRow = 1
                intMaxLoop = .MaxRows
            Else
                intMaxLoop = pintRow
            End If
            If pintDecimal <= 0 Then pintDecimal = 2
            For intLoopCounter = pintRow To intMaxLoop
                .Row = intLoopCounter
                .Row2 = intLoopCounter
                'Formula For Rate diff
                strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.New_Rate & " - R" & intLoopCounter & "C" & enumPreInvoiceDetails.Rate & "," & pintDecimal & ")"
                .Col = enumPreInvoiceDetails.Rate_diff
                .Col2 = enumPreInvoiceDetails.Rate_diff
                .BlockMode = True
                .Formula = strFormula
                .Action = FPSpreadADO.ActionConstants.ActionReCalc
                .BlockMode = False
                'Formula For New Total Packing Diff
                strFormula = "Round((R" & intLoopCounter & "C" & enumPreInvoiceDetails.Quantity & "* R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewPacking & ") - "
                strFormula = strFormula & "R" & intLoopCounter & "C" & enumPreInvoiceDetails.TotalPacking & "," & pintDecimal & ")"
                .Col = enumPreInvoiceDetails.NewTotalPacking
                .Col2 = enumPreInvoiceDetails.NewTotalPacking
                .BlockMode = True
                .Formula = strFormula
                .Action = FPSpreadADO.ActionConstants.ActionReCalc
                .BlockMode = False
                'Formula For Diff in Total Customer Basic
                If blnISBasicRoundOff = True Then
                    strFormula = "Round((R" & intLoopCounter & "C" & enumPreInvoiceDetails.New_Rate & " + "
                    strFormula = strFormula & "R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewPacking & ") * R" & intLoopCounter & "C" & enumPreInvoiceDetails.Quantity & ",0)"
                Else
                    strFormula = "Round((R" & intLoopCounter & "C" & enumPreInvoiceDetails.New_Rate & " + "
                    strFormula = strFormula & "R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewPacking & ") * R" & intLoopCounter & "C" & enumPreInvoiceDetails.Quantity & "," & rsParameterData.GetValue("Basic_Roundoff_Decimal") & ")"
                End If
                .Col = enumPreInvoiceDetails.NewBasic
                .Col2 = enumPreInvoiceDetails.NewBasic
                .BlockMode = True
                .Formula = strFormula
                .Action = FPSpreadADO.ActionConstants.ActionReCalc
                .BlockMode = False
                'Formula For Diff in Total Customer Diff Basic
                If blnISBasicRoundOff = True Then
                    strFormula = "Round((R" & intLoopCounter & "C" & enumPreInvoiceDetails.Rate_diff & " + "
                    strFormula = strFormula & "R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewPacking & ") * R" & intLoopCounter & "C" & enumPreInvoiceDetails.Quantity & ",0)"
                Else
                    strFormula = "Round((R" & intLoopCounter & "C" & enumPreInvoiceDetails.Rate_diff & " + "
                    strFormula = strFormula & "R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewPacking & ") * R" & intLoopCounter & "C" & enumPreInvoiceDetails.Quantity & "," & rsParameterData.GetValue("Basic_Roundoff_Decimal") & ")"
                End If
                .Col = enumPreInvoiceDetails.BasicDiff
                .Col2 = enumPreInvoiceDetails.BasicDiff
                .BlockMode = True
                .Formula = strFormula
                .Action = FPSpreadADO.ActionConstants.ActionReCalc
                .BlockMode = False
                'Formula For Total New Customer supplied Material
                strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewCustSuppMaterial & " * R" & intLoopCounter & "C" & enumPreInvoiceDetails.Quantity & ",2)"
                .Col = enumPreInvoiceDetails.NewTotalCustSuppMaterial
                .Col2 = enumPreInvoiceDetails.NewTotalCustSuppMaterial
                .BlockMode = True
                .Formula = strFormula
                .BlockMode = False
                'Formula For Diff in Total Customer supplied Material
                strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewTotalCustSuppMaterial & " - R" & intLoopCounter & "C" & enumPreInvoiceDetails.TotalCustSuppMaterial & "," & pintDecimal & ")"
                .Col = enumPreInvoiceDetails.CustSuppMaterial_diff
                .Col2 = enumPreInvoiceDetails.CustSuppMaterial_diff
                .BlockMode = True
                .Formula = strFormula
                .BlockMode = False
                'Formula For Diff in Total Customer Tool Cost
                strFormula = "Round((R" & intLoopCounter & "C" & enumPreInvoiceDetails.newToolCost & " * R" & intLoopCounter & "C" & enumPreInvoiceDetails.Quantity & ")," & rsParameterData.GetValue("TotalToolCostRoundOff_Decimal") & ")"
                .Col = enumPreInvoiceDetails.NewTotalToolCost
                .Col2 = enumPreInvoiceDetails.NewTotalToolCost
                .BlockMode = True
                .Formula = strFormula
                .BlockMode = False
                'Formula For Diff in Total Customer Tool Cost diff
                strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewTotalToolCost & " - R" & intLoopCounter & "C" & enumPreInvoiceDetails.ToolCost & "," & rsParameterData.GetValue("TotalToolCostRoundOff_Decimal") & ")"
                .Col = enumPreInvoiceDetails.ToolCost_diff
                .Col2 = enumPreInvoiceDetails.ToolCost_diff
                .BlockMode = True
                .Formula = strFormula
                .BlockMode = False
                'Formula For Total New Customer Accessable Value
                strFormula = "round((R" & intLoopCounter & "C" & enumPreInvoiceDetails.New_Rate & " + "
                strFormula = strFormula & "R" & intLoopCounter & "C" & enumPreInvoiceDetails.newToolCost & " + "
                strFormula = strFormula & "R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewCustSuppMaterial & " + "
                strFormula = strFormula & "((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewPacking & " * "
                strFormula = strFormula & "R" & intLoopCounter & "C" & enumPreInvoiceDetails.New_Rate & ")/100))* "
                strFormula = strFormula & "R" & intLoopCounter & "C" & enumPreInvoiceDetails.Quantity & "," & pintDecimal & ")"
                .Col = enumPreInvoiceDetails.NewAccessableValue
                .Col2 = enumPreInvoiceDetails.NewAccessableValue
                .BlockMode = True
                .Formula = strFormula
                .Action = FPSpreadADO.ActionConstants.ActionReCalc
                .BlockMode = False
                'Formula For Total New Customer Accessable Value Diff
                strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.Rate_diff & " + "
                strFormula = strFormula & "((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewPacking & " * "
                strFormula = strFormula & "R" & intLoopCounter & "C" & enumPreInvoiceDetails.Rate_diff & ")/100)) * "
                strFormula = strFormula & "R" & intLoopCounter & "C" & enumPreInvoiceDetails.Quantity
                strFormula = strFormula & ") + (R" & intLoopCounter & "C" & enumPreInvoiceDetails.ToolCost_diff & " + "
                strFormula = strFormula & "R" & intLoopCounter & "C" & enumPreInvoiceDetails.CustSuppMaterial_diff & "),2)"
                .Col = enumPreInvoiceDetails.AccessableValue_Diff
                .Col2 = enumPreInvoiceDetails.AccessableValue_Diff
                .BlockMode = True
                .Formula = strFormula
                .Action = FPSpreadADO.ActionConstants.ActionReCalc
                .BlockMode = False
                If blnEOUFlag = True Then
                    'Formula For New Excise Value
                    If blnISExciseRoundOff = True Then
                        strFormula = "Round((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewAccessableValue & " * " & Val(lblExctax_Per.Text) & ",0)"
                    Else
                        strFormula = "Round((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewAccessableValue & ") * " & Val(lblExctax_Per.Text) & "," & rsParameterData.GetValue("Excise_Roundoff_Decimal") & ")"
                    End If
                    .Col = enumPreInvoiceDetails.NewExciseValue
                    .Col2 = enumPreInvoiceDetails.NewExciseValue
                    .BlockMode = True
                    .Formula = strFormula
                    .Action = FPSpreadADO.ActionConstants.ActionReCalc
                    .BlockMode = False
                    'Formula For New CVD Value
                    If blnISExciseRoundOff = True Then
                        strFormula = "Round((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewExciseValue & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewAccessableValue & ") * "
                        strFormula = strFormula & lblCVD_Per.Text & ",0)"
                    Else
                        strFormula = "Round((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewExciseValue & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewAccessableValue & ") * "
                        strFormula = strFormula & lblCVD_Per.Text & "," & pintDecimal & ")"
                    End If
                    .Col = enumPreInvoiceDetails.NewCVDValue
                    .Col2 = enumPreInvoiceDetails.NewCVDValue
                    .BlockMode = True
                    .Formula = strFormula
                    .Action = FPSpreadADO.ActionConstants.ActionReCalc
                    .BlockMode = False
                    'Formula For New SAD Value
                    If blnISExciseRoundOff = True Then
                        strFormula = "Round((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewExciseValue & " + R" & intLoopCounter & "C"
                        strFormula = strFormula & enumPreInvoiceDetails.NewAccessableValue & "+ R" & intLoopCounter
                        strFormula = strFormula & "C" & enumPreInvoiceDetails.NewCVDValue & ") * " & lblCVD_Per.Text & ",0)"
                    Else
                        strFormula = "Round((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewExciseValue & " + R" & intLoopCounter & "C"
                        strFormula = strFormula & enumPreInvoiceDetails.NewAccessableValue & "+ R" & intLoopCounter
                        strFormula = strFormula & "C" & enumPreInvoiceDetails.NewCVDValue & ") * " & lblCVD_Per.Text & "," & pintDecimal & ")"
                    End If
                    .Col = enumPreInvoiceDetails.NewCVDValue
                    .Col2 = enumPreInvoiceDetails.NewCVDValue
                    .BlockMode = True
                    .Formula = strFormula
                    .Action = FPSpreadADO.ActionConstants.ActionReCalc
                    .BlockMode = False
                End If
                'Formula For Total New Excise Value
                If blnEOUFlag = False Then
                    If blnISExciseRoundOff = True Then
                        strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewAccessableValue & "*"
                        strFormula = strFormula & Val(lblExctax_Per.Text) & ")/100),0)"
                    Else
                        strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewAccessableValue & "*"
                        strFormula = strFormula & Val(lblExctax_Per.Text) & ")/100)" & "," & rsParameterData.GetValue("Excise_Roundoff_Decimal") & ")"
                    End If
                    .Col = enumPreInvoiceDetails.NewExciseValue
                    .Col2 = enumPreInvoiceDetails.NewExciseValue
                    .BlockMode = True
                    .Formula = strFormula
                    .Action = FPSpreadADO.ActionConstants.ActionReCalc
                    .BlockMode = False
                Else
                    If blnISExciseRoundOff = True Then
                        strFormula = "Round((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewExciseValue & "/2) + (R" & intLoopCounter
                        strFormula = strFormula & "C" & enumPreInvoiceDetails.NewCVDValue & "/2) + (R" & intLoopCounter & enumPreInvoiceDetails.NewSADValue & "/2),0)"
                    Else
                        strFormula = "Round((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewExciseValue & "/2) + (R" & intLoopCounter
                        strFormula = strFormula & "C" & enumPreInvoiceDetails.NewCVDValue & "/2) + (R" & intLoopCounter & enumPreInvoiceDetails.NewSADValue & "/2)" & "," & rsParameterData.GetValue("Excise_Roundoff_Decimal") & ")"
                    End If
                End If
                .Col = enumPreInvoiceDetails.NewTotalExciseValue
                .Col2 = enumPreInvoiceDetails.NewTotalExciseValue
                .BlockMode = True
                .Formula = strFormula
                .Action = FPSpreadADO.ActionConstants.ActionReCalc
                .BlockMode = False
                'Formula For Total New Excise Value diff
                If blnEOUFlag = False Then
                    If blnISExciseRoundOff = True Then
                        strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.AccessableValue_Diff & "*"
                        strFormula = strFormula & Val(lblExctax_Per.Text) & ")/100),0)"
                    Else
                        strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.AccessableValue_Diff & "*"
                        strFormula = strFormula & Val(lblExctax_Per.Text) & ")/100)" & "," & rsParameterData.GetValue("Excise_Roundoff_Decimal") & ")"
                    End If
                Else
                    strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewTotalExciseValue & " - R" & intLoopCounter & "C" & enumPreInvoiceDetails.TotalExciseValue & "," & pintDecimal & ")"
                End If
                .Col = enumPreInvoiceDetails.TotalExciseValueDiff
                .Col2 = enumPreInvoiceDetails.TotalExciseValueDiff
                .BlockMode = True
                .Formula = strFormula
                .Action = FPSpreadADO.ActionConstants.ActionReCalc
                .BlockMode = False
                'Calculate Ecss Value
                If blnECSSRoundoff = True Then
                    strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewTotalExciseValue & "*" & Val(Me.lblEcssCode.Text) & ") /100),0)"
                Else
                    strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewTotalExciseValue & "*" & Val(Me.lblEcssCode.Text) & ")/100)," & rsParameterData.GetValue("ECESSRoundOff_Decimal") & ")"
                End If
                .Col = enumPreInvoiceDetails.NewEcessValue
                .Col2 = enumPreInvoiceDetails.NewEcessValue
                .BlockMode = True
                .Formula = strFormula
                .Action = FPSpreadADO.ActionConstants.ActionReCalc
                .BlockMode = False
                'Calculate Ecss Difference
                If blnECSSRoundoff = True Then
                    strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewEcessValue & " - R" & intLoopCounter & "C" & enumPreInvoiceDetails.TotalEcessValue & ",0)"
                Else
                    strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewEcessValue & " - R" & intLoopCounter & "C" & enumPreInvoiceDetails.TotalEcessValue & "," & rsParameterData.GetValue("ECESSRoundOff_Decimal") & ")"
                End If
                .Col = enumPreInvoiceDetails.TotalEcssDiff
                .Col2 = enumPreInvoiceDetails.TotalEcssDiff
                .BlockMode = True
                .Formula = strFormula
                .Action = FPSpreadADO.ActionConstants.ActionReCalc
                .BlockMode = False
                ''--- Calculate SEcss Value
                If blnECSSRoundoff = True Then
                    strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewTotalExciseValue & "*" & Val(lblSEcssCode.Text) & ") /100),0)"
                Else
                    strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewTotalExciseValue & "*" & Val(lblSEcssCode.Text) & ")/100)," & rsParameterData.GetValue("ECESSRoundOff_Decimal") & ")"
                End If
                .Col = enumPreInvoiceDetails.NewsEcessValue
                .Col2 = enumPreInvoiceDetails.NewsEcessValue
                .BlockMode = True
                .Formula = strFormula
                .Action = FPSpreadADO.ActionConstants.ActionReCalc
                .BlockMode = False
                ''--- Calculate SEcss Difference
                If blnECSSRoundoff = True Then
                    strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewsEcessValue & " - R" & intLoopCounter & "C" & enumPreInvoiceDetails.TotalsEcessValue & ",0)"
                Else
                    strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewsEcessValue & " - R" & intLoopCounter & "C" & enumPreInvoiceDetails.TotalsEcessValue & "," & rsParameterData.GetValue("ECESSRoundOff_Decimal") & ")"
                End If
                .Col = enumPreInvoiceDetails.TotalsEcssDiff
                .Col2 = enumPreInvoiceDetails.TotalsEcssDiff
                .BlockMode = True
                .Formula = strFormula
                .Action = FPSpreadADO.ActionConstants.ActionReCalc
                .BlockMode = False
                'Formula For Total New SalesTaxValue Value
                If blnISSalesTaxRoundOff = True Then
                    strFormula = "round((((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewBasic & "+" & "R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewTotalExciseValue & "+" & "R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewEcessValue & "+" & "R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewsEcessValue & ")" & "*"
                    strFormula = strFormula & Val(lblSaltax_Per.Text) & ")/100),0)"
                Else
                    strFormula = "round((((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewBasic & "+" & "R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewTotalExciseValue & "+" & "R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewEcessValue & "+" & "R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewsEcessValue & ")" & "*"
                    strFormula = strFormula & Val(lblSaltax_Per.Text) & ")/100)," & rsParameterData.GetValue("SalesTax_Roundoff_Decimal") & ")"
                End If
                .Col = enumPreInvoiceDetails.NewSalesTaxValue
                .Col2 = enumPreInvoiceDetails.NewSalesTaxValue
                .BlockMode = True
                .Formula = strFormula
                .Action = FPSpreadADO.ActionConstants.ActionReCalc
                .BlockMode = False
                'Formula For Total New SalesTaxValue Value diff
                If blnISSalesTaxRoundOff = True Then
                    strFormula = "round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.BasicDiff & "*"
                    strFormula = strFormula & Val(lblSaltax_Per.Text) & ")/100),0)"
                Else
                    strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.BasicDiff & "*"
                    strFormula = strFormula & Val(lblSaltax_Per.Text) & ")/100)" & "," & rsParameterData.GetValue("SalesTax_Roundoff_Decimal") & ")"
                End If
                .Col = enumPreInvoiceDetails.SalesTaxValueDiff
                .Col2 = enumPreInvoiceDetails.SalesTaxValueDiff
                .BlockMode = True
                .Formula = strFormula
                .Action = FPSpreadADO.ActionConstants.ActionReCalc
                .BlockMode = False
                'Formula For Total New Surcharge Value
                If blnISSurChargeTaxRoundOff = True Then
                    strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewSalesTaxValue & "*"
                    strFormula = strFormula & Val(lblSurcharge_Per.Text) & ")/100),0)"
                Else
                    strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewSalesTaxValue & "*"
                    strFormula = strFormula & Val(lblSurcharge_Per.Text) & ")/100)" & "," & rsParameterData.GetValue("SST_Roundoff_Decimal") & ")"
                End If
                .Col = enumPreInvoiceDetails.NewSSTValue
                .Col2 = enumPreInvoiceDetails.NewSSTValue
                .BlockMode = True
                .Formula = strFormula
                .Action = FPSpreadADO.ActionConstants.ActionReCalc
                .BlockMode = False
                'Formula For Total New Surcharge Value diff
                If blnISSurChargeTaxRoundOff = True Then
                    strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.SalesTaxValueDiff & "*"
                    strFormula = strFormula & Val(lblSurcharge_Per.Text) & ")/100),0)"
                Else
                    strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.SalesTaxValueDiff & "*"
                    strFormula = strFormula & Val(lblSurcharge_Per.Text) & ")/100)" & "," & rsParameterData.GetValue("SST_Roundoff_Decimal") & ")"
                End If
                .Col = enumPreInvoiceDetails.SSTVlaueDiff
                .Col2 = enumPreInvoiceDetails.SSTVlaue
                .BlockMode = True
                .Formula = strFormula
                .Action = FPSpreadADO.ActionConstants.ActionReCalc
                .BlockMode = False
                '101188073 GST Changes
                If gblnGSTUnit Then
                    'Formula for New CGST AMT.
                    If blnGSTRoundOff Then
                        strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewAccessableValue & "*" & Val(Me.lblCGSTPercent.Text) & ") /100),0)"
                    Else
                        strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewAccessableValue & "*" & Val(Me.lblCGSTPercent.Text) & ") /100)," & intGSTRoundOffDecimal & ")"
                    End If
                    .Col = enumPreInvoiceDetails.NEW_CGST_AMT
                    .Col2 = enumPreInvoiceDetails.NEW_CGST_AMT
                    .BlockMode = True
                    .Formula = strFormula
                    .Action = FPSpreadADO.ActionConstants.ActionReCalc
                    .BlockMode = False
                    'Formula for CGST AMT. Diff.
                    If blnGSTRoundOff Then
                        strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NEW_CGST_AMT & " - R" & intLoopCounter & "C" & enumPreInvoiceDetails.CGST_AMT & ",0)"
                    Else
                        strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NEW_CGST_AMT & " - R" & intLoopCounter & "C" & enumPreInvoiceDetails.CGST_AMT & "," & intGSTRoundOffDecimal & ")"
                    End If
                    .Col = enumPreInvoiceDetails.DIFF_CGST_AMT
                    .Col2 = enumPreInvoiceDetails.DIFF_CGST_AMT
                    .BlockMode = True
                    .Formula = strFormula
                    .Action = FPSpreadADO.ActionConstants.ActionReCalc
                    .BlockMode = False
                    'Formula for New SGST AMT.
                    If blnGSTRoundOff Then
                        strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewAccessableValue & "*" & Val(Me.lblSGSTPercent.Text) & ") /100),0)"
                    Else
                        strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewAccessableValue & "*" & Val(Me.lblSGSTPercent.Text) & ") /100)," & intGSTRoundOffDecimal & ")"
                    End If
                    .Col = enumPreInvoiceDetails.NEW_SGST_AMT
                    .Col2 = enumPreInvoiceDetails.NEW_SGST_AMT
                    .BlockMode = True
                    .Formula = strFormula
                    .Action = FPSpreadADO.ActionConstants.ActionReCalc
                    .BlockMode = False
                    'Formula for SGST AMT. Diff.
                    If blnGSTRoundOff Then
                        strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NEW_SGST_AMT & " - R" & intLoopCounter & "C" & enumPreInvoiceDetails.SGST_AMT & ",0)"
                    Else
                        strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NEW_SGST_AMT & " - R" & intLoopCounter & "C" & enumPreInvoiceDetails.SGST_AMT & "," & intGSTRoundOffDecimal & ")"
                    End If
                    .Col = enumPreInvoiceDetails.DIFF_SGST_AMT
                    .Col2 = enumPreInvoiceDetails.DIFF_SGST_AMT
                    .BlockMode = True
                    .Formula = strFormula
                    .Action = FPSpreadADO.ActionConstants.ActionReCalc
                    .BlockMode = False
                    'Formula for New UTGST AMT.
                    If blnGSTRoundOff Then
                        strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewAccessableValue & "*" & Val(Me.lblUTGSTPercent.Text) & ") /100),0)"
                    Else
                        strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewAccessableValue & "*" & Val(Me.lblUTGSTPercent.Text) & ") /100)," & intGSTRoundOffDecimal & ")"
                    End If
                    .Col = enumPreInvoiceDetails.NEW_UTGST_AMT
                    .Col2 = enumPreInvoiceDetails.NEW_UTGST_AMT
                    .BlockMode = True
                    .Formula = strFormula
                    .Action = FPSpreadADO.ActionConstants.ActionReCalc
                    .BlockMode = False
                    'Formula for UTGST AMT. Diff.
                    If blnGSTRoundOff Then
                        strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NEW_UTGST_AMT & " - R" & intLoopCounter & "C" & enumPreInvoiceDetails.UTGST_AMT & ",0)"
                    Else
                        strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NEW_UTGST_AMT & " - R" & intLoopCounter & "C" & enumPreInvoiceDetails.UTGST_AMT & "," & intGSTRoundOffDecimal & ")"
                    End If
                    .Col = enumPreInvoiceDetails.DIFF_UTGST_AMT
                    .Col2 = enumPreInvoiceDetails.DIFF_UTGST_AMT
                    .BlockMode = True
                    .Formula = strFormula
                    .Action = FPSpreadADO.ActionConstants.ActionReCalc
                    .BlockMode = False
                    'Formula for New IGST AMT.
                    If blnGSTRoundOff Then
                        strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewAccessableValue & "*" & Val(Me.lblIGSTPercent.Text) & ") /100),0)"
                    Else
                        strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewAccessableValue & "*" & Val(Me.lblIGSTPercent.Text) & ") /100)," & intGSTRoundOffDecimal & ")"
                    End If
                    .Col = enumPreInvoiceDetails.NEW_IGST_AMT
                    .Col2 = enumPreInvoiceDetails.NEW_IGST_AMT
                    .BlockMode = True
                    .Formula = strFormula
                    .Action = FPSpreadADO.ActionConstants.ActionReCalc
                    .BlockMode = False
                    'Formula for IGST AMT. Diff.
                    If blnGSTRoundOff Then
                        strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NEW_IGST_AMT & " - R" & intLoopCounter & "C" & enumPreInvoiceDetails.IGST_AMT & ",0)"
                    Else
                        strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NEW_IGST_AMT & " - R" & intLoopCounter & "C" & enumPreInvoiceDetails.IGST_AMT & "," & intGSTRoundOffDecimal & ")"
                    End If
                    .Col = enumPreInvoiceDetails.DIFF_IGST_AMT
                    .Col2 = enumPreInvoiceDetails.DIFF_IGST_AMT
                    .BlockMode = True
                    .Formula = strFormula
                    .Action = FPSpreadADO.ActionConstants.ActionReCalc
                    .BlockMode = False
                    'Formula for New CCESS AMT.
                    If blnGSTRoundOff Then
                        strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewAccessableValue & "*" & Val(Me.lblCompCessPercent.Text) & ") /100),0)"
                    Else
                        strFormula = "Round(((R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewAccessableValue & "*" & Val(Me.lblCompCessPercent.Text) & ") /100)," & intGSTRoundOffDecimal & ")"
                    End If
                    .Col = enumPreInvoiceDetails.NEW_CCESS_AMT
                    .Col2 = enumPreInvoiceDetails.NEW_CCESS_AMT
                    .BlockMode = True
                    .Formula = strFormula
                    .Action = FPSpreadADO.ActionConstants.ActionReCalc
                    .BlockMode = False
                    'Formula for CCESS AMT. Diff.
                    If blnGSTRoundOff Then
                        strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NEW_CCESS_AMT & " - R" & intLoopCounter & "C" & enumPreInvoiceDetails.CCESS_AMT & ",0)"
                    Else
                        strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NEW_CCESS_AMT & " - R" & intLoopCounter & "C" & enumPreInvoiceDetails.CCESS_AMT & "," & intGSTRoundOffDecimal & ")"
                    End If
                    .Col = enumPreInvoiceDetails.DIFF_CCESS_AMT
                    .Col2 = enumPreInvoiceDetails.DIFF_CCESS_AMT
                    .BlockMode = True
                    .Formula = strFormula
                    .Action = FPSpreadADO.ActionConstants.ActionReCalc
                    .BlockMode = False
                End If
                '101188073
                'Formula For Total Invoice Value
                If gblnGSTUnit Then  '101188073
                    If blnAddCustMatrl = True Then
                        strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewBasic
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.NEW_CGST_AMT
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.NEW_SGST_AMT
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.NEW_UTGST_AMT
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.NEW_IGST_AMT
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.NEW_CCESS_AMT
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewTotalCustSuppMaterial
                        strFormula = strFormula & "," & pintDecimal & ")"
                    Else
                        strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewBasic
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.NEW_CGST_AMT
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.NEW_SGST_AMT
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.NEW_UTGST_AMT
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.NEW_IGST_AMT
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.NEW_CCESS_AMT
                        strFormula = strFormula & "," & pintDecimal & ")"
                    End If
                Else
                    If blnAddCustMatrl = True Then
                        strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewBasic & " + R" & intLoopCounter & "C"
                        strFormula = strFormula & enumPreInvoiceDetails.NewSalesTaxValue & "+ R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewSSTValue
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewTotalExciseValue
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewEcessValue
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewsEcessValue
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewTotalCustSuppMaterial & "," & pintDecimal & ")"
                    Else
                        strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewBasic & " + R" & intLoopCounter & "C"
                        strFormula = strFormula & enumPreInvoiceDetails.NewSalesTaxValue & " + R" & intLoopCounter & "C"
                        strFormula = strFormula & enumPreInvoiceDetails.NewSSTValue & " + R" & intLoopCounter
                        strFormula = strFormula & "C" & enumPreInvoiceDetails.NewTotalExciseValue
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewEcessValue
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.NewsEcessValue & "," & pintDecimal & ")"
                    End If
                End If
                .Col = enumPreInvoiceDetails.TotalCurrInvValue
                .Col2 = enumPreInvoiceDetails.TotalCurrInvValue
                .BlockMode = True
                .Formula = strFormula
                .Action = FPSpreadADO.ActionConstants.ActionReCalc
                .BlockMode = False
                If gblnGSTUnit Then  '101188073
                    If blnAddCustMatrl = True Then
                        strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.BasicDiff
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.DIFF_CGST_AMT
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.DIFF_SGST_AMT
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.DIFF_UTGST_AMT
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.DIFF_IGST_AMT
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.DIFF_CCESS_AMT
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.CustSuppMaterial_diff
                        strFormula = strFormula & "," & pintDecimal & ")"
                    Else
                        strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.BasicDiff
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.DIFF_CGST_AMT
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.DIFF_SGST_AMT
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.DIFF_UTGST_AMT
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.DIFF_IGST_AMT
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.DIFF_CCESS_AMT
                        strFormula = strFormula & "," & pintDecimal & ")"
                    End If
                Else
                    If blnAddCustMatrl = True Then
                        strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.BasicDiff & " + R" & intLoopCounter & "C"
                        strFormula = strFormula & enumPreInvoiceDetails.SalesTaxValueDiff & "+ R" & intLoopCounter & "C" & enumPreInvoiceDetails.SSTVlaueDiff
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.TotalExciseValueDiff
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.TotalEcssDiff
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.TotalsEcssDiff
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.CustSuppMaterial_diff & "," & pintDecimal & ")"
                    Else
                        strFormula = "Round(R" & intLoopCounter & "C" & enumPreInvoiceDetails.BasicDiff & " + R" & intLoopCounter & "C"
                        strFormula = strFormula & enumPreInvoiceDetails.SalesTaxValueDiff
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.TotalEcssDiff
                        strFormula = strFormula & " + R" & intLoopCounter & "C" & enumPreInvoiceDetails.TotalsEcssDiff
                        strFormula = strFormula & "+ R" & intLoopCounter & "C"
                        strFormula = strFormula & enumPreInvoiceDetails.SSTVlaueDiff & " + R" & intLoopCounter
                        strFormula = strFormula & "C" & enumPreInvoiceDetails.TotalExciseValueDiff & "," & pintDecimal & ")"
                    End If
                End If
                .Col = enumPreInvoiceDetails.TotalInvoiceValue
                .Col2 = enumPreInvoiceDetails.TotalInvoiceValue
                .BlockMode = True
                .Formula = strFormula
                .Action = FPSpreadADO.ActionConstants.ActionReCalc
                .BlockMode = False
                .Lock = False
            Next
        End With
        rsParameterData.ResultSetClose()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSADCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSADCode.TextChanged
        On Error GoTo ErrHandler
        If Len(txtSADCode.Text) = 0 Then
            lblSAD_per.Text = "0.00"
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSADCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSADCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If Len(txtSADCode.Text) > 0 Then
                            Call txtSADCode_Validating(txtSADCode, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            txtExciseTaxType.Focus()
                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtSADCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSADCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If cmdSADCodeHelp.Enabled Then Call cmdSADCodeHelp_Click(cmdSADCodeHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSADCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtSADCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        If Len(txtSADCode.Text) > 0 Then
            If CheckExistanceOfFieldData((txtSADCode.Text), "TxRt_Rate_No", "Gen_TaxRate", " ( Unit_code='" & gstrUNITID & "' and Tx_TaxeID='SAD') and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))") Then
                lblSAD_per.Text = CStr(GetTaxRate((txtSADCode.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Unit_code= '" & gstrUNITID & "' and Tx_TaxeID='SAD') and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"))
                If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    Call SetFormulaofColumns(0, 4)
                    ToShowDatainSummeryGrid()
                End If
                If txtExciseTaxType.Enabled Then txtExciseTaxType.Focus()
            Else
                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtSADCode.Text = ""
                If txtSADCode.Enabled Then txtSADCode.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtSaleTaxType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSaleTaxType.TextChanged
        On Error GoTo ErrHandler
        If Len(txtSaleTaxType.Text) = 0 Then
            lblSaltax_Per.Text = "0.00"
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSaleTaxType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSaleTaxType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If Len(txtSaleTaxType.Text) > 0 Then
                            Call txtSaleTaxType_Validating(txtSaleTaxType, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            txtSurchargeTaxType.Focus()
                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtSaleTaxType_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSaleTaxType.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdSaleTaxType.Enabled Then Call CmdSaleTaxType_Click(CmdSaleTaxType, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSaleTaxType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtSaleTaxType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        If Len(txtSaleTaxType.Text) > 0 Then
            If CheckExistanceOfFieldData((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "Unit_code='" & gstrUNITID & "' and (Tx_TaxeID='CST' OR Tx_TaxeID='LST' OR Tx_TaxeID='VAT') and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))") Then
                lblSaltax_Per.Text = CStr(GetTaxRate((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Unit_code= '" & gstrUNITID & "' and (Tx_TaxeID='CST' OR Tx_TaxeID='LST' or Tx_TaxeID='VAT') and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"))
                If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    Call SetFormulaofColumns(0, 4)
                    ToShowDatainSummeryGrid()
                End If
                If txtSurchargeTaxType.Enabled Then txtSurchargeTaxType.Focus()
            Else
                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtSaleTaxType.Text = ""
                If txtSaleTaxType.Enabled Then txtSaleTaxType.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtSEcssCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSEcssCode.TextChanged
        On Error GoTo ErrHandler
        If Trim(txtSEcssCode.Text) = "" Then
            lblSEcssCode.Text = "0.00"
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtSEcssCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSEcssCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            cmdSEcssCodeHelp.PerformClick()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSEcssCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSEcssCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        Call txtSEcssCode_Validating(txtSEcssCode, New System.ComponentModel.CancelEventArgs(False))
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtSEcssCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtSEcssCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        If Trim(txtSEcssCode.Text) <> "" Then
            If CheckExistanceOfFieldData(Trim(txtSEcssCode.Text), "TxRt_Rate_No", "Gen_TaxRate", "Unit_code='" & gstrUNITID & "' and Tx_TaxeID='ECSSH' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))") Then
                lblSEcssCode.Text = CStr(GetTaxRate(Trim(txtSEcssCode.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", "Unit_code= '" & gstrUNITID & "' and Tx_TaxeID='ECSSH' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"))
                If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    Call SetFormulaofColumns(0, 4)
                    Call ToShowDatainSummeryGrid()
                End If
                If txtMRP.Enabled Then txtMRP.Focus()
            Else
                MsgBox("Invalid SHEcess Tax Code!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                Cancel = True
                txtSEcssCode.Text = ""
                If txtSEcssCode.Enabled Then txtSEcssCode.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtSurchargeTaxType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSurchargeTaxType.TextChanged
        On Error GoTo ErrHandler
        If Trim(txtSurchargeTaxType.Text) = "" Then
            lblSurcharge_Per.Text = "0.00"
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSurchargeTaxType_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSurchargeTaxType.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call cmdSurchargeTaxCode_Click(cmdSurchargeTaxCode, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSurchargeTaxType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSurchargeTaxType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If Me.txtEcssCode.Enabled = True Then
                            Me.txtEcssCode.Focus()
                        Else
                            With Me.spdPrevInv
                                .Row = 1
                                .Col = 1
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                            End With
                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtSurchargeTaxType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtSurchargeTaxType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        If Trim(txtSurchargeTaxType.Text) <> "" Then
            If CheckExistanceOfFieldData((txtSurchargeTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " Unit_code='" & gstrUNITID & "' and Tx_TaxeID='SST' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))") Then
                lblSurcharge_Per.Text = CStr(GetTaxRate((txtSurchargeTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Unit_code= '" & gstrUNITID & "' and Tx_TaxeID='SST' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"))
                If spdPrevInv.Enabled Then
                    With Me.spdInvDetails
                        If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                            Call SetFormulaofColumns(0, 4)
                            ToShowDatainSummeryGrid()
                        End If
                        .Row = 1
                        .Col = 1
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                    End With
                End If
            Else
                MsgBox("This Surcharge Tax Type is not correct for help press F1.", MsgBoxStyle.Information, "eMpro")
                Cancel = True
                txtSurchargeTaxType.Text = ""
                If txtSurchargeTaxType.Enabled Then txtSurchargeTaxType.Focus()
            End If
        Else
            If spdPrevInv.Enabled Then
                With Me.spdInvDetails
                    .Row = 1
                    .Col = 1
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                End With
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Public Function ToCheckForNegativeValues(ByRef pintRow As Integer) As Boolean
        Dim intMaxLoop As Short
        Dim intLoopCounter As Short
        On Error GoTo ErrHandler
        ToCheckForNegativeValues = False
        If pintRow = 0 Then
            pintRow = 1
            intMaxLoop = spdPrevInv.MaxRows
        Else
            intMaxLoop = pintRow
        End If
        For intLoopCounter = pintRow To intMaxLoop
            With spdPrevInv
                ''--- Checking For Basic Values              
                .Row = pintRow
                .Col = enumPreInvoiceDetails.BasicDiff
                If Val(.Text) < 0 Then
                    MsgBox("Negative Basic Value Can Not Select Invoice.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    ToCheckForNegativeValues = False
                    .Row = pintRow
                    .Col = enumPreInvoiceDetails.Select_Invoice
                    .Value = System.Windows.Forms.CheckState.Unchecked
                    .Row = pintRow
                    .Col = enumPreInvoiceDetails.New_Rate
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    .Focus()
                    Exit Function
                End If
                ''--- Checking For Accessable Values
                .Row = pintRow
                .Col = enumPreInvoiceDetails.AccessableValue_Diff
                If Val(.Text) < 0 Then
                    MsgBox("Negative Accessable Value Can Not Select Invoice.")
                    ToCheckForNegativeValues = False
                    .Row = pintRow
                    .Col = enumPreInvoiceDetails.Select_Invoice
                    .Value = System.Windows.Forms.CheckState.Unchecked
                    .Row = pintRow
                    .Col = enumPreInvoiceDetails.New_Rate
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                    Exit Function
                End If
            End With
        Next
        ToCheckForNegativeValues = True
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Sub ToShowDatainSummeryGrid()
        Dim dblNewRate As Double
        Dim dblNewPacking As Double
        Dim dblNewCustSuppMat As Double
        Dim dblNewToolCost As Double
        Dim arrdblSummaryInfo() As Object
        Dim intLoopCounter As Short
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.WaitCursor)
        With spdPrevInv
            .Row = 1
            .Col = enumPreInvoiceDetails.New_Rate
            If .Text = "" Then
            Else
                dblNewRate = CDbl(.Text)
            End If
            .Col = enumPreInvoiceDetails.NewPacking
            If .Text = "" Then
            Else
                dblNewPacking = CDbl(.Text)
            End If
            .Col = enumPreInvoiceDetails.NewCustSuppMaterial
            If .Text = "" Then
            Else
                dblNewCustSuppMat = CDbl(.Text)
            End If
            .Col = enumPreInvoiceDetails.newToolCost
            If .Text = "" Then
            Else
                dblNewToolCost = CDbl(.Text)
            End If
        End With
        objInvoiceCls = New prj_InvoiceCalc.clsInvoiceCalculation(gstrUNITID)
        If bCheck = True Then
            If blnInclude = True Then
                If strSelInvoices = "" Then
                    FillInvoiceNumber()
                    For intLoopCounter = 0 To lstInv.Items.Count - 1
                        strSelInvoices = strSelInvoices & "|" & lstInv.Items.Item(intLoopCounter).Text
                    Next intLoopCounter
                    If VB.Left(strSelInvoices, 1) = "|" Then
                        strSelInvoices = VB.Right(strSelInvoices, Len(strSelInvoices) - 1)
                    End If
                End If
                bCheck = False
            Else
                bCheck = True
            End If
        End If
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then Exit Sub
        spdPrevInv.Lock = False
        objInvoiceCls.FormMode = CmdGrpChEnt.Mode
        mstrSOType.Value = GetSOType()
        Call ChangeInvTypeCaption(mstrSOType.Value)
        '101188073
        If gblnGSTUnit Then
            blnInclude = objInvoiceCls.GetInvoiceSummary(mP_Connection, strSelInvoices, bCheck, mstrSOType.Value, Val(txtMRP.Text), arrdblSummaryInfo, dblNewRate, dblNewPacking, dblNewCustSuppMat, dblNewToolCost, lblExctax_Per.Text, lblCVD_Per.Text, lblSAD_per.Text, lblEcssCode.Text, lblSaltax_Per.Text, lblSurcharge_Per.Text, getDateForDB(dtpDateFrom.Value), getDateForDB(dtpDateTo.Value), txtCustCode.Text, lblItemCode.Text, txtCustPartCode.Text, txtChallanNo.Text, lblSEcssCode.Text, lblCGSTPercent.Text, lblSGSTPercent.Text, lblUTGSTPercent.Text, lblIGSTPercent.Text, lblCompCessPercent.Text)
        Else
            blnInclude = objInvoiceCls.GetInvoiceSummary(mP_Connection, strSelInvoices, bCheck, mstrSOType.Value, Val(txtMRP.Text), arrdblSummaryInfo, dblNewRate, dblNewPacking, dblNewCustSuppMat, dblNewToolCost, lblExctax_Per.Text, lblCVD_Per.Text, lblSAD_per.Text, lblEcssCode.Text, lblSaltax_Per.Text, lblSurcharge_Per.Text, getDateForDB(dtpDateFrom.Value), getDateForDB(dtpDateTo.Value), txtCustCode.Text, lblItemCode.Text, txtCustPartCode.Text, txtChallanNo.Text, lblSEcssCode.Text)
        End If
        '101188073
        If blnInclude = False Then
            cmdSelectInvoice_Click(cmdSelectInvoice, New System.EventArgs())
            Exit Sub
        End If
        With spdInvDetails
            .MaxRows = 1
            If LenOfArray(arrdblSummaryInfo) <> 0 Then
                If dblNewRate <= 0 Then
                    Call .SetText(enumInvoiceSummery.Rate, 1, 0.0#)
                    Call .SetText(enumInvoiceSummery.BasicValue, 1, 0.0#)
                    Call .SetText(enumInvoiceSummery.CustSuppMat, 1, 0.0#)
                    Call .SetText(enumInvoiceSummery.ToolCost, 1, 0.0#)
                    Call .SetText(enumInvoiceSummery.AccessableValue, 1, 0.0#)
                    Call .SetText(enumInvoiceSummery.ExciseValue, 1, 0.0#)
                    Call .SetText(enumInvoiceSummery.EcssValue, 1, 0.0#)
                    Call .SetText(enumInvoiceSummery.sEcssValue, 1, 0.0#)
                    Call .SetText(enumInvoiceSummery.PackingValue, 1, 0.0#)
                    Call .SetText(enumInvoiceSummery.SalesTaxType, 1, 0.0#)
                    Call .SetText(enumInvoiceSummery.SSTType, 1, 0.0#)
                    Call .SetText(enumInvoiceSummery.CGST, 1, 0.0#)
                    Call .SetText(enumInvoiceSummery.SGST, 1, 0.0#)
                    Call .SetText(enumInvoiceSummery.UTGST, 1, 0.0#)
                    Call .SetText(enumInvoiceSummery.IGST, 1, 0.0#)
                    Call .SetText(enumInvoiceSummery.CCESS, 1, 0.0#)
                    Call .SetText(enumInvoiceSummery.SummeryInvoiceValue, 1, 0.0#)
                Else
                    Call .SetText(enumInvoiceSummery.Rate, 1, arrdblSummaryInfo(0))
                    Call .SetText(enumInvoiceSummery.BasicValue, 1, arrdblSummaryInfo(1))
                    Call .SetText(enumInvoiceSummery.CustSuppMat, 1, arrdblSummaryInfo(2))
                    Call .SetText(enumInvoiceSummery.ToolCost, 1, arrdblSummaryInfo(3))
                    Call .SetText(enumInvoiceSummery.AccessableValue, 1, arrdblSummaryInfo(4))
                    If gblnGSTUnit Then
                        Call .SetText(enumInvoiceSummery.ExciseValue, 1, 0)
                        Call .SetText(enumInvoiceSummery.EcssValue, 1, 0)
                        Call .SetText(enumInvoiceSummery.sEcssValue, 1, 0)
                        Call .SetText(enumInvoiceSummery.SalesTaxType, 1, 0)
                        Call .SetText(enumInvoiceSummery.SSTType, 1, 0)
                        Call .SetText(enumInvoiceSummery.CGST, 1, arrdblSummaryInfo(12))
                        Call .SetText(enumInvoiceSummery.SGST, 1, arrdblSummaryInfo(13))
                        Call .SetText(enumInvoiceSummery.UTGST, 1, arrdblSummaryInfo(14))
                        Call .SetText(enumInvoiceSummery.IGST, 1, arrdblSummaryInfo(15))
                        Call .SetText(enumInvoiceSummery.CCESS, 1, arrdblSummaryInfo(16))
                    Else
                        Call .SetText(enumInvoiceSummery.ExciseValue, 1, arrdblSummaryInfo(5))
                        Call .SetText(enumInvoiceSummery.EcssValue, 1, arrdblSummaryInfo(6))
                        Call .SetText(enumInvoiceSummery.sEcssValue, 1, arrdblSummaryInfo(10))
                        Call .SetText(enumInvoiceSummery.SalesTaxType, 1, arrdblSummaryInfo(7))
                        Call .SetText(enumInvoiceSummery.SSTType, 1, arrdblSummaryInfo(8))
                        Call .SetText(enumInvoiceSummery.CGST, 1, 0)
                        Call .SetText(enumInvoiceSummery.SGST, 1, 0)
                        Call .SetText(enumInvoiceSummery.UTGST, 1, 0)
                        Call .SetText(enumInvoiceSummery.IGST, 1, 0)
                        Call .SetText(enumInvoiceSummery.CCESS, 1, 0)
                    End If
                    Call .SetText(enumInvoiceSummery.PackingValue, 1, Val(arrdblSummaryInfo(11)))
                    Call .SetText(enumInvoiceSummery.SummeryInvoiceValue, 1, arrdblSummaryInfo(9))
                End If
            End If
        End With
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Sub ToShowDatainSummeryGrid_Old()
        Dim intMaxLoop As Short
        Dim intLoopCounter As Short
        Dim dblRate As Double
        Dim dblCustSuppMat As Double
        Dim dblToolCost As Double
        Dim dblBasic As Double
        Dim dblAccessableValue As Double
        Dim dblCVDValue As Double
        Dim dblSVDValue As Double
        Dim dblExciseOnly As Double
        Dim dblExciseValue As Double
        Dim dblEcssValue As Double
        Dim dblSAlesTaxValue As Double
        Dim dblSSTax As Double
        Dim dblSummeryInvValue As Double
        On Error GoTo ErrHandler
        dblRate = 0
        dblCustSuppMat = 0
        dblToolCost = 0
        dblBasic = 0
        dblAccessableValue = 0
        dblExciseValue = 0
        dblSAlesTaxValue = 0
        dblSSTax = 0
        dblSummeryInvValue = 0
        For intLoopCounter = 1 To spdPrevInv.MaxRows
            spdPrevInv.Col = enumPreInvoiceDetails.Select_Invoice
            spdPrevInv.Row = intLoopCounter
            If spdPrevInv.Value = System.Windows.Forms.CheckState.Checked Then
                spdPrevInv.Row = intLoopCounter
                spdPrevInv.Col = enumPreInvoiceDetails.Rate_diff
                dblRate = dblRate + Val(spdPrevInv.Text)
                spdPrevInv.Row = intLoopCounter
                spdPrevInv.Col = enumPreInvoiceDetails.CustSuppMaterial_diff
                dblCustSuppMat = dblCustSuppMat + Val(spdPrevInv.Text)
                spdPrevInv.Row = intLoopCounter
                spdPrevInv.Col = enumPreInvoiceDetails.ToolCost_diff
                dblToolCost = dblToolCost + Val(spdPrevInv.Text)
                spdPrevInv.Row = intLoopCounter
                spdPrevInv.Col = enumPreInvoiceDetails.BasicDiff
                dblBasic = dblBasic + Val(spdPrevInv.Text)
                spdPrevInv.Row = intLoopCounter
                spdPrevInv.Col = enumPreInvoiceDetails.AccessableValue_Diff
                dblAccessableValue = dblAccessableValue + Val(spdPrevInv.Text)
                spdPrevInv.Row = intLoopCounter
                spdPrevInv.Col = enumPreInvoiceDetails.TotalExciseValueDiff
                dblExciseValue = dblExciseValue + Val(spdPrevInv.Text)
                spdPrevInv.Row = intLoopCounter
                spdPrevInv.Col = enumPreInvoiceDetails.TotalEcssDiff
                dblEcssValue = dblEcssValue + Val(spdPrevInv.Text)
                spdPrevInv.Row = intLoopCounter
                spdPrevInv.Col = enumPreInvoiceDetails.SalesTaxValueDiff
                dblSAlesTaxValue = dblSAlesTaxValue + Val(spdPrevInv.Text)
                spdPrevInv.Row = intLoopCounter
                spdPrevInv.Col = enumPreInvoiceDetails.SSTVlaueDiff
                dblSSTax = dblSSTax + Val(spdPrevInv.Text)
                spdPrevInv.Row = intLoopCounter
                spdPrevInv.Col = enumPreInvoiceDetails.TotalInvoiceValue
                dblSummeryInvValue = dblSummeryInvValue + Val(spdPrevInv.Text)
            End If
        Next
        With spdInvDetails
            .MaxRows = 1
            Call .SetText(enumInvoiceSummery.Rate, 1, dblRate)
            Call .SetText(enumInvoiceSummery.BasicValue, 1, dblBasic)
            Call .SetText(enumInvoiceSummery.CustSuppMat, 1, dblCustSuppMat)
            Call .SetText(enumInvoiceSummery.ToolCost, 1, dblToolCost)
            Call .SetText(enumInvoiceSummery.AccessableValue, 1, dblAccessableValue)
            Call .SetText(enumInvoiceSummery.ExciseValue, 1, dblExciseValue)
            Call .SetText(enumInvoiceSummery.EcssValue, 1, dblEcssValue)
            Call .SetText(enumInvoiceSummery.SalesTaxType, 1, dblSAlesTaxValue)
            Call .SetText(enumInvoiceSummery.SSTType, 1, dblSSTax)
            Call .SetText(enumInvoiceSummery.SummeryInvoiceValue, 1, dblSummeryInvValue)
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function ValidatebeforeSave() As Boolean
        On Error GoTo ErrHandler
        Dim lstrControls As String
        Dim blnNegValCheck As Boolean
        Dim lNo As Integer
        Dim lctrFocus As System.Windows.Forms.Control
        ValidatebeforeSave = True
        blnNegValCheck = False
        lNo = 1
        lstrControls = ResolveResString(10059)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            Call SetFormulaofColumns(0, 4)
            bCheck = True
            ToShowDatainSummeryGrid()
        End If
        If dtpDateTo.Value < dtpDateFrom.Value Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Date To < Date From."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = Me.dtpDateFrom
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Me.txtLocationCode.Text) = 0) Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Location Code."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = Me.txtLocationCode
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Me.txtCustCode.Text) = 0) Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Customer Code."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = Me.txtCustCode
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Me.txtCustPartCode.Text) = 0) Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Customer Part Code"
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = Me.txtCustCode
            End If
            ValidatebeforeSave = False
        End If
        With spdInvDetails
            .Row = 1
            .Col = enumInvoiceSummery.Rate
            If Val(.Text) < 0 Then
                blnNegValCheck = True
            End If
        End With
        With spdInvDetails
            .Row = 1
            .Col = enumInvoiceSummery.BasicValue
            If Val(.Text) < 0 Then
                blnNegValCheck = True
            End If
        End With
        With spdInvDetails
            .Row = 1
            .Col = enumInvoiceSummery.CustSuppMat
            If Val(.Text) < 0 Then
                blnNegValCheck = True
            End If
        End With
        With spdInvDetails
            .Row = 1
            .Col = enumInvoiceSummery.ExciseValue
            If Val(.Text) < 0 Then
                blnNegValCheck = True
            End If
        End With
        With spdInvDetails
            .Row = 1
            .Col = enumInvoiceSummery.EcssValue
            If Val(.Text) < 0 Then
                blnNegValCheck = True
            End If
        End With
        With spdInvDetails
            .Row = 1
            .Col = enumInvoiceSummery.sEcssValue
            If Val(.Text) < 0 Then
                blnNegValCheck = True
            End If
        End With
        With spdInvDetails
            .Row = 1
            .Col = enumInvoiceSummery.PackingValue
            If Val(.Text) < 0 Then
                blnNegValCheck = True
            End If
        End With
        With spdInvDetails
            .Row = 1
            .Col = enumInvoiceSummery.SalesTaxType
            If Val(.Text) < 0 Then
                blnNegValCheck = True
            End If
        End With
        With spdInvDetails
            .Row = 1
            .Col = enumInvoiceSummery.SSTType
            If Val(.Text) < 0 Then
                blnNegValCheck = True
            End If
        End With
        With spdInvDetails
            .Row = 1
            .Col = enumInvoiceSummery.AccessableValue
            If Val(.Text) < 0 Then
                blnNegValCheck = True
            End If
        End With
        With spdInvDetails
            .Row = 1
            .Col = enumInvoiceSummery.ToolCost
            If Val(.Text) < 0 Then
                blnNegValCheck = True
            End If
        End With
        '101188073
        If gblnGSTUnit Then
            With spdInvDetails
                .Row = 1
                .Col = enumInvoiceSummery.CGST
                If Val(.Text) < 0 Then
                    blnNegValCheck = True
                End If
            End With
            With spdInvDetails
                .Row = 1
                .Col = enumInvoiceSummery.SGST
                If Val(.Text) < 0 Then
                    blnNegValCheck = True
                End If
            End With
            With spdInvDetails
                .Row = 1
                .Col = enumInvoiceSummery.UTGST
                If Val(.Text) < 0 Then
                    blnNegValCheck = True
                End If
            End With
            With spdInvDetails
                .Row = 1
                .Col = enumInvoiceSummery.IGST
                If Val(.Text) < 0 Then
                    blnNegValCheck = True
                End If
            End With
            With spdInvDetails
                .Row = 1
                .Col = enumInvoiceSummery.CCESS
                If Val(.Text) < 0 Then
                    blnNegValCheck = True
                End If
            End With
        End If
        '101188073
        With spdInvDetails
            .Row = 1
            .Col = enumInvoiceSummery.SummeryInvoiceValue
            If Val(.Text) < 0 Then
                blnNegValCheck = True
            End If
        End With
        If blnNegValCheck = True Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Please check for negative values."
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = Me.spdInvDetails
            End If
            ValidatebeforeSave = False
        End If
        With spdInvDetails
            .Row = 1
            .Col = enumInvoiceSummery.SummeryInvoiceValue
            If Val(.Text) <= 0 Then
                lstrControls = lstrControls & vbCrLf & lNo & ". Total Invoice Value can Not Be Zero"
                lNo = lNo + 1
                If lctrFocus Is Nothing Then
                    lctrFocus = Me.spdInvDetails
                End If
                ValidatebeforeSave = False
            End If
        End With
        If Len(Trim(txtCVDCode.Text)) = 0 Then
            lblCVD_Per.Text = "0.00"
        End If
        If Len(Trim(txtSADCode.Text)) = 0 Then
            lblSAD_per.Text = "0.00"
        End If
        If Len(Trim(txtExciseTaxType.Text)) = 0 Then
            lblExctax_Per.Text = "0.00"
        End If
        If Len(Trim(txtSaleTaxType.Text)) = 0 Then
            lblSaltax_Per.Text = "0.00"
        End If
        If Len(Trim(txtSurchargeTaxType.Text)) = 0 Then
            lblSurcharge_Per.Text = "0.00"
        End If
        '101188073
        If gblnGSTUnit Then
            If Len(Trim(txtCGSTType.Text)) = 0 Then
                lblCGSTPercent.Text = "0.00"
            End If
            If Len(Trim(txtSGSTType.Text)) = 0 Then
                lblSGSTPercent.Text = "0.00"
            End If
            If Len(Trim(txtUTGSTType.Text)) = 0 Then
                lblUTGSTPercent.Text = "0.00"
            End If
            If Len(Trim(txtIGSTType.Text)) = 0 Then
                lblIGSTPercent.Text = "0.00"
            End If
            If Len(Trim(txtCompCessType.Text)) = 0 Then
                lblCompCessPercent.Text = "0.00"
            End If
        End If
        '101188073
        If (Len(lblCurrencyDes.Text) = 0) Then
            lblCurrencyDes.Text = gstrCURRENCYCODE
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Arrow)
        If Not ValidatebeforeSave Then
            MsgBox(lstrControls, MsgBoxStyle.Information, ResolveResString(10059))
            If UCase(lctrFocus.Name) = "SPDINVDETAILS" Then
                sstbInvoiceDtl.SelectedIndex = 1
                lctrFocus.Focus()
            Else
                lctrFocus.Focus()
            End If
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
    End Function
    Public Sub SaveData()
        On Error GoTo ErrHandler
        Dim strsqlHdr As String
        Dim strSqlDtl As String
        Dim strSqlCreditAdvDtl As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim dblSummeryRate As Double
        Dim dblSummeryBasic_Amount As Double
        Dim dblSummeryAccessible_amount As Double
        Dim dblSummeryTotalExciseAmount As Double
        Dim dblSummeryCustMtrl_Amount As Double
        Dim dblSummeryToolCost_amount As Double
        Dim dblSummerySales_Tax_Amount As Double
        Dim dblSummerySurcharge_Sales_Tax_Amount As Double
        Dim dblSummeryEcss_amount As Double
        Dim dblSummerySEcss_amount As Double
        Dim dblSummerytotal_amount As Double
        Dim dblSummeryPacking_amount As Double
        Dim dblSuppInvoiceNo As Double
        Dim strSuppInvDate As String
        Dim strUpdateSalesDtl As String
        Dim strDeleteSuppDtl As String
        Dim strDeleteSuppCrAdvise As String
        Dim varFlag As Object
        Dim dblNewRate As Double
        Dim dblNewPacking As Double
        Dim dblNewCustSuppMat As Double
        Dim dblNewToolCost As Double
        Dim rsParameterData As New ClsResultSetDB
        Dim blnISInsExcisable As Boolean
        Dim blnEOUFlag As Boolean
        Dim blnISExciseRoundOff As Boolean
        Dim blnISSalesTaxRoundOff As Boolean
        Dim blnISSurChargeTaxRoundOff As Boolean
        Dim blnAddCustMatrl As Boolean
        Dim blnISBasicRoundOff As Boolean
        Dim blnTotalToolCostRoundOff As Boolean
        Dim ldblExciseValueForSaleTax As Double
        Dim blnInsIncSTax As Boolean
        Dim blnTCSTax As Boolean
        Dim VarDelete As Object
        Dim intNonDeletedRowCount As Short
        Dim intBasicRoundOffDecimal As Short
        Dim intSaleTaxRoundOffDecimal As Short
        Dim intExciseRoundOffDecimal As Short
        Dim intSSTRoundOffDecimal As Short
        Dim intTCSRoundOffDecimal As Short
        Dim intToolCostRoundOffDecimal As Short
        Dim blnActiveTrans As Boolean
        Dim blnECSSTax As Boolean
        Dim intECSRoundOffDecimal As Short
        Dim ldblTotalECSSTaxAmount As Double
        Dim blnECSSOnSaleTax As Boolean
        Dim intECSSOnSaleRoundOffDecimal As Short
        Dim ldblTotalECSSOnSaleTaxAmount As Double
        Dim blnTurnOverTax As Boolean
        Dim intTurnOverTaxRoundOffDecimal As Short
        Dim ldblTotalTurnOverTaxAmount As Double
        Dim blnTotalInvoiceAmount As Boolean
        Dim intTotalInvoiceAmountRoundOffDecimal As Short
        Dim ldblTotalInvoiceValueRoundOff As Double
        Dim blnEcssonCVD As Boolean
        Dim intEcssOnCVDRoundOff As Short
        Dim strParamQuery As String
        Dim dblTotalInvoiceAmtRoundOff_diff As Double
        Dim dblTotalInvoiceAmt As Double
        Dim blnGSTRoundOff As Boolean
        Dim intGSTRoundOffDecimal As Integer
        Dim dblCGSTAmount As Double = 0
        Dim dblSGSTAmount As Double = 0
        Dim dblUTGSTAmount As Double = 0
        Dim dblIGSTAmount As Double = 0
        Dim dblCCESSAmount As Double = 0
        dblSummeryRate = 0 : dblSummeryBasic_Amount = 0 : dblSummeryAccessible_amount = 0 : dblSummeryTotalExciseAmount = 0
        dblSummeryCustMtrl_Amount = 0 : dblSummeryToolCost_amount = 0 : dblSummerySales_Tax_Amount = 0 : dblSummerySurcharge_Sales_Tax_Amount = 0
        dblSummeryEcss_amount = 0
        With spdInvDetails
            Call .GetFloat(enumInvoiceSummery.Rate, 1, dblSummeryRate)
            Call .GetFloat(enumInvoiceSummery.BasicValue, 1, dblSummeryBasic_Amount)
            Call .GetFloat(enumInvoiceSummery.CustSuppMat, 1, dblSummeryCustMtrl_Amount)
            Call .GetFloat(enumInvoiceSummery.ToolCost, 1, dblSummeryToolCost_amount)
            Call .GetFloat(enumInvoiceSummery.AccessableValue, 1, dblSummeryAccessible_amount)
            Call .GetFloat(enumInvoiceSummery.ExciseValue, 1, dblSummeryTotalExciseAmount)
            Call .GetFloat(enumInvoiceSummery.EcssValue, 1, dblSummeryEcss_amount)
            Call .GetFloat(enumInvoiceSummery.sEcssValue, 1, dblSummerySEcss_amount)
            .Row = 1
            .Col = enumInvoiceSummery.PackingValue
            dblSummeryPacking_amount = CDbl(.Text)
            Call .GetFloat(enumInvoiceSummery.SalesTaxType, 1, dblSummerySales_Tax_Amount)
            Call .GetFloat(enumInvoiceSummery.SSTType, 1, dblSummerySurcharge_Sales_Tax_Amount)
            Call .GetFloat(enumInvoiceSummery.CGST, 1, dblCGSTAmount)
            Call .GetFloat(enumInvoiceSummery.SGST, 1, dblSGSTAmount)
            Call .GetFloat(enumInvoiceSummery.UTGST, 1, dblUTGSTAmount)
            Call .GetFloat(enumInvoiceSummery.IGST, 1, dblIGSTAmount)
            Call .GetFloat(enumInvoiceSummery.CCESS, 1, dblCCESSAmount)
            Call .GetFloat(enumInvoiceSummery.SummeryInvoiceValue, 1, dblSummerytotal_amount)
        End With

        'Total GST Amount RoundOff
        If blnGSTRoundOff = False Then
            dblCGSTAmount = System.Math.Round(dblCGSTAmount, intGSTRoundOffDecimal)
            dblSGSTAmount = System.Math.Round(dblSGSTAmount, intGSTRoundOffDecimal)
            dblUTGSTAmount = System.Math.Round(dblUTGSTAmount, intGSTRoundOffDecimal)
            dblIGSTAmount = System.Math.Round(dblIGSTAmount, intGSTRoundOffDecimal)
            dblCCESSAmount = System.Math.Round(dblCCESSAmount, intGSTRoundOffDecimal)
        Else
            dblCGSTAmount = System.Math.Round(dblCGSTAmount, 0)
            dblSGSTAmount = System.Math.Round(dblSGSTAmount, 0)
            dblUTGSTAmount = System.Math.Round(dblUTGSTAmount, 0)
            dblIGSTAmount = System.Math.Round(dblIGSTAmount, 0)
            dblCCESSAmount = System.Math.Round(dblCCESSAmount, 0)
        End If

        If gblnGSTUnit Then
            dblTotalInvoiceAmt = Val(CStr(dblSummeryPacking_amount + dblSummeryBasic_Amount + dblCGSTAmount + dblSGSTAmount + dblUTGSTAmount + dblIGSTAmount + dblCCESSAmount))
        Else
            dblTotalInvoiceAmt = Val(CStr(dblSummeryPacking_amount + dblSummeryBasic_Amount + dblSummeryTotalExciseAmount + dblSummeryEcss_amount + dblSummerySEcss_amount + dblSummerySales_Tax_Amount + dblSummerySurcharge_Sales_Tax_Amount))
        End If
        With spdPrevInv
            .Row = 1
            .Col = enumPreInvoiceDetails.New_Rate
            If .Text = "" Then
            Else
                dblNewRate = CDbl(.Text)
            End If
            .Col = enumPreInvoiceDetails.NewPacking
            If .Text = "" Then
            Else
                dblNewPacking = CDbl(.Text)
            End If
            .Col = enumPreInvoiceDetails.NewCustSuppMaterial
            If .Text = "" Then
            Else
                dblNewCustSuppMat = CDbl(.Text)
            End If
            .Col = enumPreInvoiceDetails.newToolCost
            If .Text = "" Then
            Else
                dblNewToolCost = CDbl(.Text)
            End If
        End With
        strParamQuery = "SELECT InsExc_Excise,CustSupp_Inc,EOU_Flag, Basic_Roundoff, Basic_Roundoff_decimal, SalesTax_Roundoff, SalesTax_Roundoff_decimal, Excise_Roundoff, Excise_Roundoff_decimal, "
        strParamQuery = strParamQuery & "ECESSOnCVD_Roundoff=isnull(ECESSOnCVD_Roundoff,0),ECESSOnCVDRoundOff_Decimal= isnull(ECESSOnCVDRoundOff_Decimal,0),"
        strParamQuery = strParamQuery & " SST_Roundoff, SST_Roundoff_decimal, InsInc_SalesTax, TCSTax_Roundoff, TCSTax_Roundoff_decimal, TotalToolCostRoundoff, TotalToolCostRoundoff_Decimal, ECESS_Roundoff, ECESSRoundoff_Decimal, ECESSOnSaleTax_Roundoff, ECESSOnSaleTaxRoundOff_Decimal, "
        strParamQuery = strParamQuery & " TurnOverTax_RoundOff, TurnOverTaxRoundOff_Decimal, TotalInvoiceAmount_RoundOff,TotalInvoiceAmountRoundOff_Decimal, SDTRoundOff, SDTRoundOff_Decimal,ISNULL(GSTTAX_ROUNDOFF_DECIMAL,0) GSTTAX_ROUNDOFF_DECIMAL,ISNULL(GSTTAX_ROUNDOFF,0) GSTTAX_ROUNDOFF FROM Sales_Parameter where Unit_code='" & gstrUnitId & "'"
        rsParameterData = New ClsResultSetDB
        rsParameterData.GetResult(strParamQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsParameterData.GetNoRows > 0 Then
            blnISInsExcisable = rsParameterData.GetValue("InsExc_Excise")
            blnEOUFlag = rsParameterData.GetValue("EOU_Flag")
            blnISBasicRoundOff = rsParameterData.GetValue("Basic_Roundoff")
            blnISExciseRoundOff = rsParameterData.GetValue("Excise_Roundoff")
            blnISSalesTaxRoundOff = rsParameterData.GetValue("SalesTax_Roundoff")
            blnISSurChargeTaxRoundOff = rsParameterData.GetValue("SST_Roundoff")
            blnAddCustMatrl = rsParameterData.GetValue("CustSupp_Inc")
            blnInsIncSTax = rsParameterData.GetValue("InsInc_SalesTax")
            blnTotalToolCostRoundOff = rsParameterData.GetValue("TotalToolCostRoundoff")
            blnTCSTax = rsParameterData.GetValue("TCSTax_Roundoff")
            intBasicRoundOffDecimal = rsParameterData.GetValue("Basic_Roundoff_decimal")
            intSaleTaxRoundOffDecimal = rsParameterData.GetValue("SalesTax_Roundoff_decimal")
            intExciseRoundOffDecimal = rsParameterData.GetValue("Excise_Roundoff_decimal")
            intSSTRoundOffDecimal = rsParameterData.GetValue("SST_Roundoff_decimal")
            intTCSRoundOffDecimal = rsParameterData.GetValue("TCSTax_Roundoff_decimal")
            intToolCostRoundOffDecimal = rsParameterData.GetValue("TotalToolCostRoundoff_decimal")
            If blnEOU_FLAG = True Then
                blnEcssonCVD = rsParameterData.GetValue("ECESSOnCVD_Roundoff")
                intEcssOnCVDRoundOff = rsParameterData.GetValue("ECESSOnCVDRoundOff_Decimal")
            End If
            blnECSSTax = rsParameterData.GetValue("ECESS_Roundoff")
            intECSRoundOffDecimal = rsParameterData.GetValue("ECESSRoundoff_Decimal")
            blnECSSOnSaleTax = rsParameterData.GetValue("ECESSOnSaleTax_Roundoff")
            intECSSOnSaleRoundOffDecimal = rsParameterData.GetValue("ECESSOnSaleTaxRoundOff_Decimal")
            blnTurnOverTax = rsParameterData.GetValue("TurnOverTax_RoundOff")
            intTurnOverTaxRoundOffDecimal = rsParameterData.GetValue("TurnOverTaxRoundOff_Decimal")
            blnTotalInvoiceAmount = rsParameterData.GetValue("TotalInvoiceAmount_RoundOff")
            intTotalInvoiceAmountRoundOffDecimal = rsParameterData.GetValue("TotalInvoiceAmountRoundOff_Decimal")
            blnGSTRoundOff = rsParameterData.GetValue("GSTTAX_ROUNDOFF")
            intGSTRoundOffDecimal = rsParameterData.GetValue("GSTTAX_ROUNDOFF_DECIMAL")
        Else
            MsgBox("No data define in Sales_Parameter Table", MsgBoxStyle.Critical, "empower")
            rsParameterData.ResultSetClose()
            rsParameterData = Nothing
            Exit Sub
        End If
        rsParameterData.ResultSetClose()
        rsParameterData = Nothing
        'Roundoff rate
        dblSummeryRate = System.Math.Round(dblSummeryRate, 4)
        'Basic Amount roundoff
        If blnISBasicRoundOff = False Then
            dblSummeryBasic_Amount = System.Math.Round(dblSummeryBasic_Amount, intBasicRoundOffDecimal)
        ElseIf blnISBasicRoundOff = True Then
            dblSummeryBasic_Amount = System.Math.Round(dblSummeryBasic_Amount, 0)
        End If
        'CustMtrl_Amount roundoff
        dblSummeryCustMtrl_Amount = System.Math.Round(dblSummeryCustMtrl_Amount, 4)
        'Tool cost roundoff
        If blnTotalToolCostRoundOff = False Then
            dblSummeryToolCost_amount = System.Math.Round(dblSummeryToolCost_amount, intToolCostRoundOffDecimal)
        ElseIf blnTotalToolCostRoundOff = True Then
            dblSummeryToolCost_amount = System.Math.Round(dblSummeryToolCost_amount, 0)
        End If
        'Accessiable Amount
        dblSummeryAccessible_amount = System.Math.Round(dblSummeryAccessible_amount, 4)
        'Excise Amount Roundoff
        If blnISExciseRoundOff = False Then
            dblSummeryTotalExciseAmount = System.Math.Round(dblSummeryTotalExciseAmount, intExciseRoundOffDecimal)
        ElseIf blnISExciseRoundOff = True Then
            dblSummeryTotalExciseAmount = System.Math.Round(dblSummeryTotalExciseAmount, 0)
        End If
        'ECESS Amount Roundoff
        If blnECSSTax = False Then
            dblSummeryEcss_amount = System.Math.Round(dblSummeryEcss_amount, intECSRoundOffDecimal)
            dblSummerySEcss_amount = System.Math.Round(dblSummerySEcss_amount, intECSRoundOffDecimal)
        ElseIf blnECSSTax = True Then
            dblSummeryEcss_amount = System.Math.Round(dblSummeryEcss_amount, 0)
            dblSummerySEcss_amount = System.Math.Round(dblSummerySEcss_amount, 0)
        End If
        'Sales Tax Roundoff
        If blnISSalesTaxRoundOff = False Then
            dblSummerySales_Tax_Amount = System.Math.Round(dblSummerySales_Tax_Amount, intSaleTaxRoundOffDecimal)
        ElseIf blnISSalesTaxRoundOff = True Then
            dblSummerySales_Tax_Amount = System.Math.Round(dblSummerySales_Tax_Amount, 0)
        End If
        'SurChargeTaxRoundOff Amount Round off
        If blnISSurChargeTaxRoundOff = False Then
            dblSummerySurcharge_Sales_Tax_Amount = System.Math.Round(dblSummerySurcharge_Sales_Tax_Amount, intSSTRoundOffDecimal)
        ElseIf blnISSurChargeTaxRoundOff = True Then
            dblSummerySurcharge_Sales_Tax_Amount = System.Math.Round(dblSummerySurcharge_Sales_Tax_Amount, 0)
        End If
        'Total Invoice Amount Roundoff
        If blnTotalInvoiceAmount = False Then
            dblSummerytotal_amount = System.Math.Round(dblSummerytotal_amount, intTotalInvoiceAmountRoundOffDecimal)
        ElseIf blnTotalInvoiceAmount = True Then
            dblSummerytotal_amount = System.Math.Round(dblSummerytotal_amount, 0)
        End If
      

        If blnTotalInvoiceAmount Then
            dblTotalInvoiceAmtRoundOff_diff = System.Math.Round(dblTotalInvoiceAmt - dblSummerytotal_amount, 4)
        Else
            dblTotalInvoiceAmtRoundOff_diff = (dblSummerytotal_amount - System.Math.Round(dblTotalInvoiceAmt, intTotalInvoiceAmountRoundOffDecimal))
        End If
        Select Case CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                strsqlHdr = "Insert into SupplementaryInv_hdr ("
                strsqlHdr = strsqlHdr & "Unit_code,TotalInvoiceAmtRoundOff_diff,Location_Code,Account_Code,Cust_name,Cust_Ref,Amendment_No,Doc_No,Invoice_DateFrom,Invoice_DateTo,Invoice_Date,Bill_Flag,Cancel_flag,Item_Code,"
                strsqlHdr = strsqlHdr & "Cust_Item_Code,Currency_Code,Rate,Basic_Amount,Accessible_amount,Excise_type,CVD_type,"
                strsqlHdr = strsqlHdr & "SAD_type,Excise_per,CVD_per,SVD_per,TotalExciseAmount,CustMtrl_Amount,"
                strsqlHdr = strsqlHdr & "ToolCost_amount,SalesTax_Type,SalesTax_Per,Sales_Tax_Amount,Surcharge_salesTaxType,Surcharge_SalesTax_Per,"
                strsqlHdr = strsqlHdr & "Surcharge_Sales_Tax_Amount,total_amount,"
                strsqlHdr = strsqlHdr & "SuppInv_Remarks,remarks,Ent_dt,Ent_UserId,Upd_dt,Upd_Userid"
                strsqlHdr = strsqlHdr & ",ECESS_Type,ECESS_Per,ECESS_Amount"
                strsqlHdr = strsqlHdr & ",SECESS_Type,SECESS_Per,SECESS_Amount,Packing,Packing_Amount,MRP"
                strsqlHdr = strsqlHdr & ",HSN_SAC_CODE,ISHSNORSAC,CGSTTXRT_TYPE,CGST_PERCENT,SGSTTXRT_TYPE,SGST_PERCENT,UTGSTTXRT_TYPE,UTGST_PERCENT,IGSTTXRT_TYPE,IGST_PERCENT,COMPENSATION_CESS_TYPE,COMPENSATION_CESS_PERCENT,CGST_AMT,SGST_AMT,UTGST_AMT,IGST_AMT,CCESS_AMT"
                strsqlHdr = strsqlHdr & ") Values"
                strsqlHdr = strsqlHdr & " ('" & gstrUNITID & "'," & dblTotalInvoiceAmtRoundOff_diff & ",'" & Trim(txtLocationCode.Text) & "','" & Trim(txtCustCode.Text) & "','" & Trim(lblCustCodeDes.Text) & "','"
                strsqlHdr = strsqlHdr & Trim(txtRefNo.Text) & "','" & Trim(txtAmendment.Text) & "'," & txtChallanNo.Text & ",'" & getDateForDB(dtpDateFrom.Value) & "','" & getDateForDB(dtpDateTo.Value) & "','" & getDateForDB(GetServerDate()) & "',0,0,'" & Trim(lblItemCode.Text)
                strsqlHdr = strsqlHdr & "','" & Trim(txtCustPartCode.Text) & "','" & Trim(lblCurrencyDes.Text) & "'," & dblSummeryRate & ","
                strsqlHdr = strsqlHdr & dblSummeryBasic_Amount & "," & dblSummeryAccessible_amount & ",'" & Trim(txtExciseTaxType.Text) & "','"
                strsqlHdr = strsqlHdr & Trim(txtCVDCode.Text) & "','" & Trim(txtSADCode.Text) & "'," & Trim(lblExctax_Per.Text) & ","
                strsqlHdr = strsqlHdr & Trim(lblCVD_Per.Text) & "," & Trim(lblSAD_per.Text) & "," & dblSummeryTotalExciseAmount & "," & dblSummeryCustMtrl_Amount
                strsqlHdr = strsqlHdr & "," & dblSummeryToolCost_amount & ",'" & Trim(txtSaleTaxType.Text) & "'," & lblSaltax_Per.Text & "," & dblSummerySales_Tax_Amount
                strsqlHdr = strsqlHdr & ",'" & Trim(txtSurchargeTaxType.Text) & "'," & lblSurcharge_Per.Text & "," & dblSummerySurcharge_Sales_Tax_Amount
                strsqlHdr = strsqlHdr & "," & dblSummerytotal_amount & ",'" & Trim(txtCustRefRemarks.Text) & "','" & Trim(txtRemarks.Text) & "','" & getDateForDB(GetServerDate()) & "','"
                strsqlHdr = strsqlHdr & mP_User & "','" & getDateForDB(GetServerDate()) & "','" & mP_User & "',"
                strsqlHdr = strsqlHdr & "'" & Trim(Me.txtEcssCode.Text) & "'," & Trim(Me.lblEcssCode.Text) & "," & Val(CStr(dblSummeryEcss_amount))
                strsqlHdr = strsqlHdr & ",'" & Trim(txtSEcssCode.Text) & "'," & Trim(lblSEcssCode.Text) & "," & Val(CStr(dblSummerySEcss_amount))
                strsqlHdr = strsqlHdr & "," & dblNewPacking & "," & dblSummeryPacking_amount & "," & CDec(txtMRP.Text) & ""
                strsqlHdr = strsqlHdr & ",'" & Trim(lblHSNSACCODE.Text) & "','" & Trim(lblHSNSAC.Text) & "','" & Trim(txtCGSTType.Text) & "'," & Val(lblCGSTPercent.Text) & ""
                strsqlHdr = strsqlHdr & ",'" & Trim(txtSGSTType.Text) & "'," & Val(lblSGSTPercent.Text) & ""
                strsqlHdr = strsqlHdr & ",'" & Trim(txtUTGSTType.Text) & "'," & Val(lblUTGSTPercent.Text) & ""
                strsqlHdr = strsqlHdr & ",'" & Trim(txtIGSTType.Text) & "'," & Val(lblIGSTPercent.Text) & ""
                strsqlHdr = strsqlHdr & ",'" & Trim(txtCompCessType.Text) & "'," & Val(lblCompCessPercent.Text) & "," & dblCGSTAmount & "," & dblSGSTAmount & "," & dblUTGSTAmount & "," & dblIGSTAmount & "," & dblCCESSAmount & ")"
                mstrSOType.Value = GetSOType()
                Call ChangeInvTypeCaption(mstrSOType.Value)
                '101188073
                If gblnGSTUnit Then
                    strSqlDtl = objInvoiceCls.SaveDataString(mP_Connection, strSelInvoices, strNotSelInvoices, CDec(txtMRP.Text), dblNewRate, dblNewPacking, dblNewCustSuppMat, dblNewToolCost, lblExctax_Per.Text, lblCVD_Per.Text, lblSAD_per.Text, lblEcssCode.Text, lblSaltax_Per.Text, lblSurcharge_Per.Text, txtLocationCode.Text, txtChallanNo.Text, lblItemCode.Text, txtCustPartCode.Text, getDateForDB(dtpDateFrom.Value), getDateForDB(dtpDateTo.Value), txtCustCode.Text, mP_User, lblSEcssCode.Text, lblCGSTPercent.Text, lblSGSTPercent.Text, lblUTGSTPercent.Text, lblIGSTPercent.Text, lblCompCessPercent.Text)
                Else
                    strSqlDtl = objInvoiceCls.SaveDataString(mP_Connection, strSelInvoices, strNotSelInvoices, CDec(txtMRP.Text), dblNewRate, dblNewPacking, dblNewCustSuppMat, dblNewToolCost, lblExctax_Per.Text, lblCVD_Per.Text, lblSAD_per.Text, lblEcssCode.Text, lblSaltax_Per.Text, lblSurcharge_Per.Text, txtLocationCode.Text, txtChallanNo.Text, lblItemCode.Text, txtCustPartCode.Text, getDateForDB(dtpDateFrom.Value), getDateForDB(dtpDateTo.Value), txtCustCode.Text, mP_User, lblSEcssCode.Text)
                End If
                '101188073
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strsqlHdr = "update SupplementaryInv_hdr set TotalInvoiceAmtRoundOff_diff = " & dblTotalInvoiceAmtRoundOff_diff & " ,Cust_ref = '" & Trim(txtRefNo.Text) & "',Amendment_No = '" & Trim(txtAmendment.Text) & "',"
                strsqlHdr = strsqlHdr & "Currency_Code ='" & Trim(lblCurrencyDes.Text) & "' ,Rate = " & dblSummeryRate & ",Basic_Amount = " & dblSummeryBasic_Amount
                strsqlHdr = strsqlHdr & ",Accessible_amount = " & dblSummeryAccessible_amount & ",Excise_type = '" & Trim(txtExciseTaxType.Text) & "',"
                strsqlHdr = strsqlHdr & " CVD_type = '" & Trim(txtCVDCode.Text) & "',SAD_type = '" & Trim(txtSADCode.Text) & "', Excise_per = " & Val(lblExctax_Per.Text)
                strsqlHdr = strsqlHdr & ",CVD_per = " & Val(lblCVD_Per.Text) & ",SVD_per = " & Val(lblSAD_per.Text)
                strsqlHdr = strsqlHdr & ",TotalExciseAmount = " & dblSummeryTotalExciseAmount & ",CustMtrl_Amount = " & dblSummeryCustMtrl_Amount & ","
                strsqlHdr = strsqlHdr & "ToolCost_amount = " & dblSummeryToolCost_amount & ",SalesTax_Type = '" & Trim(txtSaleTaxType.Text) & "',"
                strsqlHdr = strsqlHdr & "SalesTax_Per =" & Val(lblSaltax_Per.Text) & ",Sales_Tax_Amount = " & dblSummerySales_Tax_Amount
                strsqlHdr = strsqlHdr & ",Surcharge_salesTaxType = '" & Trim(txtSurchargeTaxType.Text) & "',Surcharge_SalesTax_Per = " & Val(lblSurcharge_Per.Text)
                strsqlHdr = strsqlHdr & ",Surcharge_Sales_Tax_Amount = " & dblSummerySurcharge_Sales_Tax_Amount & ",total_amount = " & dblSummerytotal_amount
                strsqlHdr = strsqlHdr & ", SuppInv_Remarks = '" & Trim(txtCustRefRemarks.Text) & "', remarks = '" & Trim(txtRemarks.Text) & "'"
                strsqlHdr = strsqlHdr & ",ECESS_Type = '" & Trim(Me.txtEcssCode.Text) & "',ECESS_Per = " & Trim(Me.lblEcssCode.Text) & ",ECESS_Amount = " & dblSummeryEcss_amount
                strsqlHdr = strsqlHdr & ",SECESS_Type = '" & Trim(txtSEcssCode.Text) & "',SECESS_Per = " & Trim(lblSEcssCode.Text) & ",SECESS_Amount = " & dblSummerySEcss_amount
                strsqlHdr = strsqlHdr & ",Packing = " & dblNewPacking
                strsqlHdr = strsqlHdr & ",Packing_Amount = " & dblSummeryPacking_amount & ",MRP=" & Val(txtMRP.Text)
                strsqlHdr = strsqlHdr & ", Upd_dt = '" & getDateForDB(GetServerDate()) & "',Upd_Userid = '" & mP_User & "' "
                strsqlHdr = strsqlHdr & ", HSN_SAC_CODE='" & Trim(lblHSNSACCODE.Text) & "',ISHSNORSAC='" & Trim(lblHSNSAC.Text) & "'"
                strsqlHdr = strsqlHdr & ", CGSTTXRT_TYPE='" & Trim(txtCGSTType.Text) & "',CGST_PERCENT=" & Val(lblCGSTPercent.Text) & ""
                strsqlHdr = strsqlHdr & ", SGSTTXRT_TYPE='" & Trim(txtSGSTType.Text) & "',SGST_PERCENT=" & Val(lblSGSTPercent.Text) & ""
                strsqlHdr = strsqlHdr & ", UTGSTTXRT_TYPE='" & Trim(txtUTGSTType.Text) & "',UTGST_PERCENT=" & Val(lblUTGSTPercent.Text) & ""
                strsqlHdr = strsqlHdr & ", IGSTTXRT_TYPE='" & Trim(txtIGSTType.Text) & "',IGST_PERCENT=" & Val(lblIGSTPercent.Text) & ""
                strsqlHdr = strsqlHdr & ", COMPENSATION_CESS_TYPE='" & Trim(txtCompCessType.Text) & "',COMPENSATION_CESS_PERCENT=" & Val(lblCompCessPercent.Text) & ",CGST_AMT=" & dblCGSTAmount & ""
                strsqlHdr = strsqlHdr & ", SGST_AMT=" & dblSGSTAmount & ",UTGST_AMT=" & dblUTGSTAmount & ",IGST_AMT=" & dblIGSTAmount & ",CCESS_AMT=" & dblCCESSAmount & ""
                strsqlHdr = strsqlHdr & " where Unit_code='" & gstrUnitId & "' and Location_code = '"
                strsqlHdr = strsqlHdr & Trim(txtLocationCode.Text) & "' and Doc_no = " & Trim(txtChallanNo.Text)
                '************String For Deletion
                strDeleteSuppDtl = "Delete from SupplementaryInv_Dtl Where Unit_code='" & gstrUnitId & "' and Doc_no = '" & Trim(txtChallanNo.Text) & "'"
                strDeleteSuppCrAdvise = "Delete from SuppCreditAdvise_Dtl Where Unit_code='" & gstrUnitId & "' and Doc_no = '" & Trim(txtChallanNo.Text) & "'"
                mstrSOType.Value = GetSOType()
                Call ChangeInvTypeCaption(mstrSOType.Value)
                '101188073
                If gblnGSTUnit Then
                    strSqlDtl = objInvoiceCls.SaveDataString(mP_Connection, strSelInvoices, strNotSelInvoices, Convert.ToDecimal(IIf(txtMRP.Text = "", 0, txtMRP.Text)), dblNewRate, dblNewPacking, dblNewCustSuppMat, dblNewToolCost, lblExctax_Per.Text, lblCVD_Per.Text, lblSAD_per.Text, lblEcssCode.Text, lblSaltax_Per.Text, lblSurcharge_Per.Text, txtLocationCode.Text, txtChallanNo.Text, lblItemCode.Text, txtCustPartCode.Text, getDateForDB(dtpDateFrom.Value), getDateForDB(dtpDateTo.Value), txtCustCode.Text, mP_User, lblSEcssCode.Text, lblCGSTPercent.Text, lblSGSTPercent.Text, lblUTGSTPercent.Text, lblIGSTPercent.Text, lblCompCessPercent.Text)
                Else
                    strSqlDtl = objInvoiceCls.SaveDataString(mP_Connection, strSelInvoices, strNotSelInvoices, Convert.ToDecimal(IIf(txtMRP.Text = "", 0, txtMRP.Text)), dblNewRate, dblNewPacking, dblNewCustSuppMat, dblNewToolCost, lblExctax_Per.Text, lblCVD_Per.Text, lblSAD_per.Text, lblEcssCode.Text, lblSaltax_Per.Text, lblSurcharge_Per.Text, txtLocationCode.Text, txtChallanNo.Text, lblItemCode.Text, txtCustPartCode.Text, getDateForDB(dtpDateFrom.Value), getDateForDB(dtpDateTo.Value), txtCustCode.Text, mP_User, lblSEcssCode.Text)
                End If
                '101188073

        End Select
        Dim strTemp1 As String
        Dim intOccur As Short
        If Len(Trim(strsqlHdr)) > 0 Then
            mP_Connection.BeginTrans()
            mP_Connection.Execute("set dateFormat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            mP_Connection.Execute(strsqlHdr, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                If Len(Trim(strDeleteSuppDtl)) > 0 Then
                    mP_Connection.Execute(strDeleteSuppDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
                If Len(Trim(strDeleteSuppCrAdvise)) > 0 Then
                    mP_Connection.Execute(strDeleteSuppCrAdvise, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
            End If
            strTemp1 = Replace(strSqlDtl, "|", "")
            intOccur = Len(strSqlDtl) - Len(strTemp1)
            If intOccur > 1 Then
                strSqlCreditAdvDtl = Mid(strSqlDtl, InStrRev(strSqlDtl, "|") + 1, Len(strSqlDtl))
            End If
            strUpdateSalesDtl = Mid(strSqlDtl, InStr(1, strSqlDtl, "|") + 1, (InStrRev(strSqlDtl, "|") - InStr(1, strSqlDtl, "|")) - 1)
            strSqlDtl = VB.Left(strSqlDtl, InStr(1, strSqlDtl, "|") - 1)
            mP_Connection.Execute(strSqlDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            If Len(Trim(strSqlCreditAdvDtl)) > 0 Then
                mP_Connection.Execute(strSqlCreditAdvDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            If Len(Trim(strUpdateSalesDtl)) > 0 Then
                mP_Connection.Execute(strUpdateSalesDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            mP_Connection.CommitTrans()
        End If
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        MsgBox("Transaction Completed successfully.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SelectChallanNoFromSupplementatryInvHdr()
        On Error GoTo ErrHandler
        Dim strChallanNo As String
        Dim rsChallanNo As ClsResultSetDB
        strChallanNo = "Select max(Doc_No) as Doc_No from SupplementaryInv_hdr where Unit_code='" & gstrUNITID & "' and Doc_No>" & 99000000
        rsChallanNo = New ClsResultSetDB
        rsChallanNo.GetResult(strChallanNo, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsChallanNo.GetNoRows > 0 Then
            If Val(rsChallanNo.GetValue("Doc_No")) = 0 Then
                txtChallanNo.Text = "99000001"
            Else
                txtChallanNo.Text = CStr(Val(rsChallanNo.GetValue("Doc_No")) + 1)
            End If
        Else
            txtChallanNo.Text = "99000001"
        End If
        rsChallanNo.ResultSetClose()
        rsChallanNo = Nothing
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Sub ToLockCellofGrid()
        On Error GoTo ErrHandler
        With spdPrevInv
            .Row = 1
            .Row2 = .MaxRows
            .Col = enumPreInvoiceDetails.New_Rate
            .Col2 = enumPreInvoiceDetails.New_Rate
            .BlockMode = True
            .Lock = False
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewPacking
            .Col2 = enumPreInvoiceDetails.NewPacking
            .BlockMode = True
            .Lock = False
            .BlockMode = False
            .Col = enumPreInvoiceDetails.NewCustSuppMaterial
            .Col2 = enumPreInvoiceDetails.NewCustSuppMaterial
            .BlockMode = True
            .Lock = False
            .BlockMode = False
            .Col = enumPreInvoiceDetails.newToolCost
            .Col2 = enumPreInvoiceDetails.newToolCost
            .BlockMode = True
            .Lock = False
            .BlockMode = False
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Sub DeleteRecords()
        Dim strUpdateSalesDtl As String
        Dim strDelete As String
        Dim dblPrevInvoiceNo As Double
        Dim intLoopCounter As Short
        Dim varFalg As Object
        Dim sarrInvoices() As String
        On Error GoTo ErrHandler
        strUpdateSalesDtl = ""
        If strSelInvoices = "" Then
            FillInvoiceNumber()
            For intLoopCounter = 0 To lstInv.Items.Count - 1
                strSelInvoices = strSelInvoices & "|" & lstInv.Items.Item(intLoopCounter).Text
            Next intLoopCounter
            If VB.Left(strSelInvoices, 1) = "|" Then
                strSelInvoices = VB.Right(strSelInvoices, Len(strSelInvoices) - 1)
            End If
        End If
        sarrInvoices = Split(strSelInvoices, "|")
        For intLoopCounter = 0 To UBound(sarrInvoices)
            strUpdateSalesDtl = strUpdateSalesDtl & "update sales_dtl set SupplementaryInvoiceFlag = 0 where Unit_code='" & gstrUNITID & "' and doc_no = "
            strUpdateSalesDtl = strUpdateSalesDtl & sarrInvoices(intLoopCounter) & " and Location_code = '" & Trim(txtLocationCode.Text)
            strUpdateSalesDtl = strUpdateSalesDtl & "' and Cust_item_code = '" & Trim(txtCustPartCode.Text)
            strUpdateSalesDtl = strUpdateSalesDtl & "' and Item_code = '" & Trim(lblItemCode.Text) & "'" & vbCrLf
        Next
        strDelete = ""
        strDelete = " Delete From SuppCreditAdvise_Dtl where Unit_code='" & gstrUNITID & "' and doc_no = " & Trim(txtChallanNo.Text)
        strDelete = strDelete & " and Location_code = '" & Trim(txtLocationCode.Text) & "'" & vbCrLf
        strDelete = strDelete & "Delete From supplementaryInv_dtl where Unit_code='" & gstrUNITID & "' and doc_no = " & Trim(txtChallanNo.Text)
        strDelete = strDelete & " and Location_code = '" & Trim(txtLocationCode.Text) & "'" & vbCrLf
        strDelete = strDelete & " Delete From supplementaryInv_hdr where Unit_code='" & gstrUNITID & "' and doc_no = " & Trim(txtChallanNo.Text)
        strDelete = strDelete & " and Location_code = '" & Trim(txtLocationCode.Text) & "'"
        If Len(Trim(strUpdateSalesDtl)) > 0 Then
            mP_Connection.BeginTrans()
            mP_Connection.Execute(strUpdateSalesDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            mP_Connection.Execute(strDelete, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            mP_Connection.CommitTrans()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Sub SetWidthofColumnsinGrid_Old()
        On Error GoTo ErrHandler
        With spdPrevInv
            .set_ColWidth(enumPreInvoiceDetails.New_Rate, 1100)
            .set_ColWidth(enumPreInvoiceDetails.NewPacking, 1000)
            .set_ColWidth(enumPreInvoiceDetails.NewCustSuppMaterial, 1375)
            .set_ColWidth(enumPreInvoiceDetails.newToolCost, 1100)
            .set_ColWidth(enumPreInvoiceDetails.NewExciseValue, 1000)
            .set_ColWidth(enumPreInvoiceDetails.NewEcessValue, 1000)
            .set_ColWidth(enumPreInvoiceDetails.NewCVDValue, 1000)
            .set_ColWidth(enumPreInvoiceDetails.NewSADValue, 1000)
            .set_ColWidth(enumPreInvoiceDetails.NewSalesTaxValue, 1000)
            .set_ColWidth(enumPreInvoiceDetails.NewSSTValue, 1000)
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Sub SetWidthofColumnsinGrid()
        On Error GoTo ErrHandler
        With spdPrevInv
            .set_ColWidth(enumPreInvoiceDetails.Select_Invoice, 0)
            .Col = enumPreInvoiceDetails.Select_Invoice
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.Invoice_No, 0)
            .Col = enumPreInvoiceDetails.Invoice_No
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.Invoice_Date, 0)
            .Col = enumPreInvoiceDetails.Invoice_Date
            .Lock = False
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.LastSupplementary, 0)
            .Col = enumPreInvoiceDetails.LastSupplementary
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.SupplementaryDate, 0)
            .Col = enumPreInvoiceDetails.SupplementaryDate
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.Quantity, 0)
            .Col = enumPreInvoiceDetails.Quantity
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.Rate, 0)
            .Col = enumPreInvoiceDetails.Rate
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.New_Rate, 1100)
            .set_ColWidth(enumPreInvoiceDetails.Rate_diff, 0)
            .Col = enumPreInvoiceDetails.Rate_diff
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.TotalPacking, 0)
            .Col = enumPreInvoiceDetails.TotalPacking
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewPacking, 1000)
            .set_ColWidth(enumPreInvoiceDetails.NewTotalPacking, 0)
            .Col = enumPreInvoiceDetails.NewTotalPacking
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.Basic, 0)
            .Col = enumPreInvoiceDetails.Basic
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewBasic, 0)
            .Col = enumPreInvoiceDetails.NewBasic
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.BasicDiff, 0)
            .Col = enumPreInvoiceDetails.BasicDiff
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.TotalCustSuppMaterial, 0)
            .Col = enumPreInvoiceDetails.TotalCustSuppMaterial
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewCustSuppMaterial, 1375)
            .set_ColWidth(enumPreInvoiceDetails.NewTotalCustSuppMaterial, 0)
            .Col = enumPreInvoiceDetails.NewTotalCustSuppMaterial
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.CustSuppMaterial_diff, 0)
            .Col = enumPreInvoiceDetails.CustSuppMaterial_diff
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.ToolCost, 0)
            .Col = enumPreInvoiceDetails.ToolCost
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.newToolCost, 1100)
            .set_ColWidth(enumPreInvoiceDetails.NewTotalToolCost, 0)
            .Col = enumPreInvoiceDetails.NewTotalToolCost
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.ToolCost_diff, 0)
            .Col = enumPreInvoiceDetails.ToolCost_diff
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.AccessableValue, 0)
            .Col = enumPreInvoiceDetails.AccessableValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewAccessableValue, 0)
            .Col = enumPreInvoiceDetails.NewAccessableValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.AccessableValue_Diff, 0)
            .Col = enumPreInvoiceDetails.AccessableValue_Diff
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.TotalExciseValue, 0)
            .Col = enumPreInvoiceDetails.TotalExciseValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewExciseValue, 0)
            .Col = enumPreInvoiceDetails.NewExciseValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.TotalEcessValue, 0)
            .Col = enumPreInvoiceDetails.TotalEcessValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewEcessValue, 0)
            .Col = enumPreInvoiceDetails.NewEcessValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.TotalEcssDiff, 0)
            .Col = enumPreInvoiceDetails.TotalEcssDiff
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.TotalsEcessValue, 0)
            .Col = enumPreInvoiceDetails.TotalsEcessValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewsEcessValue, 0)
            .Col = enumPreInvoiceDetails.NewsEcessValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.TotalsEcssDiff, 0)
            .Col = enumPreInvoiceDetails.TotalsEcssDiff
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewCVDValue, 0)
            .Col = enumPreInvoiceDetails.NewCVDValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewSADValue, 0)
            .Col = enumPreInvoiceDetails.NewSADValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewTotalExciseValue, 0)
            .Col = enumPreInvoiceDetails.NewTotalExciseValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.TotalExciseValueDiff, 0)
            .Col = enumPreInvoiceDetails.TotalExciseValueDiff
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.SalesTaxValue, 0)
            .Col = enumPreInvoiceDetails.SalesTaxValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewSalesTaxValue, 0)
            .Col = enumPreInvoiceDetails.NewSalesTaxValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.SalesTaxValueDiff, 0)
            .Col = enumPreInvoiceDetails.SalesTaxValueDiff
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.SSTVlaue, 0)
            .Col = enumPreInvoiceDetails.SSTVlaue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NewSSTValue, 0)
            .Col = enumPreInvoiceDetails.NewSSTValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.SSTVlaueDiff, 0)
            .Col = enumPreInvoiceDetails.SSTVlaueDiff
            '101188073
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.CGST_AMT, 0)
            .Col = enumPreInvoiceDetails.CGST_AMT
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NEW_CGST_AMT, 0)
            .Col = enumPreInvoiceDetails.NEW_CGST_AMT
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.DIFF_CGST_AMT, 0)
            .Col = enumPreInvoiceDetails.DIFF_CGST_AMT
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.SGST_AMT, 0)
            .Col = enumPreInvoiceDetails.SGST_AMT
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NEW_SGST_AMT, 0)
            .Col = enumPreInvoiceDetails.NEW_SGST_AMT
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.DIFF_SGST_AMT, 0)
            .Col = enumPreInvoiceDetails.DIFF_SGST_AMT
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.UTGST_AMT, 0)
            .Col = enumPreInvoiceDetails.UTGST_AMT
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NEW_UTGST_AMT, 0)
            .Col = enumPreInvoiceDetails.NEW_UTGST_AMT
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.DIFF_UTGST_AMT, 0)
            .Col = enumPreInvoiceDetails.DIFF_UTGST_AMT
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.IGST_AMT, 0)
            .Col = enumPreInvoiceDetails.IGST_AMT
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NEW_IGST_AMT, 0)
            .Col = enumPreInvoiceDetails.NEW_IGST_AMT
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.DIFF_IGST_AMT, 0)
            .Col = enumPreInvoiceDetails.DIFF_IGST_AMT
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.CCESS_AMT, 0)
            .Col = enumPreInvoiceDetails.CCESS_AMT
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.NEW_CCESS_AMT, 0)
            .Col = enumPreInvoiceDetails.NEW_CCESS_AMT
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.DIFF_CCESS_AMT, 0)
            .Col = enumPreInvoiceDetails.DIFF_CCESS_AMT
            '101188073
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.TotalCurrInvValue, 0)
            .Col = enumPreInvoiceDetails.TotalCurrInvValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.TotalInvoiceValue, 0)
            .Col = enumPreInvoiceDetails.TotalInvoiceValue
            .ColHidden = True
            .set_ColWidth(enumPreInvoiceDetails.flag, 0)
            .Col = enumPreInvoiceDetails.flag
            .ColHidden = True
            .Lock = False
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Sub DisplayDetailsinEditMode()
        Dim strSalesDetailData As String
        Dim rsSalesDtl As ClsResultSetDB
        Dim intMaxRow As Short
        Dim intLoopCounter As Short
        Dim strSupplementary As String
        Dim strRefDocNo As String
        Dim strCustItemCode As String
        Dim strItemCode As String
        Dim dblPrevInvoiceNo As Double
        On Error GoTo ErrHandler
        rsSalesDtl = New ClsResultSetDB
        intMaxRow = spdPrevInv.MaxRows
        strInvoiceNo = ""
        Dim sarrInvoiceNo() As String
        sarrInvoiceNo = Split(strSelInvoices, "|")
        For intLoopCounter = 0 To UBound(sarrInvoiceNo)
            dblPrevInvoiceNo = Val(sarrInvoiceNo(intLoopCounter))
            If Len(Trim(strInvoiceNo)) > 0 Then
                strInvoiceNo = strInvoiceNo & "," & dblPrevInvoiceNo
            Else
                strInvoiceNo = CStr(dblPrevInvoiceNo)
            End If
        Next
        If VB.Left(strInvoiceNo, 1) = "," Then
            strInvoiceNo = VB.Right(strInvoiceNo, Len(strInvoiceNo) - 1)
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        Call SetCellTypeofGrids()
        Call SetMaxLengthofGrid(4)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Arrow)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Sub DisplayDetailsinEditMode_Old()
        Dim strSalesDetailData As String
        Dim rsSalesDtl As ClsResultSetDB
        Dim intMaxRow As Short
        Dim intLoopCounter As Short
        Dim rsLastSupplementary As ClsResultSetDB
        Dim strSupplementary As String
        Dim strRefDocNo As String
        Dim strCustItemCode As String
        Dim strItemCode As String
        Dim dblPrevInvoiceNo As Double
        On Error GoTo ErrHandler
        rsSalesDtl = New ClsResultSetDB
        intMaxRow = spdPrevInv.MaxRows
        strInvoiceNo = ""
        Dim sarrInvoiceNo() As String
        For intLoopCounter = 1 To intMaxRow
            Call spdPrevInv.GetFloat(enumPreInvoiceDetails.Invoice_No, intLoopCounter, dblPrevInvoiceNo)
            If Len(Trim(CStr(dblPrevInvoiceNo))) > 0 Then
                strInvoiceNo = strInvoiceNo & "," & dblPrevInvoiceNo
            Else
                strInvoiceNo = CStr(dblPrevInvoiceNo)
            End If
        Next
        If VB.Left(strInvoiceNo, 1) = "," Then
            strInvoiceNo = VB.Right(strInvoiceNo, Len(strInvoiceNo) - 1)
        End If
        strSalesDetailData = "select * from SupplementaryData('" & gstrUNITID & "','" & Format(dtpDateFrom.Value, "dd MMM yyyy") & "','"
        strSalesDetailData = strSalesDetailData & Format(dtpDateTo.Value, "dd MMM yyyy") & "','" & txtCustCode.Text & "','" & lblItemCode.Text
        strSalesDetailData = strSalesDetailData & "','" & txtCustPartCode.Text & "') Where Doc_no Not in (" & strInvoiceNo & ")"
        rsSalesDtl.GetResult(strSalesDetailData)
        intMaxRow = rsSalesDtl.GetNoRows
        rsSalesDtl.MoveFirst()
        If intMaxRow > 0 Then
            intLoopCounter = spdPrevInv.MaxRows + 1
            spdPrevInv.MaxRows = spdPrevInv.MaxRows + intMaxRow
            intMaxRow = spdPrevInv.MaxRows
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            Call SetCellTypeofGrids()
            Call SetMaxLengthofGrid(4)
            With spdPrevInv
                rsLastSupplementary = New ClsResultSetDB
                For intLoopCounter = intLoopCounter To intMaxRow
                    Call .SetText(enumPreInvoiceDetails.Invoice_No, intLoopCounter, rsSalesDtl.GetValue("Doc_No"))
                    strRefDocNo = rsSalesDtl.GetValue("Doc_No")
                    .SetText(enumPreInvoiceDetails.Invoice_Date, intLoopCounter, VB6.Format(rsSalesDtl.GetValue("Invoice_date"), gstrDateFormat))
                    .SetText(enumPreInvoiceDetails.Quantity, intLoopCounter, rsSalesDtl.GetValue("Sales_Quantity"))
                    .SetText(enumPreInvoiceDetails.Rate, intLoopCounter, rsSalesDtl.GetValue("Rate"))
                    .SetText(enumPreInvoiceDetails.New_Rate, intLoopCounter, 0)
                    .SetText(enumPreInvoiceDetails.TotalCustSuppMaterial, intLoopCounter, rsSalesDtl.GetValue("CustMtrl_amount"))
                    .SetText(enumPreInvoiceDetails.ToolCost, intLoopCounter, rsSalesDtl.GetValue("ToolCost_amount"))
                    .SetText(enumPreInvoiceDetails.TotalPacking, intLoopCounter, rsSalesDtl.GetValue("Packing"))
                    .SetText(enumPreInvoiceDetails.TotalExciseValue, intLoopCounter, rsSalesDtl.GetValue("Excise_amount"))
                    .SetText(enumPreInvoiceDetails.TotalEcessValue, intLoopCounter, rsSalesDtl.GetValue("Ecess_Amount"))
                    .SetText(enumPreInvoiceDetails.SalesTaxValue, intLoopCounter, rsSalesDtl.GetValue("SalesTax_Amount"))
                    .SetText(enumPreInvoiceDetails.SSTVlaue, intLoopCounter, rsSalesDtl.GetValue("SSalesTax_Amount"))
                    .SetText(enumPreInvoiceDetails.Basic, intLoopCounter, rsSalesDtl.GetValue("Basic_amount"))
                    .SetText(enumPreInvoiceDetails.AccessableValue, intLoopCounter, rsSalesDtl.GetValue("Accessible_amount"))
                    If rsSalesDtl.GetValue("SupplemnetryInv") = True Then
                        strSupplementary = "select a.Doc_No,b.Invoice_date,a.RefDoc_No,a.Item_code,a.Cust_Item_Code from"
                        strSupplementary = strSupplementary & " SupplementaryInv_dtl a,SupplementaryInv_hdr b Where "
                        strSupplementary = strSupplementary & " a.Unit_code = b.Unit_code and a.Unit_code='" & gstrUNITID & "' and a.Doc_No = b.Doc_No And a.Location_Code = b.Location_Code"
                        strSupplementary = strSupplementary & " and a.RefDoc_no = '" & strRefDocNo & "' and a.Item_code = '"
                        strSupplementary = strSupplementary & Trim(lblItemCode.Text) & "' and a.Cust_Item_code = '"
                        strSupplementary = strSupplementary & Trim(txtCustPartCode.Text) & "' and b.Invoice_date < '" & Format(GetServerDate, "dd MMM yyyy") & "'"
                        strSupplementary = strSupplementary & " order by Invoice_date,a.Doc_no "
                        rsLastSupplementary.GetResult(strSupplementary)
                        If rsLastSupplementary.GetNoRows > 0 Then
                            rsLastSupplementary.MoveLast()
                            .SetText(enumPreInvoiceDetails.LastSupplementary, intLoopCounter, rsLastSupplementary.GetValue("Doc_no"))
                            .SetText(enumPreInvoiceDetails.SupplementaryDate, intLoopCounter, VB6.Format(rsLastSupplementary.GetValue("Invoice_date"), gstrDateFormat))
                        End If
                        rsLastSupplementary.ResultSetClose()
                        .Row = intLoopCounter
                        .Row2 = intLoopCounter
                        .Col = 1
                        .Col2 = .MaxCols
                        .BlockMode = True : .ForeColor = System.Drawing.Color.Blue
                        .BlockMode = False
                    Else
                        .Row = intLoopCounter
                        .Row2 = intLoopCounter
                        .Col = 1
                        .Col2 = .MaxCols
                        .BlockMode = True : .ForeColor = System.Drawing.SystemColors.WindowText
                        .BlockMode = False
                    End If
                    rsSalesDtl.MoveNext()
                Next
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Arrow)
                .SetRefStyle(2)
                If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    Call SetFormulaofColumns(0, 4)
                    ToShowDatainSummeryGrid()
                End If
            End With
        End If
        rsSalesDtl.ResultSetClose()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function ToCheckForNegativeValuesNoMessages(ByRef pintRow As Integer) As Boolean
        Dim intMaxLoop As Short
        Dim intLoopCounter As Short
        On Error GoTo ErrHandler
        ToCheckForNegativeValuesNoMessages = False
        intMaxLoop = pintRow
        For intLoopCounter = pintRow To intMaxLoop
            With spdPrevInv
                'Checking For Basic Values
                .Row = pintRow
                .Col = enumPreInvoiceDetails.BasicDiff
                If Val(.Text) < 0 Then
                    ToCheckForNegativeValuesNoMessages = False
                    Exit Function
                End If
                'Checking For Accessable Values
                .Row = pintRow
                .Col = enumPreInvoiceDetails.AccessableValue_Diff
                If Val(.Text) < 0 Then
                    ToCheckForNegativeValuesNoMessages = False
                    Exit Function
                End If
                'Checking For Total InvoiceValue
                .Row = pintRow
                .Col = enumPreInvoiceDetails.TotalInvoiceValue
                If Val(.Text) < 0 Then
                    ToCheckForNegativeValuesNoMessages = False
                    Exit Function
                End If
            End With
        Next
        ToCheckForNegativeValuesNoMessages = True
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function FillInvoiceNumber() As Boolean
        Dim strSQL As String
        Dim rsSalesDtl As New ClsResultSetDB
        Dim rsLastSupplementary As ClsResultSetDB
        Dim intMaxRow As Short
        Dim Intcounter As Short
        Dim intListCounter As Short
        Dim intListIndex As Short
        Dim sarrSelInvoices() As String
        lstInv.Items.Clear()
        lstInv.Columns.Item(0).Width = VB6.TwipsToPixelsX(2000)
        lstInv.Columns.Item(1).Width = VB6.TwipsToPixelsX(1000)
        lstInv.View = System.Windows.Forms.View.Details
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            sarrSelInvoices = Split(strSelInvoices, "|")
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
        mstrSOType.Value = GetSOType()
        Call ChangeInvTypeCaption(mstrSOType.Value)
        strSQL = "select * from SupplementaryData('" & gstrUNITID & "','" & Format(dtpDateFrom.Value, "dd MMM yyyy") & "','"
        strSQL = strSQL & Format(dtpDateTo.Value, "dd MMM yyyy") & "','" & txtCustCode.Text & "','" & lblItemCode.Text
        strSQL = strSQL & "','" & txtCustPartCode.Text & "', " & CmdGrpChEnt.Mode & ", '" & txtChallanNo.Text & "','" & mstrSOType.Value & "') Order By Doc_No"
        rsSalesDtl.GetResult(strSQL)
        intMaxRow = rsSalesDtl.GetNoRows
        rsSalesDtl.MoveFirst()
        If intMaxRow > 0 Then
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            rsSalesDtl.MoveFirst()
            For Intcounter = 0 To intMaxRow - 1
                lstInv.Items.Insert(Intcounter, rsSalesDtl.GetValue("Doc_No"))
                If lstInv.Items.Item(Intcounter).SubItems.Count > 1 Then
                    lstInv.Items.Item(Intcounter).SubItems(1).Text = rsSalesDtl.GetValue("Sales_Quantity")
                Else
                    lstInv.Items.Item(Intcounter).SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsSalesDtl.GetValue("Sales_Quantity")))
                End If
                rsSalesDtl.MoveNext()
            Next Intcounter
            If strSelInvoices <> "" Then
                chkUnCheckall.CheckState = System.Windows.Forms.CheckState.Checked
                chkUnCheckall_CheckStateChanged(chkUnCheckall, New System.EventArgs())
                chkUnCheckall.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkSelectAll.CheckState = System.Windows.Forms.CheckState.Unchecked
                sarrSelInvoices = Split(strSelInvoices, "|")
                For intListCounter = 0 To UBound(sarrSelInvoices)
                    If sarrSelInvoices(intListCounter) <> "" Then
                        intListIndex = lstInv.FindItemWithText(sarrSelInvoices(intListCounter)).Index
                        lstInv.Items.Item(intListIndex).Checked = True
                    End If
                Next intListCounter
            Else
                chkUnCheckall.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkSelectAll.CheckState = System.Windows.Forms.CheckState.Checked
                chkSelectAll_CheckStateChanged(chkSelectAll, New System.EventArgs())
            End If
            FillInvoiceNumber = True
        Else
            FillInvoiceNumber = False
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
    End Function
    Private Function SelectData(ByVal pstrFldName As String, ByRef pstrTblName As String, ByRef pstrCond As String) As String
        Dim rstemp As New ClsResultSetDB
        Dim strSQL As String
        Dim strRetString As String
        strSQL = "Select " & pstrFldName & " from " & pstrTblName & pstrCond
        rstemp.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        If rstemp.GetNoRows > 0 Then
            rstemp.MoveFirst()
            Do While Not rstemp.EOFRecord
                strRetString = strRetString & "|" & rstemp.GetValue(pstrFldName)
                rstemp.MoveNext()
            Loop
        Else
            strRetString = ""
        End If
        rstemp.ResultSetClose()
        If VB.Left(strRetString, 1) = "|" Then
            strRetString = VB.Right(strRetString, Len(strRetString) - 1)
        End If
        SelectData = strRetString
    End Function
    Private Sub SetWidthOfCols()
        On Error GoTo ErrHandler
        With spdPrevInv
            .set_ColWidth(enumPreInvoiceDetails.New_Rate, 1100)
            .set_ColWidth(enumPreInvoiceDetails.NewPacking, 1000)
            .set_ColWidth(enumPreInvoiceDetails.NewCustSuppMaterial, 1375)
            .set_ColWidth(enumPreInvoiceDetails.newToolCost, 1100)
            .set_ColWidth(enumPreInvoiceDetails.NewExciseValue, 1000)
            .set_ColWidth(enumPreInvoiceDetails.NewEcessValue, 1000)
            .set_ColWidth(enumPreInvoiceDetails.NewCVDValue, 1000)
            .set_ColWidth(enumPreInvoiceDetails.NewSADValue, 1000)
            .set_ColWidth(enumPreInvoiceDetails.NewSalesTaxValue, 1000)
            .set_ColWidth(enumPreInvoiceDetails.NewSSTValue, 1000)
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function LenOfArray(ByRef psArrayList() As Object) As Short
        On Error GoTo ErrHandler
        Dim lngElementCount As Integer
        Dim lngLoopCount As Integer
        For lngLoopCount = 0 To UBound(psArrayList)
            If psArrayList(lngLoopCount).ToString() <> "" Then
                lngElementCount = lngElementCount + 1
            End If
        Next
AssignNumber:
        LenOfArray = lngElementCount
ErrHandler:
        If Err.Number = 9 Then
            Err.Clear()
            lngElementCount = 0
            GoTo AssignNumber
        ElseIf Err.Number <> 0 Then
            MsgBox(Err.Description, MsgBoxStyle.Information)
        End If
    End Function
    Private Function GetInvoiceNumbers(ByVal Rs As ClsResultSetDB) As Object
        strSelInvoices = ""
        strNotSelInvoices = ""
        With Rs
            .MoveFirst()
            Do While Not .EOFRecord
                'if the invoice is selected
                If .GetValue("SelectInvoice") = 1 Then
                    strSelInvoices = strSelInvoices & "|" & .GetValue("RefDoc_No")
                    'if the invoice is not selected
                ElseIf .GetValue("SelectInvoice") = 0 Then
                    strNotSelInvoices = strNotSelInvoices & "|" & .GetValue("RefDoc_No")
                End If
                .MoveNext()
            Loop
        End With
        'remove the first character
        If VB.Left(strSelInvoices, 1) = "|" Then
            strSelInvoices = VB.Right(strSelInvoices, Len(strSelInvoices) - 1)
        End If
        'remove the first character
        If VB.Left(strNotSelInvoices, 1) = "|" Then
            strSelInvoices = VB.Right(strNotSelInvoices, Len(strNotSelInvoices) - 1)
        End If
    End Function
    Public Property GridColor() As Object
        Get
            GridColor = System.Drawing.ColorTranslator.ToOle(cColor)
        End Get
        Set(ByVal Value As Object)
            cColor = System.Drawing.ColorTranslator.FromOle(Value)
        End Set
    End Property
    Private Sub ChangeInvTypeCaption(ByVal pstrSOType As String)
        On Error GoTo ErrHandler
        If pstrSOType = "M" Then
            lblTypeOfInvDisplay.Text = "             INVOICES MADE AGAINST MRP WILL BE                 CONSIDERED ONLY"
        Else
            lblTypeOfInvDisplay.Text = "             INVOICES MADE WITHOUT MRP WILL BE                 CONSIDERED ONLY"
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function GetSOType() As String
        On Error GoTo ErrHandler
        If Val(txtMRP.Text) > 0 Then
            GetSOType = "M"
        Else
            GetSOType = ""
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub HideMRPlabel(ByVal pblnHide As Boolean)
        On Error GoTo ErrHandler
        lblTypeOfInvDisplay.Visible = pblnHide
        Image1.Visible = pblnHide
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub dtpDateFrom_KeyDown1(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpDateFrom.KeyDown
        On Error GoTo errHandler
        If e.KeyCode = Keys.Return Then
            dtpDateTo.Focus()
        End If
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub dtpDateFrom_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDateFrom.LostFocus
        On Error GoTo errHandler
        Dim dtToDate As Date
        If CheckFinancialYearDates() = True Then
            If dtpDateFrom.Value < Financial_Start_Date Then
                MsgBox("Date Can Not Be Less Then Business Period Start Date [" & VB6.Format(Financial_Start_Date, gstrDateFormat) & "].", vbInformation, "empower")
                dtpDateFrom.Focus()
                Exit Sub
            End If
        End If
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub dtpDateTo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpDateTo.KeyDown
        On Error GoTo errHandler
        If e.KeyCode = Keys.Return Then
            txtLocationCode.Focus()
        End If
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub dtpDateTo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDateTo.LostFocus
        On Error GoTo errHandler
        If CheckFinancialYearDates() = True Then
            If dtpDateTo.Value > Financial_End_Date Then
                MsgBox("Date Can Not Be Greater Then Business Period End Date [" & VB6.Format(Financial_End_Date, gstrDateFormat) & "].", vbInformation, "empower")
                dtpDateTo.Focus()
                Exit Sub
            ElseIf dtpDateTo.Value < dtpDateFrom.Value Then
                MsgBox("Date Can Not Be Less Then From Date [" & VB6.Format(dtpDateFrom.Value, gstrDateFormat) & "].", vbInformation, "empower")
                dtpDateFrom.Focus()
                Exit Sub
            End If
        End If
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub lstInv_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles lstInv.ItemChecked
        Dim Item As System.Windows.Forms.ListViewItem = (e.Item)
        Dim intListCount As Short
        Dim bChecked As Boolean
        Dim bUnCheck As Boolean
        On Error GoTo errHandler
        Dim lstItm As ListViewItem
        'For intListCount = 0 To lstInv.Items.Count - 1
        '    If lstInv.Items.Item(intListCount).Checked = True Then
        '        chkUnCheckall.CheckState = System.Windows.Forms.CheckState.Unchecked
        '        bUnCheck = True
        '        Exit For
        '    End If
        'Next intListCount
        For Each lstItm In lstInv.Items
            If IsNothing(lstItm) = True Then Continue For
            If lstItm.Checked = True Then
                chkUnCheckall.CheckState = System.Windows.Forms.CheckState.Unchecked
                bUnCheck = True
                Exit For
            End If
        Next

        'For intListCount = 0 To lstInv.Items.Count - 1
        '    If lstInv.Items.Item(intListCount).Checked = False Then
        '        chkSelectAll.CheckState = System.Windows.Forms.CheckState.Unchecked
        '        bChecked = True
        '        Exit For
        '    End If
        'Next intListCount
        lstItm = New ListViewItem
        For Each lstItm In lstInv.Items
            If IsNothing(lstItm) = True Then Continue For
            If lstItm.Checked = False Then
                chkSelectAll.CheckState = System.Windows.Forms.CheckState.Unchecked
                bChecked = True
                Exit For
            End If
        Next

        If bUnCheck = False Then
            chkUnCheckall.CheckState = System.Windows.Forms.CheckState.Checked
        End If
        If bChecked = False Then
            chkSelectAll.CheckState = System.Windows.Forms.CheckState.Checked
        End If
        Exit Sub
errHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    '101188073
    Private Sub txtCGSTType_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCGSTType.TextChanged, txtSGSTType.TextChanged, txtUTGSTType.TextChanged, txtIGSTType.TextChanged, txtCompCessType.TextChanged
        Try
            If Not gblnGSTUnit Then Exit Sub
            Dim txtGST As New TextBox
            txtGST = DirectCast(sender, TextBox)
            If Len(txtGST.Text.Trim) = 0 Then
                If txtGST.Name.ToUpper = "TXTCGSTTYPE" Then
                    lblCGSTPercent.Text = "0.00"
                ElseIf txtGST.Name.ToUpper = "TXTSGSTTYPE" Then
                    lblSGSTPercent.Text = "0.00"
                ElseIf txtGST.Name.ToUpper = "TXTUTGSTTYPE" Then
                    lblUTGSTPercent.Text = "0.00"
                ElseIf txtGST.Name.ToUpper = "TXTIGSTTYPE" Then
                    lblIGSTPercent.Text = "0.00"
                ElseIf txtGST.Name.ToUpper = "TXTCOMPCESSTYPE" Then
                    lblCompCessPercent.Text = "0.00"
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub txtCGSTType_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCGSTType.Validating, txtSGSTType.Validating, txtUTGSTType.Validating, txtIGSTType.Validating, txtCompCessType.Validating
        Try
            If Not gblnGSTUnit Then Exit Sub
            Dim txtGST As New TextBox
            Dim strTaxType As String = String.Empty
            Dim lblGST As New Label
            Dim txtFocusGST As New TextBox
            txtGST = DirectCast(sender, TextBox)
            If Len(txtGST.Text.Trim) > 0 Then
                If Not ValidationBeforeGSTSelection() Then
                    txtGST.Text = ""
                    txtGST.Focus()
                    Exit Sub
                End If

                If txtGST.Name.ToUpper = "TXTCGSTTYPE" Then
                    strTaxType = "CGST"
                    lblGST = lblCGSTPercent
                    txtFocusGST = txtSGSTType
                ElseIf txtGST.Name.ToUpper = "TXTSGSTTYPE" Then
                    strTaxType = "SGST"
                    lblGST = lblSGSTPercent
                    txtFocusGST = txtIGSTType
                ElseIf txtGST.Name.ToUpper = "TXTUTGSTTYPE" Then
                    strTaxType = "UTGST"
                    lblGST = lblUTGSTPercent
                    txtFocusGST = txtIGSTType
                ElseIf txtGST.Name.ToUpper = "TXTIGSTTYPE" Then
                    strTaxType = "IGST"
                    lblGST = lblIGSTPercent
                    txtFocusGST = txtCompCessType
                ElseIf txtGST.Name.ToUpper = "TXTCOMPCESSTYPE" Then
                    strTaxType = "GSTEC"
                    lblGST = lblCompCessPercent
                    txtFocusGST = Nothing
                End If
                If CheckExistanceOfFieldData((txtGST.Text), "TxRt_Rate_No", "Gen_TaxRate", "Unit_code='" & gstrUnitId & "' and Tx_TaxeID='" & strTaxType & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))") Then
                    lblGST.Text = CStr(GetTaxRate((txtGST.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", "(Unit_code='" & gstrUnitId & "' and Tx_TaxeID='" & strTaxType & "') and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"))
                    If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                        Call SetFormulaofColumns(0, 4)
                        ToShowDatainSummeryGrid()
                    End If
                    If txtFocusGST IsNot Nothing Then
                        If txtFocusGST.Enabled Then txtFocusGST.Focus()
                    End If
                Else
                    Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    txtGST.Text = ""
                    If txtGST.Enabled Then txtGST.Focus()
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub cmdCGSTType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCGSTType.Click, cmdSGSTType.Click, cmdUTGSTType.Click, cmdIGSTType.Click, cmdCompCessType.Click
        Dim strHelp As String
        Dim strGSTHelp() As String
        Dim strTaxType As String = String.Empty
        Dim lblGST As New Label
        Dim cmdGST As New Button
        Dim txtFocusGST As New TextBox
        Try
            If Not gblnGSTUnit Then Exit Sub
            cmdGST = DirectCast(sender, Button)
            Select Case Me.CmdGrpChEnt.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                    If Not ValidationBeforeGSTSelection() Then Exit Sub
                    If cmdGST.Name.ToUpper = "CMDCGSTTYPE" Then
                        strTaxType = "CGST"
                        lblGST = lblCGSTPercent
                        txtFocusGST = txtCGSTType
                    ElseIf cmdGST.Name.ToUpper = "CMDSGSTTYPE" Then
                        strTaxType = "SGST"
                        lblGST = lblSGSTPercent
                        txtFocusGST = txtSGSTType
                    ElseIf cmdGST.Name.ToUpper = "CMDUTGSTTYPE" Then
                        strTaxType = "UTGST"
                        lblGST = lblUTGSTPercent
                        txtFocusGST = txtUTGSTType
                    ElseIf cmdGST.Name.ToUpper = "CMDIGSTTYPE" Then
                        strTaxType = "IGST"
                        lblGST = lblIGSTPercent
                        txtFocusGST = txtIGSTType
                    ElseIf cmdGST.Name.ToUpper = "CMDCOMPCESSTYPE" Then
                        strTaxType = "GSTEC"
                        lblGST = lblCompCessPercent
                        txtFocusGST = txtCompCessType
                    End If
                    strHelp = "Select TxRT_Rate_No,TxRt_Percentage from Gen_taxRate where Unit_code='" & gstrUnitId & "' and Tx_TaxeID ='" & strTaxType & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                    strGSTHelp = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, strTaxType & " Tax Help")
                    If UBound(strGSTHelp) < 0 Then Exit Sub
                    If strGSTHelp(0) = "0" Then
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtFocusGST.Text = "" : txtFocusGST.Focus() : Exit Sub
                    Else
                        If strGSTHelp(0) <> "" Then
                            txtFocusGST.Text = strGSTHelp(0)
                            lblGST.Text = strGSTHelp(1)
                            txtFocusGST.Focus()
                        End If
                    End If
            End Select
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub txtCGSTType_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCGSTType.KeyPress, txtSGSTType.KeyPress, txtUTGSTType.KeyPress, txtIGSTType.KeyPress, txtCompCessType.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        On Error GoTo ErrHandler
        If Not gblnGSTUnit Then Exit Sub
        Dim txtGST As New TextBox
        Dim txtFocusGST As New TextBox
        txtGST = DirectCast(sender, TextBox)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If txtGST.Name.ToUpper = "TXTCGSTTYPE" Then
                            txtFocusGST = txtSGSTType
                        ElseIf txtGST.Name.ToUpper = "TXTSGSTTYPE" Then
                            txtFocusGST = txtIGSTType
                        ElseIf txtGST.Name.ToUpper = "TXTUTGSTTYPE" Then
                            txtFocusGST=txtIGSTType
                        ElseIf txtGST.Name.ToUpper = "TXTIGSTTYPE" Then
                            txtFocusGST = txtCompCessType
                        ElseIf txtGST.Name.ToUpper = "TXTCOMPCESSTYPE" Then
                            txtFocusGST = Nothing
                        End If
                        If Len(txtGST.Text) > 0 Then
                            Call txtCGSTType_Validating(txtGST, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            If txtFocusGST IsNot Nothing Then
                                txtFocusGST.Focus()
                            End If
                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtCGSTType_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCGSTType.KeyUp, txtSGSTType.KeyUp, txtUTGSTType.KeyUp, txtIGSTType.KeyUp, txtCompCessType.KeyUp
        Try
            If Not gblnGSTUnit Then Exit Sub
            Dim KeyCode As Short = e.KeyCode
            Dim Shift As Short = e.KeyData \ &H10000
            Dim txtGST As New TextBox
            Dim cmdGST As New Button
            txtGST = DirectCast(sender, TextBox)
            If KeyCode = 112 Then
                If txtGST.Name.ToUpper = "TXTCGSTTYPE" Then
                    cmdGST = cmdCGSTType
                ElseIf txtGST.Name.ToUpper = "TXTSGSTTYPE" Then
                    cmdGST = cmdSGSTType
                ElseIf txtGST.Name.ToUpper = "TXTUTGSTTYPE" Then
                    cmdGST = cmdUTGSTType
                ElseIf txtGST.Name.ToUpper = "TXTIGSTTYPE" Then
                    cmdGST = cmdIGSTType
                ElseIf txtGST.Name.ToUpper = "TXTCOMPCESSTYPE" Then
                    cmdGST = cmdCompCessType
                End If
                If cmdGST.Enabled Then Call cmdCGSTType_Click(cmdGST, New System.EventArgs())
            End If
            Exit Sub
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub txtCustPartCode_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCustPartCode.Validating
        Dim dtHSN As New DataTable
        Try
            If gblnGSTUnit Then
                Dim strSql As String = String.Empty
                If Len(txtCustPartCode.Text.Trim) > 0 And Len(lblItemCode.Text.Trim) > 0 Then
                    strSql = "SELECT ISNULL(HSN_SAC,'') HSN_SAC,ISNULL(HSN_SAC_CODE,'') HSN_SAC_CODE FROM ITEM_MST WHERE UNIT_CODE='" & gstrUnitId & "' AND ITEM_CODE='" & lblItemCode.Text & "'"
                    dtHSN = SqlConnectionclass.GetDataTable(strSql)
                    If dtHSN IsNot Nothing AndAlso dtHSN.Rows.Count > 0 Then
                        lblHSNSAC.Text = dtHSN.Rows(0)("HSN_SAC")
                        lblHSNSACCODE.Text = dtHSN.Rows(0)("HSN_SAC_CODE")
                    Else
                        lblHSNSAC.Text = ""
                        lblHSNSACCODE.Text = ""
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            dtHSN.Dispose()
        End Try
    End Sub
    Private Function ValidationBeforeGSTSelection() As Boolean
        Dim result As Boolean = True
        If Len(txtLocationCode.Text.Trim) = 0 Then
            MsgBox("Please Select Location Code.", MsgBoxStyle.Information, "eMPro")
            result = False
        ElseIf Len(txtCustCode.Text.Trim) = 0 Then
            MsgBox("Please Select Customer Code.", MsgBoxStyle.Information, "eMPro")
            result = False
        End If
        Return result
    End Function
    Private Function ValidateGSTTaxes() As Boolean
        Dim sqlCmd As SqlCommand
        Dim strMsg As String = String.Empty
        Dim result As Boolean = True
        Try
            If Not gblnGSTUnit Then
                Return result
                Exit Function
            End If
            If Len(txtCustCode.Text.Trim) > 0 Then
                sqlCmd = New SqlCommand()
                sqlCmd.CommandType = CommandType.StoredProcedure
                sqlCmd.CommandText = "USP_VALIDATE_GST_SUPPLIMENTARY_INVOICE"
                sqlCmd.Parameters.AddWithValue("@CUSTOMER_CODE", txtCustCode.Text.Trim)
                sqlCmd.Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
                sqlCmd.Parameters.Add("@CGST_APPLICABLE", SqlDbType.Bit).Value = 0
                sqlCmd.Parameters("@CGST_APPLICABLE").Direction = ParameterDirection.InputOutput
                sqlCmd.Parameters.Add("@SGST_APPLICABLE", SqlDbType.Bit).Value = 0
                sqlCmd.Parameters("@SGST_APPLICABLE").Direction = ParameterDirection.InputOutput
                sqlCmd.Parameters.Add("@UTGST_APPLICABLE", SqlDbType.Bit).Value = 0
                sqlCmd.Parameters("@UTGST_APPLICABLE").Direction = ParameterDirection.InputOutput
                sqlCmd.Parameters.Add("@IGST_APPLICABLE", SqlDbType.Bit).Value = 0
                sqlCmd.Parameters("@IGST_APPLICABLE").Direction = ParameterDirection.InputOutput
                sqlCmd.Parameters.Add("@COMPCESS_APPLICABLE", SqlDbType.Bit).Value = 0
                sqlCmd.Parameters("@COMPCESS_APPLICABLE").Direction = ParameterDirection.InputOutput
                sqlCmd.Parameters.Add("@ERROR_MSG", SqlDbType.VarChar, 8000).Value = String.Empty
                sqlCmd.Parameters("@ERROR_MSG").Direction = ParameterDirection.InputOutput
                SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                _cgstApplicable = CBool(sqlCmd.Parameters("@CGST_APPLICABLE").Value)
                _sgstApplicable = CBool(sqlCmd.Parameters("@SGST_APPLICABLE").Value)
                _utgstApplicable = CBool(sqlCmd.Parameters("@UTGST_APPLICABLE").Value)
                _igstApplicable = CBool(sqlCmd.Parameters("@IGST_APPLICABLE").Value)
                _ccessApplicable = CBool(sqlCmd.Parameters("@COMPCESS_APPLICABLE").Value)
                If Len(sqlCmd.Parameters("@ERROR_MSG").Value) > 0 Then
                    MsgBox(sqlCmd.Parameters("@ERROR_MSG").Value.ToString(), MsgBoxStyle.Information, "eMPro")
                    result = False
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
            result = False
        End Try
        Return result
    End Function
    '101188073
    Private Sub txtCustCode_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCustCode.Validating
        EnableDisableGST()
    End Sub
    Private Sub EnableDisableGST()
        Try
            If Not gblnGSTUnit Then Exit Sub
            If Len(txtCustCode.Text.Trim) > 0 Then
                If Not ValidateGSTTaxes() Then
                    txtCGSTType.Enabled = False
                    txtCGSTType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    lblCGSTPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    cmdCGSTType.Enabled = False
                    txtSGSTType.Enabled = False
                    cmdSGSTType.Enabled = False
                    txtSGSTType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    lblSGSTPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtUTGSTType.Enabled = False
                    cmdUTGSTType.Enabled = False
                    txtUTGSTType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    lblUTGSTPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtIGSTType.Enabled = False
                    cmdIGSTType.Enabled = False
                    txtIGSTType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    lblIGSTPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtCompCessType.Enabled = False
                    cmdCompCessType.Enabled = False
                    txtCompCessType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    lblCompCessPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    Exit Sub
                Else
                    txtCGSTType.Enabled = _cgstApplicable
                    cmdCGSTType.Enabled = _cgstApplicable
                    If _cgstApplicable Then
                        txtCGSTType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        lblCGSTPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    Else
                        txtCGSTType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        lblCGSTPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    End If
                    txtSGSTType.Enabled = _sgstApplicable
                    cmdSGSTType.Enabled = _sgstApplicable
                    If _sgstApplicable Then
                        txtSGSTType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        lblSGSTPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    Else
                        txtSGSTType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        lblSGSTPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    End If
                    txtUTGSTType.Enabled = _utgstApplicable
                    cmdUTGSTType.Enabled = _utgstApplicable
                    If _utgstApplicable Then
                        txtUTGSTType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        lblUTGSTPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    Else
                        txtUTGSTType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        lblUTGSTPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    End If
                    txtIGSTType.Enabled = _igstApplicable
                    cmdIGSTType.Enabled = _igstApplicable
                    If _igstApplicable Then
                        txtIGSTType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        lblIGSTPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    Else
                        txtIGSTType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        lblIGSTPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    End If
                    txtCompCessType.Enabled = _ccessApplicable
                    cmdCompCessType.Enabled = _ccessApplicable
                    If _ccessApplicable Then
                        txtCompCessType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        lblCompCessPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    Else
                        txtCompCessType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        lblCompCessPercent.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
End Class