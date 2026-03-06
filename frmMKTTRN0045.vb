Option Strict Off
Option Explicit On
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Friend Class frmMKTTRN0045
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
	'--------------------------------------------------------------------------------
	' Revision Date     :   25/05/2006
	' Revision By       :   Davinder Singh
	' Issue ID          :   17936
	' Revision History  :   1) To save the data also in the Supplementary_Dtl table
	'                       2) To calculate the sales Tax on the basis of Basic rate Instead of calculating it on the Accesible rate
	'                       3) To calculate the Total Amount on the basis of Basic rate Instead of calculating it on the Accesible rate
	'--------------------------------------------------------------------------------
	' Author              - Davinder Singh
	' Revision Date       - 20 Apr 2007
	' Revision History    - To include the Secondary Ecess(SECESS) in Supplementary Invoice
	' Issue ID            - 19786
	'--------------------------------------------------------------------------------
	'Revised By           : Manoj Kr. Vaish
	'Issue ID             : 21551
	'Revision Date        : 22-Nov-2007
	'History              : Add New Tax VAT with Sale Tax help
    '***********************************************************************************
    'Revised By           : Manoj Kr. Vaish
    'Issue ID             : eMpro-20090401-29562
    'Revision Date        : 28-Apr-2009
    'History              : Wrong total amount was saving while making supplementary Invoice 
    '***********************************************************************************
    'Revised By           : Siddharth Ranjan
    'Issue ID             : eMpro-20090910-36205
    'Revision Date        : 10-Sep-2009
    'History              : Add Additional VAT functionality
    '---------------------------------------------------------------------------------------
    'Modified by    :   Virendra Gupta
    'Modified ON    :   18/05/2011
    'Modified to support MultiUnit functionality
    '-----------------------------------------------------------------------
    '***********************************************************************************
    'Revised By           : Prashant Rajpal
    'Issue ID             : 10125156
    'Revision Date        : 10-Aug-2011
    'History              : Error message :Input String was not in correct Format
    '***********************************************************************************
    'MODIFIED BY VIRENDRA GUPTA 0N 09 NOV 2011 FOR CHANGE MANAGEMENT
    '***********************************************************************************
    'Revised By           : Prashant Rajpal
    'Issue ID             : 10158545
    'Revision Date        : 11-Nov-2011
    'History              : Supplementary Invoice , No validation on Customer item Code Field 
    '***********************************************************************************
    'Modified By Roshan Singh on 20 Dec 2011 for multiUnit change management  

    Dim mintIndex As Short
	Dim Financial_Start_Date As Date
	Dim Financial_End_Date As Date
	Dim strDefaultLocation As String
    Private strSelInvoices As String 'to store the selected invoice numbers
	Private strNotSelInvoices As String ' to store the deselected invoice numbers
	Private bCheck As Boolean 'boolean to specify the class function if invoice inclusion check is required
	Private blnInclude As Boolean ' retval of get invoice summary function from the dll
    Dim strInvoiceNo As String
	Dim cColor As System.Drawing.Color
    Dim mblnExciseRoundOFFFlag As Boolean
	Dim mblnEcessRoundOFFFlag As Boolean
	Dim mblnTotalInvoiceRoundOFFFlag As Boolean
	Dim mblnsalestaxRoundOFFFlag As Boolean
	Dim intExciseRoundOFFplace As Short
	Dim intEcessRoundOFFplace As Short
	Dim intTotalInvoiceRoundOFFplace As Short
    Dim intsalestaxRoundOFFplace As Short
    Dim mblnISSurChargeTaxRoundOff As Boolean
    Dim stSSTRoundOffDecimal As Short
    Private Sub CmdChallanNo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdChallanNo.Click
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Display Help From SalesChallan_Dtl
        '*****************************************************************************************
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
                strHelpString = "Select DISTINCT Doc_No,Location_Code from SupplementaryInv_hdr where Location_code ='" & txtLocationCode.Text & "' and Unit_Code = '" & gstrUNITID & "'"
                strHelpString = Trim(strHelpString) & " and Cancel_flag = 0"
        End Select
        strChallanNo = ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelpString, "Supplementary Invoice No")
        If UBound(strChallanNo) < 0 Then Exit Sub
        If strChallanNo(0) = "0" Then
            MsgBox("No Supplementary Invoice Available To Display", MsgBoxStyle.Information, ResolveResString(100)) : txtChallanNo.Text = "" : txtChallanNo.Focus() : Exit Sub
        Else
            txtChallanNo.Text = strChallanNo(0)
            Call txtChallanNo_Validating(txtChallanNo, New System.ComponentModel.CancelEventArgs(False))
            CmdGrpChEnt.Revert()
            CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        End If
        txtChallanNo.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub CmdCustCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdCustCodeHelp.Click
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Display Help From Customer_Mst
        '*****************************************************************************************
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
                strCustHelp = "Select DISTINCT customer_code,Cust_Name from customer_mst where Unit_Code = '" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
        End Select
        strCustomer = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strCustHelp, "Customer List")
        If UBound(strCustomer) <= 0 Then Exit Sub
        If strCustomer(0) = "0" Then
            MsgBox("No Customer Available to Display", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100)) : txtCustCode.Text = "" : txtCustCode.Focus() : Exit Sub
        Else
            txtCustCode.Text = strCustomer(0)
            lblCustCodeDes.Text = strCustomer(1)
        End If
        txtCustPartCode.Text = ""
        txtCustPartCode.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdcustPartCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdcustPartCode.Click
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Display Help For Customer Item Code
        '*****************************************************************************************
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
            MsgBox("First Select Customer Code.", MsgBoxStyle.Information, ResolveResString(100))
            If txtCustCode.Enabled Then txtCustCode.Focus()
            Exit Sub
        End If
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                'ISSUE ID Starts: 10158545 
                'strCustItem = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrDSNName, gstrDatabaseName, "Select DISTINCT Cust_DRgNo,Item_code,Drg_desc from custitem_mst")
                strCustItem = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select DISTINCT Cust_DRgNo,Item_code,Drg_desc from custitem_mst where account_code='" & Me.txtCustCode.Text & "' and active=1 and Unit_Code = '" & gstrUNITID & "'")
                'ISSUE ID End : 10158545 
        End Select
        If UBound(strCustItem) <= 0 Then Exit Sub
        txtCustPartCode.Text = strCustItem(0)
        lblItemCode.Text = strCustItem(1)
        lblCustItemDesc.Text = strCustItem(2)
        If Len(Trim(txtCustCode.Text)) > 0 Then
            rsCustItemMst = New ClsResultSetDB
            strCustItemMst = "SELECT Bill_Address1 + ', '  +  Bill_Address2 + ', ' + Bill_City + ' - ' + Bill_Pin as  invoiceAddress from Customer_Mst where Customer_code ='" & txtCustCode.Text & "' and Unit_Code = '" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
            rsCustItemMst.GetResult(strCustItemMst)
            If rsCustItemMst.GetNoRows > 0 Then
                lblAddressDes.Text = rsCustItemMst.GetValue("InvoiceAddress")
            End If
            rsCustItemMst = Nothing
        End If
        Me.txtRefNo.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdEcssCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdEcssCodeHelp.Click
        '****************************************************************
        ' Author              - Sourabh Khatri
        ' Created Date        - 13 oct 2004
        ' Arguments           - None
        ' Return Value        - None
        ' Function            - To display help for Ecss Type
        '*****************************************************************************************
        On Error GoTo Errorhandler
        Dim strHelp As String
        Dim strSSTaxHelp() As String
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strHelp = "Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where tx_TaxeID ='ECS' and Unit_code = '" & gstrUNITID & "'"
                strSSTaxHelp = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, "E CESS Tax Help")
                If UBound(strSSTaxHelp) <= 0 Then Exit Sub
                If strSSTaxHelp(0) = "0" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : Me.txtEcssCode.Text = "" : Me.txtEcssCode.Focus() : Exit Sub
                Else
                    Me.txtEcssCode.Text = strSSTaxHelp(0)
                    Me.lblEcssCode.Text = strSSTaxHelp(1)
                    txtEcssCode.Focus()
                End If
        End Select
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub CmdExciseTaxType_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdExciseTaxType.Click
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Display Help From SaleTax Master
        '*****************************************************************************************
        Dim strHelp As String
        Dim strSTaxHelp() As String
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strHelp = "Select Distinct TxRT_Rate_No,TxRt_Percentage from Gen_taxRate where Tx_TaxeID ='EXC' and Unit_code = '" & gstrUNITID & "'"
                strSTaxHelp = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, "Exc.Tax Help")
                If UBound(strSTaxHelp) <= 0 Then Exit Sub
                If strSTaxHelp(0) = "0" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtExciseTaxType.Text = "" : txtExciseTaxType.Focus() : Exit Sub
                Else
                    txtExciseTaxType.Text = strSTaxHelp(0)
                    lblExctax_Per.Text = strSTaxHelp(1)
                    txtExciseTaxType.Focus()
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdPrint_Click()
        On Error GoTo ErrHandler
        Dim strsql As String
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
        '<<<<CR11 Code Starts>>>>
        Dim objRpt As ReportDocument
        Dim frmReportViewer As New eMProCrystalReportViewer
        objRpt = frmReportViewer.GetReportDocument()
        frmReportViewer.ShowPrintButton = True
        frmReportViewer.ShowTextSearchButton = True
        frmReportViewer.ShowZoomButton = True
        frmReportViewer.ReportHeader = Me.ctlFormHeader1.HeaderString()
        '<<<<CR11 Code Ends>>>>
        With objRpt
            'load the report
            .Load(My.Application.Info.DirectoryPath & "\Reports\rptInvoiceDetails.rpt")
            ' Checking Number of Checked records
            strsupplierlist = txtCustCode.Text
            strsupplierlist = "{saleschallan_dtl.account_code} IN ('" & strsupplierlist & "') and "
            strLocationList = strLocationList & strsupplierlist
            ' Checking for Date Range
            strLocationList = strLocationList & "{saleschallan_dtl.invoice_date} >= date(" & strdate & ")  and "
            strLocationList = strLocationList & "{saleschallan_dtl.invoice_date}  <= date(" & strdate & ")"
            ' To Check whether the User is Viewing all Records all Selected Records
            straddress = gstr_WRK_ADDRESS1 & " " & gstr_WRK_ADDRESS2
            .DataDefinition.FormulaFields("comp_name").Text = "'" & gstrCOMPANY & "'"
            .DataDefinition.FormulaFields("comp_address").Text = "'" & straddress & "'"
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            .RecordSelectionFormula = strLocationList & " and {saleschallan_dtl.UNIT_CODE} = '" & gstrUNITID & "'"
            frmReportViewer.Zoom = 150
            frmReportViewer.Show()
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            Exit Sub
        End With
        Exit Sub
ErrHandler:
        If Err.Number = 20545 Then
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            Resume Next
        Else
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End If
    End Sub
    Private Sub CmdSaleTaxType_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdSaleTaxType.Click
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Display Help From SaleTax Master
        '*****************************************************************************************
        Dim strHelp As String
        Dim strSTaxHelp() As String
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strHelp = "Select TxRT_Rate_No,TxRt_Percentage from Gen_taxRate where ( Tx_TaxeID ='CST' OR Tx_TaxeID ='LST' OR Tx_TaxeID ='VAT') and Unit_code = '" & gstrUNITID & "'"
                strSTaxHelp = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, "S.Tax Help")
                If UBound(strSTaxHelp) <= 0 Then Exit Sub
                If strSTaxHelp(0) = "0" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtSaleTaxType.Text = "" : txtSaleTaxType.Focus() : Exit Sub
                Else
                    txtSaleTaxType.Text = strSTaxHelp(0)
                    lblSaltax_Per.Text = strSTaxHelp(1)
                    txtSaleTaxType.Focus()
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdSEcssCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSEcssCodeHelp.Click
        '****************************************************************
        ' Author              - Davinder Singh
        ' Created Date        - 20 Apr 2007
        ' Function            - To display help for Secondary Ecss
        ' Issue ID            - 19786 ETo include the Secondary Ecess in Supplementary Invoice
        '*****************************************************************************************
        On Error GoTo Errorhandler
        Dim strHelp As String
        Dim strSSTaxHelp() As String
        Select Case CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strHelp = "Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where tx_TaxeID ='ECSSH' and Unit_code = '" & gstrUNITID & "'"
                strSSTaxHelp = ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, "E CESS Tax Help")
                If UBound(strSSTaxHelp) <= 0 Then Exit Sub
                If strSSTaxHelp(0) = "0" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    txtSEcssCode.Text = ""
                    txtSEcssCode.Focus()
                    Exit Sub
                Else
                    lblSEcssCode.Text = strSSTaxHelp(1)
                    txtSEcssCode.Text = strSSTaxHelp(0)
                    txtEcssCode.Focus()
                End If
        End Select
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAmendment_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAmendment.TextChanged
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - On Change of Amendment change
        '*****************************************************************************************
        On Error GoTo ErrHandler
        txtAmendment.Text = Replace(txtAmendment.Text, "'", "")
        If Len(Trim(txtAmendment.Text)) = 0 Then
            txtRefNo.Text = ""
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtAmendment_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAmendment.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Set focus on Next Control
        '*****************************************************************************************
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
    Private Sub txtChallanNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtChallanNo.TextChanged
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Refresh all the details
        '*****************************************************************************************
        On Error GoTo ErrHandler
        txtChallanNo.Text = Replace(txtChallanNo.Text, "'", "")
        If Len(Trim(txtChallanNo.Text)) = 0 Then
            Call RefreshCtrls()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtChallanNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtChallanNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim rsSalesChallan As New ClsResultSetDB
        Dim strField As String
        strField = "SELECT CM.BILL_ADDRESS1 + ', '  +  CM.BILL_ADDRESS2 + ', ' + CM.BILL_CITY + ' - ' + CM.BILL_PIN AS  INVOICEADDRESS, isnull(CIM.DRG_DESC,'') as DRG_DESC, SH.* FROM SUPPLEMENTARYINV_HDR SH LEFT OUTER JOIN CUSTOMER_MST CM ON CM.CUSTOMER_CODE=SH.ACCOUNT_CODE AND CM.UNIT_CODE = SH.UNIT_CODE LEFT OUTER JOIN CUSTITEM_MST CIM ON SH.ITEM_CODE=CIM.ITEM_CODE AND SH.CUST_ITEM_CODE=CIM.CUST_DRGNO AND SH.UNIT_CODE = CIM.UNIT_CODE Where SH.Doc_No = '" & Trim(txtChallanNo.Text) & "' and cancel_flag = 0 AND SH.UNIT_CODE = '" & gstrUNITID & "'"
        rsSalesChallan.GetResult(strField)
        If rsSalesChallan.GetNoRows > 0 Then
            txtCustCode.Text = Trim(rsSalesChallan.GetValue("Account_Code"))
            lblCustCodeDes.Text = Trim(rsSalesChallan.GetValue("Cust_Name"))
            txtRefNo.Text = Trim(rsSalesChallan.GetValue("Cust_Ref"))
            txtAmendment.Text = Trim(rsSalesChallan.GetValue("Amendment_No"))
            lblItemCode.Text = rsSalesChallan.GetValue("Item_code")
            txtCustPartCode.Text = Trim(rsSalesChallan.GetValue("Cust_Item_Code"))
            lblCustItemDesc.Text = rsSalesChallan.GetValue("Drg_desc")
            lblCurrencyDes.Text = Trim(rsSalesChallan.GetValue(" Currency_Code"))
            ctlAccrate.Text = Trim(rsSalesChallan.GetValue("Rate"))
            ctlSalesQuantity.Text = Val(rsSalesChallan.GetValue("Accessible_amount")) / IIf(Val(Trim(rsSalesChallan.GetValue("Rate"))) = 0, 1, Val(Trim(rsSalesChallan.GetValue("Rate"))))
            ctlAccrate.Text = Val(CStr((rsSalesChallan.GetValue("Accessible_amount")) / Val(ctlSalesQuantity.Text)))
            txtExciseTaxType.Text = rsSalesChallan.GetValue("Excise_Type")
            lblExctax_Per.Text = rsSalesChallan.GetValue("Excise_per")
            txtSaleTaxType.Text = rsSalesChallan.GetValue("SalesTax_Type")
            lblSaltax_Per.Text = rsSalesChallan.GetValue("SalesTax_Per")
            ctlTotals.Text = CStr(Val(rsSalesChallan.GetValue("total_amount")))
            txtCustRefRemarks.Text = rsSalesChallan.GetValue("SuppInv_Remarks")
            txtRemarks.Text = rsSalesChallan.GetValue("remarks")
            txtEcssCode.Text = rsSalesChallan.GetValue("ECESS_Type")
            lblEcssCode.Text = rsSalesChallan.GetValue("ECESS_Per")
            txtSEcssCode.Text = rsSalesChallan.GetValue("SECESS_Type")
            lblSEcssCode.Text = rsSalesChallan.GetValue("SECESS_Per")
            '--------------------------------------------10125156-----------------------------------------------------------------------------------
            If Val(rsSalesChallan.GetValue("sales_quantity")) > 0 Then
                ctlBasicRate.Text = System.Math.Round(rsSalesChallan.GetValue("basic_amount") / rsSalesChallan.GetValue("sales_quantity"), 4)
            Else
                ctlBasicRate.Text = 0
            End If
            '-----------------------------------------------------------------------------------------------------------------------------------------
            lblAddressDes.Text = rsSalesChallan.GetValue("INVOICEADDRESS")
            ctlSalesQuantity.Text = rsSalesChallan.GetValue("sales_Quantity")
            txtAddVAT.Text = rsSalesChallan.GetValue("ADDVAT_TYPE")
            lblAddress.Text = rsSalesChallan.GetValue("ADDVAT_PER")
            txtSurcharge_VAT.Text = rsSalesChallan.GetValue("SURCHARGE_SALESTAXTYPE")
            lblSurcharge_VAT.Text = rsSalesChallan.GetValue("SURCHARGE_SALESTAX_PER")
            chkShowDetailAnexture.Enabled = True
            optsupp.Enabled = True
            optcredit.Enabled = True
        End If
        rsSalesChallan.ResultSetClose()
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtCustCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustCode.TextChanged
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Refresh all the details
        '*****************************************************************************************
        On Error GoTo ErrHandler
        txtCustCode.Text = Replace(txtCustCode.Text, "'", "")
        If Len(Trim(txtCustCode.Text)) = 0 Then
            lblCustCodeDes.Text = "" : lblAddressDes.Text = ""
            txtCustPartCode.Text = "" : txtRefNo.Text = "" : txtAmendment.Text = ""
            txtCustRefRemarks.Text = "" : txtRemarks.Text = "" : txtExciseTaxType.Text = ""
            txtSaleTaxType.Text = ""
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtCustCode_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustCode.Leave
        On Error GoTo ErrHandler
        Dim strCustMst As String
        Dim rsCustMst As ClsResultSetDB
        If Len(Trim(txtCustCode.Text)) > 0 Then
            rsCustMst = New ClsResultSetDB
            strCustMst = "SELECT isnull(Bill_Address1,'') + ', '  +  isnull(Bill_Address2,'') + ', ' + isnull(Bill_City,'') + ' - ' + isnull(Bill_Pin,'') as  invoiceAddress,isnull(Currency_Code,'INR') as Currency_Code from Customer_Mst where Customer_code ='" & txtCustCode.Text & "' and Unit_Code = '" & gstrUNITID & "'  and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
            rsCustMst.GetResult(strCustMst)
            If rsCustMst.GetNoRows > 0 Then
                lblAddressDes.Text = rsCustMst.GetValue("InvoiceAddress")
                lblCurrencyDes.Text = rsCustMst.GetValue("Currency_Code")
            End If
            rsCustMst = Nothing
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtCustPartCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustPartCode.TextChanged
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Refresh All the Details
        '*****************************************************************************************
        On Error GoTo ErrHandler
        txtCustPartCode.Text = Replace(txtCustPartCode.Text, "'", "")
        If Len(Trim(txtCustPartCode.Text)) = 0 Then
            lblItemCode.Text = "" : txtRefNo.Text = "" : txtAmendment.Text = ""
            txtCustRefRemarks.Text = "" : txtRemarks.Text = "" : txtExciseTaxType.Text = ""
            txtSaleTaxType.Text = "" : lblCustItemDesc.Text = ""
        End If
        strSelInvoices = ""
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtCustPartCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustPartCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call help on F1 Press
        '*****************************************************************************************
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
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Set focus on Next Control
        '*****************************************************************************************
        On Error GoTo ErrHandler
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
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
    Private Sub CmdGrpChEnt_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles CmdGrpChEnt.ButtonClick
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To ADD Functionality of ADD/EDIT/SAVE/UPDATE/CLOSE
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Dim rsCustOrdHdr As ClsResultSetDB
        Dim strsql As String
        Dim strToCheckPrevInvoices As String
        Dim rsToCheckPrevInvoices As ClsResultSetDB
        Dim intCase As String
        Dim strMessageString As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim strInputBox As String
        Select Case CmdGrpChEnt.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                Call SetFinancialYearDates()
                Call EnableControls(True, Me, True)
                Call SelectChallanNoFromSupplementatryInvHdr()
                txtChallanNo.Enabled = False : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : CmdChallanNo.Enabled = False
                txtChallanNo.Enabled = False : lblCustCodeDes.Text = ""
                txtLocationCode.Text = strDefaultLocation
                txtLocationCode.Focus()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                Call frmMKTTRN0045_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                If ValidatebeforeSave() = False Then Exit Sub
                Call SaveData()
                Me.CmdGrpChEnt.Revert()
                CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                gblnCancelUnload = False : gblnFormAddEdit = False
                Call EnableControls(False, Me)
                txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                CmdLocCodeHelp.Enabled = True : CmdChallanNo.Enabled = True 'Me.SpChEntry.Enabled = False
                If txtLocationCode.Enabled Then
                    txtLocationCode.Focus()
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
                If ConfirmWindow(10054, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    rsCustOrdHdr = New ClsResultSetDB
                    strsql = "select * from SupplementaryInv_hdr where  doc_no='" & txtChallanNo.Text & "' and cancel_flag=" & "0" & " and bill_flag=1 and Unit_Code = '" & gstrUNITID & "'"
                    rsCustOrdHdr.GetResult(strsql)
                    If rsCustOrdHdr.GetNoRows > 0 Then
                        MsgBox("Locked Invoices can't be deleted", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        Exit Sub
                    Else
                        mP_Connection.Execute("Delete SupplementaryInv_hdr where doc_no='" & txtChallanNo.Text & "' and bill_flag<>1 and Unit_Code = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute("Delete SupplementaryInv_dtl where doc_no='" & txtChallanNo.Text & "' and Unit_Code = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        Call EnableControls(False, Me, True)
                        txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        CmdLocCodeHelp.Enabled = True : CmdChallanNo.Enabled = True 'Me.SpChEntry.Enabled = False
                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                        CmdGrpChEnt.Enabled(2) = False
                        If txtLocationCode.Enabled Then
                            txtLocationCode.Focus()
                        End If
                    End If
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                MsgBox("You can Delete the Record but can't Update", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                Me.CmdGrpChEnt.Revert()
                CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                Exit Sub
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
                '<<<<CR11 Code Starts>>>>
                Dim objRpt As ReportDocument
                Dim strRepPath As String
                Dim frmReportViewer As New eMProCrystalReportViewer
                objRpt = frmReportViewer.GetReportDocument()
                frmReportViewer.ShowPrintButton = True
                frmReportViewer.ShowTextSearchButton = True
                frmReportViewer.ShowZoomButton = True
                frmReportViewer.ReportHeader = Me.ctlFormHeader1.HeaderString()
                '<<<<CR11 Code Ends>>>>
                With objRpt
                    If optsupp.Checked = True Then
                        If chkShowDetailAnexture.CheckState = System.Windows.Forms.CheckState.Checked Then
                            strRepPath = My.Application.Info.DirectoryPath & "\Reports\rptSuppInvAnnexure.rpt"
                        Else
                            strRepPath = My.Application.Info.DirectoryPath & "\Reports\rptSuppInvAnnexureSummary.rpt"
                        End If
                    Else
                        strRepPath = My.Application.Info.DirectoryPath & "\Reports\rptSuppCrAdvise.rpt"
                    End If
                    'load the report
                    .Load(strRepPath)
                    strsql = "{SupplementaryInv_hdr.Location_Code}='" & Trim(txtLocationCode.Text) & "' and {SupplementaryInv_hdr.Doc_No} =" & Trim(txtChallanNo.Text)
                    .RecordSelectionFormula = strsql & " and {SupplementaryInv_hdr.UNIT_CODE} = '" & gstrUNITID & "'"
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                    frmReportViewer.SetReportDocument()
                    objRpt.PrintToPrinter(1, False, 0, 0)
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                End With
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub CmdLocCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdLocCodeHelp.Click
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Display Help From Location Master
        '*****************************************************************************************
        Dim strLocationCode() As String
        On Error GoTo ErrHandler
        Select Case CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                strLocationCode = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select DISTINCT s.Location_Code,l.Description from Location_Mst l,SaleConf s where s.Location_code = l.Location_code and s.Unit_code = l.Unit_code and s.Location_code like'" & txtLocationCode.Text & "%' and s.Unit_code = '" & gstrUNITID & "'", "Accounting Locations")
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                strLocationCode = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select DISTINCT s.Location_Code,l.Description from Location_Mst l,SupplementaryInv_hdr s where s.Location_code = l.Location_code and s.Unit_code = l.Unit_code and s.Location_code like'" & txtLocationCode.Text & "%' and s.Unit_code = '" & gstrUNITID & "'", "Accounting Locations")
        End Select
        If UBound(strLocationCode) < 0 Then Exit Sub
        If strLocationCode(0) = "0" Then
            MsgBox("No Accounting Location Available to Display.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100)) : txtLocationCode.Text = "" : txtLocationCode.Focus() : Exit Sub
        Else
            txtLocationCode.Text = strLocationCode(0)
            strDefaultLocation = Trim(txtLocationCode.Text)
        End If
        'Call SelectDescriptionForField("Description", "Location_Code", "Location_Mst", lblLocCodeDes, txtLocationCode.Text)
        Call txtLocationCode_Leave(txtLocationCode, New System.EventArgs())
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0045_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add Code on Form Activate
        '*****************************************************************************************
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex
        frmModules.NodeFontBold(Me.Tag) = True
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0045_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To ADD Code Form Deactivate
        '*****************************************************************************************
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0045_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call empower help on F4 Click
        '*****************************************************************************************
        On Error GoTo Err_Handler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0045_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add Code on Form Key Press
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                If Me.CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call Me.CmdGrpChEnt.Revert()
                        Call EnableControls(False, Me, True)
                        gblnCancelUnload = False
                        gblnFormAddEdit = False
                        txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdLocCodeHelp.Enabled = True
                        txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdChallanNo.Enabled = True
                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                        CmdGrpChEnt.Enabled(2) = False
                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
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
    Private Sub frmMKTTRN0045_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add The Code on Form Load
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Dim rsCompanyMst As ClsResultSetDB
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FillLabelFromResFile(Me) 'Fill Labels >From Resource File
        Call FitToClient(Me, FraChEnt, ctlFormHeader1, CmdGrpChEnt)
        Call EnableControls(False, Me, True)
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
        CmdExciseTaxType.Image = My.Resources.ico111.ToBitmap
        CmdSaleTaxType.Image = My.Resources.ico111.ToBitmap
        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        CmdGrpChEnt.Enabled(2) = False
        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        strDefaultLocation = ""
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0045_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add The Code on Quaery Unload
        '*****************************************************************************************
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
    Private Sub frmMKTTRN0045_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add Code on Form Unload
        '*****************************************************************************************
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Me.Dispose()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SelectDescriptionForField(ByRef pstrFieldName1 As String, ByRef pstrFieldName2 As String, ByRef pstrTableName As String, ByRef pContrName As System.Windows.Forms.Control, ByRef pstrControlText As String)
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           -  pstrFieldName1 - Field Name1,pstrFieldName2 - Field Name2,pstrTableName - Table Name
        '                    -  pContName - Name Of The Control where Caption Is To Be Set
        '                    -  pstrControlText - Field Text
        'Return Value        - None
        'Function            - To Select The Field Description In The Description Labels
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Dim strDesSql As String 'Declared to make Select Query
        Dim rsDescription As ClsResultSetDB
        strDesSql = "Select " & Trim(pstrFieldName1) & " from " & Trim(pstrTableName) & " where " & Trim(pstrFieldName2) & "='" & Trim(pstrControlText) & "'"
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
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Set focus on Next Control
        '*****************************************************************************************
        On Error GoTo ErrHandler
        If KeyAscii = Asc("'") Then KeyAscii = 0
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            Call txtChallanNo_Validating(txtChallanNo, New System.ComponentModel.CancelEventArgs(False))
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
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Display on F1 Click
        '*****************************************************************************************
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call CmdChallanNo_Click(CmdChallanNo, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtCustCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Set focus on Next Control
        '*****************************************************************************************
        On Error GoTo ErrHandler
        If KeyAscii = Asc("'") Then KeyAscii = 0
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
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
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Display help on F1 click
        '*****************************************************************************************
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call CmdCustCodeHelp_Click(CmdCustCodeHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtCustRefRemarks_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustRefRemarks.TextChanged
        On Error GoTo ErrHandler
        txtCustRefRemarks.Text = Replace(txtCustRefRemarks.Text, "'", "")
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
    Private Sub txtEcssCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEcssCode.TextChanged
        '*****************************************************************************************
        'Author              - Sourabh Khatri
        'Function            - To Refresh Ecss Label
        '*****************************************************************************************
        On Error GoTo ErrHandler
        txtEcssCode.Text = Replace(txtEcssCode.Text, "'", "")
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
        '*****************************************************************************************
        'Author              - Sourabh Khatri
        'Function            - To Call Help on F1 Click
        '*****************************************************************************************
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
        If KeyAscii = Asc("'") Then KeyAscii = 0
        If KeyAscii = 13 Then txtEcssCode_Validating(txtEcssCode, New System.ComponentModel.CancelEventArgs((False)))
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtEcssCode_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEcssCode.Leave
        On Error GoTo ErrHandler
        Call CalculateTotalIncoiceValue()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtEcssCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtEcssCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If Trim(Me.txtEcssCode.Text) <> "" Then
            If CheckExistanceOfFieldData((Me.txtEcssCode.Text), "TxRt_Rate_No", "Gen_TaxRate", " Tx_TaxeID='ECS'") Then
                Me.lblEcssCode.Text = CStr(GetTaxRate((Me.txtEcssCode.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='ECS'"))
                txtSEcssCode.Focus()
            Else
                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtEcssCode.Text = ""
                If txtEcssCode.Enabled Then txtEcssCode.Focus()
            End If
        Else
            ''ctlTotals.SetFocus
            CmdGrpChEnt.Focus()
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtExciseTaxType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtExciseTaxType.TextChanged
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Refresh all the Details
        '*****************************************************************************************
        On Error GoTo ErrHandler
        txtExciseTaxType.Text = Replace(txtExciseTaxType.Text, "'", "")
        If Len(txtExciseTaxType.Text) = 0 Then
            lblExctax_Per.Text = "0.00"
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtExciseTaxType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtExciseTaxType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - At Enter Key Press Set Focus To Next Control
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Call txtExciseTaxType_Validating(txtExciseTaxType, New System.ComponentModel.CancelEventArgs(False))
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
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To show the help on F1 key Press
        '*****************************************************************************************
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdExciseTaxType.Enabled Then Call CmdExciseTaxType_Click(CmdExciseTaxType, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtExciseTaxType_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtExciseTaxType.Leave
        On Error GoTo ErrHandler
        Call CalculateTotalIncoiceValue()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtExciseTaxType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtExciseTaxType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Validate Excise Tax Type
        '*****************************************************************************************
        On Error GoTo ErrHandler
        If (CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT) Then
            If Len(txtExciseTaxType.Text) > 0 Then
                If CheckExistanceOfFieldData((txtExciseTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='EXC')") Then
                    lblExctax_Per.Text = CStr(GetTaxRate((txtExciseTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='EXC')"))
                    If txtSaleTaxType.Enabled Then txtSaleTaxType.Focus()
                Else
                    Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    txtExciseTaxType.Text = ""
                    Cancel = True
                    ''If txtExciseTaxType.Enabled Then txtExciseTaxType.SetFocus
                End If
            Else
                txtSaleTaxType.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtLocationCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLocationCode.TextChanged
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Refresh all the data
        '*****************************************************************************************
        On Error GoTo ErrHandler
        txtLocationCode.Text = Replace(txtLocationCode.Text, "'", "")
        If Len(Trim(txtLocationCode.Text)) = 0 Then
            If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                txtChallanNo.Text = ""
            End If
            Call RefreshCtrls()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLocationCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtLocationCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Display Help on F1 click
        '*****************************************************************************************
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
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Set focus on Next Control
        '*****************************************************************************************
        On Error GoTo ErrHandler
        If KeyAscii = Asc("'") Then KeyAscii = 0
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
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
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - NA
        'Return Value        - None
        'Function            - Check Validity Of Location Code In The Location_Mst
        '*****************************************************************************************
        On Error GoTo ErrHandler
        txtLocationCode.Text = UCase(txtLocationCode.Text)
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Len(txtLocationCode.Text) > 0 Then
                    If CheckExistanceOfFieldData((txtLocationCode.Text), "Location_Code", "SupplementaryInv_hdr") Then
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
                    If CheckExistanceOfFieldData((txtLocationCode.Text), "Location_Code", "SaleConf") Then
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
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           -  pstrFieldText - Field Text,pstrColumnName - Column Name
        '                    -  pstrTableName - Table Name,pstrCondition - Optional Parameter For Condition
        'Return Value        - None
        'Function            - To Check Validity Of Field Data Whethet it Exists In The Database Or Not
        '*****************************************************************************************
        On Error GoTo ErrHandler
        CheckExistanceOfFieldData = False
        Dim strTableSql As String 'Declared To Make Select Query
        Dim rsExistData As ClsResultSetDB
        If Len(Trim(pstrCondition)) > 0 Then
            strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and Unit_code = '" & gstrUNITID & "' and " & pstrCondition
        Else
            strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and Unit_code = '" & gstrUNITID & "'"
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
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - pstrFieldText - Field Text,pstrColumnName - Column Name
        '                    - pstrTableName - Table Name,pstrCondition - Optional Parameter For Condition
        'Return Value        - None
        'Function            - To Check Validity Of Field Data Whether it Exists In The Database
        '                      Or Not and Return it's Tax Rate
        '*****************************************************************************************
        On Error GoTo ErrHandler
        GetTaxRate = 0
        Dim strTableSql As String 'Declared To Make Select Query
        Dim rsExistData As ClsResultSetDB
        If Len(Trim(pstrCondition)) > 0 Then
            strTableSql = "select " & Trim(pstrFieldName_WhichValueRequire) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and Unit_code = '" & gstrUNITID & "' and " & pstrCondition
        Else
            strTableSql = "select " & Trim(pstrFieldName_WhichValueRequire) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and Unit_code = '" & gstrUNITID & "'"
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
    Private Function SetFinancialYearDates() As Object
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - NA
        'Return Value        - None
        'Function            - To Set Year Start & End Dates in Date From and Date To
        '*****************************************************************************************
        Dim rsBusinessPeriodDates As ClsResultSetDB
        On Error GoTo ErrHandler
        rsBusinessPeriodDates = New ClsResultSetDB
        rsBusinessPeriodDates.GetResult("Select Per_PeriodStart,Per_PeriodEnd from Gen_BusienssPeriod where Unit_code = '" & gstrUNITID & "'")
        If rsBusinessPeriodDates.GetNoRows > 0 Then
            SetFinancialYearDates = True
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function CheckFinancialYearDates() As Boolean
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - NA
        'Return Value        - None
        'Function            - To Check if Business Period is Defined or Not
        '*****************************************************************************************
        Dim rsBusinessPeriodDates As ClsResultSetDB
        On Error GoTo ErrHandler
        CheckFinancialYearDates = False
        rsBusinessPeriodDates = New ClsResultSetDB
        rsBusinessPeriodDates.GetResult("Select Per_PeriodStart,Per_PeriodEnd from Gen_BusienssPeriod where Unit_code = '" & gstrUNITID & "'")
        If rsBusinessPeriodDates.GetNoRows > 0 Then
            Financial_Start_Date = rsBusinessPeriodDates.GetValue("Per_PeriodStart")
            Financial_End_Date = rsBusinessPeriodDates.GetValue("Per_PeriodEnd")
            CheckFinancialYearDates = True
        Else
            MsgBox("No Business Period Defined in Gen_business Period.", MsgBoxStyle.Information, ResolveResString(100))
            CheckFinancialYearDates = False
            Exit Function
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub DisplayNewSOHdrDetails_Old()
        ''*****************************************************************************************
        ''Author              - Nisha Rai
        ''Create Date         - 14/10/2003
        ''Arguments           - None
        ''Return Value        - None
        ''Function            - To Display Detials in Add Mode When We Select Customer Referance
        ''                      and Amendment
        ''*****************************************************************************************
        'Dim rsCustOrdHdr As ClsResultSetDB
        'Dim rsCustOrdDtl As ClsResultSetDB
        'Dim intLoopCounter As Integer
        'Dim intMaxLoopCount As Integer
        'Dim strSQL As String
        'On Error GoTo ErrHandler
        'Set rsCustOrdHdr = New ClsResultSetDB
        'Set rsCustOrdDtl = New ClsResultSetDB
        'strSQL = "Select Currency_code,SalesTax_Type,Surcharge_code from Cust_ord_hdr where active_flag= 'A' and Authorized_flag  = 1 and Account_code = '" & Trim(txtCustCode.Text)
        'strSQL = strSQL & "' and Cust_ref = '" & Trim(txtRefNo.Text) & "' and amendment_no = '" & Trim(txtAmendment.Text) & "'"
        'rsCustOrdHdr.GetResult strSQL
        'If rsCustOrdHdr.GetNoRows > 0 Then
        '    lblCurrency.Caption = rsCustOrdHdr.GetValue("Currency_code")
        '    txtSaleTaxType.Text = rsCustOrdHdr.GetValue("SalesTax_Type")
        '    txtSurchargeTaxType.Text = rsCustOrdHdr.GetValue("Surcharge_code")
        '
        '    strSQL = "select Rate,Cust_Mtrl,Packing,Excise_Duty,Tool_Cost from Cust_ord_dtl where active_flag= 'A' and Authorized_flag  = 1 "
        '    strSQL = strSQL & " and Account_code = '" & Trim(txtCustCode.Text) & "' and Cust_ref = '" & Trim(txtRefNo.Text)
        '    strSQL = strSQL & "' and amendment_no = '" & Trim(txtAmendment.Text) & "' and cust_drgNo = '" & Trim(txtCustPartCode.Text) & "' and "
        '    strSQL = strSQL & " Item_code = '" & Trim(lblItemCode.Caption) & "'"
        '    rsCustOrdDtl.GetResult strSQL
        '    If rsCustOrdDtl.GetNoRows > 0 Then
        '        txtExciseTaxType.Text = rsCustOrdDtl.GetValue("Excise_duty")
        '        intMaxLoopCount = spdPrevInv.maxRows
        '        With spdPrevInv
        '            For intLoopCounter = 1 To intMaxLoopCount
        '                Call .SetText(enumPreInvoiceDetails.New_Rate, intLoopCounter, rsCustOrdDtl.GetValue("Rate"))
        '                Call .SetText(enumPreInvoiceDetails.NewCustSuppMaterial, intLoopCounter, rsCustOrdDtl.GetValue("Cust_mtrl"))
        '                Call .SetText(enumPreInvoiceDetails.NewPacking, intLoopCounter, rsCustOrdDtl.GetValue("Packing"))
        '                Call .SetText(enumPreInvoiceDetails.newToolCost, intLoopCounter, rsCustOrdDtl.GetValue("Tool_cost"))
        '            Next
        '        End With
        '        lblExctax_Per.Caption = GetTaxRate(txtExciseTaxType.Text, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='EXC')")
        '        lblSaltax_Per.Caption = GetTaxRate(txtSaleTaxType.Text, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST')")
        '        lblSurcharge_Per.Caption = GetTaxRate(txtSurchargeTaxType.Text, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SST'")
        '    End If
        'End If
        '    Exit Sub
        'ErrHandler:                             'The Error Handling Code Starts here
        '    Call gobjError.RaiseError(Err.number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        '    Exit Sub
    End Sub
    Private Sub txtRefNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRefNo.TextChanged
        On Error GoTo ErrHandler
        txtRefNo.Text = Replace(txtRefNo.Text, "'", "")
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
                Me.txtAmendment.Focus()
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
    Private Sub txtRemarks_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRemarks.TextChanged
        On Error GoTo ErrHandler
        txtRemarks.Text = Replace(txtRemarks.Text, "'", "")
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtRemarks_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRemarks.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Set focus on Next Control
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
            Case System.Windows.Forms.Keys.Return
                Me.ctlSalesQuantity.Focus()
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
    Private Sub txtSaleTaxType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSaleTaxType.TextChanged
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Refresh S.Tax Type
        '*****************************************************************************************
        On Error GoTo ErrHandler
        txtSaleTaxType.Text = Replace(txtSaleTaxType.Text, "'", "")
        If Len(txtSaleTaxType.Text) = 0 Then
            lblSaltax_Per.Text = "0.00"
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSaleTaxType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSaleTaxType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - At Enter Key Press Set Focus To Next Control
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If Len(txtSaleTaxType.Text) > 0 Then
                            Call txtSaleTaxType_Validating(txtSaleTaxType, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            'txtSurchargeTaxType.SetFocus
                            txtEcssCode.Focus()
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
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Show Help on F1 Key Press
        '*****************************************************************************************
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdSaleTaxType.Enabled Then Call CmdSaleTaxType_Click(CmdSaleTaxType, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSaleTaxType_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSaleTaxType.Leave
        On Error GoTo ErrHandler
        Call CalculateTotalIncoiceValue()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSaleTaxType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtSaleTaxType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Validate STax Type
        '*****************************************************************************************
        On Error GoTo ErrHandler
        If Len(txtSaleTaxType.Text) > 0 Then
            If CheckExistanceOfFieldData((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST' OR Tx_TaxeID='VAT')") Then
                lblSaltax_Per.Text = CStr(GetTaxRate((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST' OR Tx_TaxeID='VAT')"))
                If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                End If
                txtEcssCode.Focus()
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
    Public Function Find_Value(ByRef strField As String) As String
        '----------------------------------------------------------------------------
        'Author         :   Arshad Ali
        'Argument       :   Sql query string as strField
        'Return Value   :   selected table field value as String
        'Function       :   Return a field value from a table
        'Comments       :   Nil
        '----------------------------------------------------------------------------
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
    Private Function ValidatebeforeSave() As Boolean
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To be called on Save Button Click
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Dim lstrControls As String
        Dim lNo As Integer
        Dim blnCheckAddVAT As Boolean
        Dim lctrFocus As System.Windows.Forms.Control
        ValidatebeforeSave = True
        lNo = 1
        lstrControls = ResolveResString(10059)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        blnCheckAddVAT = Find_Value("SELECT ISNULL(CHECKADDITONALVAT,0) AS CHECKADDITONALVAT FROM SALES_PARAMETER where Unit_Code = '" & gstrUNITID & "'")
        If (Len(Me.txtLocationCode.Text) = 0) Then
            lstrControls = lstrControls & vbCrLf & " " & lNo & ". Location Code"
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = Me.txtLocationCode
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Me.txtCustCode.Text) = 0) Then
            lstrControls = lstrControls & vbCrLf & " " & lNo & ". Customer Code"
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = Me.txtCustCode
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Me.txtCustPartCode.Text) = 0) Then
            lstrControls = lstrControls & vbCrLf & " " & lNo & ". Customer Part Code"
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = Me.txtCustCode
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Me.ctlSalesQuantity.Text) = 0) Then
            lstrControls = lstrControls & vbCrLf & " " & lNo & ". Sale Quantity"
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = Me.ctlSalesQuantity
            End If
            ValidatebeforeSave = False
        End If
        If (Len(Me.ctlAccrate.Text) = 0) Then
            lstrControls = lstrControls & vbCrLf & " " & lNo & ". Accessible Rate"
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = Me.ctlAccrate
            End If
            ValidatebeforeSave = False
        End If
        If blnCheckAddVAT = True Then
            If Me.txtAddVAT.Text.Length = 0 Then
                lstrControls = lstrControls & vbCrLf & lNo & ". Additional VAT"
                lNo = lNo + 1
                If lctrFocus Is Nothing Then
                    lctrFocus = Me.txtAddVAT
                End If
                ValidatebeforeSave = False
            End If
        End If
        If Len(Trim(txtExciseTaxType.Text)) = 0 Then
            lblExctax_Per.Text = "0.00"
        End If
        If Len(Trim(txtSaleTaxType.Text)) = 0 Then
            lblSaltax_Per.Text = "0.00"
        End If
        If (Len(lblCurrencyDes.Text) = 0) Then
            lblCurrencyDes.Text = gstrCURRENCYCODE
        End If
        If ValidatebeforeSave = False Then
            MsgBox(lstrControls, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Arrow)
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
    End Function
    Private Sub SaveData()
        '*****************************************************************************************
        'Author              - Nisha Rai
        'Create Date         - 14/10/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Generate Save/Update string
        '*****************************************************************************************
        'Revised By          - Davinder Singh
        'Revision Date       - 20 Apr 2007
        'Revision History    - To include the Secondary Ecess in Supplementary Invoice
        'Issue ID            - 19786
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Dim strsqlHdr As String
        Dim strSqlDtl As String '' Davinder
        Dim strSqlCreditAdvDtl As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim gobjDB As New ClsResultSetDB
        'Variable Decleration for Summery Table
        Dim curRounoffDiff As Decimal
        Dim dblTotalValue As Double
        Dim dblExcise As Double
        Dim dblEcess As Double
        Dim dblSEcess As Double
        Dim dblSaleTax As Double
        Dim salesQuantity As Double
        Dim dblAddVATamount As Double
        Dim dblSurchargeVAT As Double
        'Added for Issue ID eMpro-20090401-29562 Starts
        dblTotalValue = 0
        dblExcise = 0
        dblEcess = 0
        dblSEcess = 0
        dblSaleTax = 0
        dblSurchargeVAT = 0
        'Added for Issue ID eMpro-20090401-29562 Ends
        gobjDB.GetResult("SELECT salestax_roundoff,Excise_Roundoff,ecess_roundoff,ecessroundoff_decimal,salestax_roundoff_decimal,excise_roundoff_decimal,totalinvoiceamount_roundoff,totalinvoiceamountroundoff_decimal, SST_Roundoff,SST_Roundoff_decimal FROM sales_parameter where Unit_Code = '" & gstrUNITID & "'")
        mblnExciseRoundOFFFlag = gobjDB.GetValue("Excise_Roundoff")
        mblnEcessRoundOFFFlag = gobjDB.GetValue("ecess_roundoff")
        mblnTotalInvoiceRoundOFFFlag = gobjDB.GetValue("totalinvoiceamount_roundoff")
        mblnsalestaxRoundOFFFlag = gobjDB.GetValue("salestax_roundoff")
        intExciseRoundOFFplace = gobjDB.GetValue("excise_roundoff_decimal")
        intEcessRoundOFFplace = gobjDB.GetValue("ecessroundoff_decimal")
        intTotalInvoiceRoundOFFplace = gobjDB.GetValue("totalinvoiceamountroundoff_decimal")
        intsalestaxRoundOFFplace = gobjDB.GetValue("salestax_roundoff_decimal")
        mblnISSurChargeTaxRoundOff = gobjDB.GetValue("SST_Roundoff")
        stSSTRoundOffDecimal = gobjDB.GetValue("SST_Roundoff_decimal")
        If mblnExciseRoundOFFFlag Then
            dblExcise = System.Math.Round((Val(ctlSalesQuantity.Text) * Val(ctlAccrate.Text)) * CDbl(lblExctax_Per.Text) / 100)
        Else
            dblExcise = System.Math.Round((Val(ctlSalesQuantity.Text) * Val(ctlAccrate.Text)) * CDbl(lblExctax_Per.Text) / 100, intEcessRoundOFFplace)
        End If
        If mblnEcessRoundOFFFlag Then
            dblEcess = System.Math.Round(Val(CStr(dblExcise * CDbl(lblEcssCode.Text) / 100)))
        Else
            dblEcess = System.Math.Round(Val(CStr(dblExcise * CDbl(lblEcssCode.Text) / 100)), intEcessRoundOFFplace)
        End If
        If mblnEcessRoundOFFFlag Then
            dblSEcess = System.Math.Round(Val(CStr(dblExcise * CDbl(lblSEcssCode.Text) / 100)))
        Else
            dblSEcess = System.Math.Round(Val(CStr(dblExcise * CDbl(lblSEcssCode.Text) / 100)), intEcessRoundOFFplace)
        End If
        If mblnsalestaxRoundOFFFlag Then
            dblSaleTax = System.Math.Round(((Val(ctlSalesQuantity.Text) * Val(ctlBasicRate.Text)) + dblExcise + dblEcess + dblSEcess) * CDbl(lblSaltax_Per.Text) / 100)
            dblAddVATamount = System.Math.Round(((Val(ctlSalesQuantity.Text) * Val(ctlBasicRate.Text)) + dblExcise + dblEcess + dblSEcess) * CDbl(lblAddVAT.Text) / 100)
        Else
            dblSaleTax = System.Math.Round(((Val(ctlSalesQuantity.Text) * Val(ctlBasicRate.Text)) + dblExcise + dblEcess + dblSEcess) * CDbl(lblSaltax_Per.Text) / 100, intsalestaxRoundOFFplace)
            dblAddVATamount = System.Math.Round(((Val(ctlSalesQuantity.Text) * Val(ctlBasicRate.Text)) + dblExcise + dblEcess + dblSEcess) * CDbl(lblAddVAT.Text) / 100, intsalestaxRoundOFFplace)
        End If
        If mblnISSurChargeTaxRoundOff Then
            dblSurchargeVAT = System.Math.Round((dblSaleTax * Convert.ToDouble(lblSurcharge_VAT.Text) / 100))
        Else
            dblSurchargeVAT = System.Math.Round((dblSaleTax * Convert.ToDouble(lblSurcharge_VAT.Text) / 100), stSSTRoundOffDecimal)
        End If
        If mblnTotalInvoiceRoundOFFFlag Then
            dblTotalValue = (Val(ctlSalesQuantity.Text) * Val(ctlBasicRate.Text)) + dblEcess + dblSEcess + dblExcise + dblSaleTax + Val(ctlOthers.Text) + dblAddVATamount + dblSurchargeVAT
            curRounoffDiff = dblTotalValue - System.Math.Round(dblTotalValue)
            dblTotalValue = System.Math.Round(dblTotalValue)
        Else
            dblTotalValue = (Val(ctlSalesQuantity.Text) * Val(ctlBasicRate.Text)) + dblEcess + dblSEcess + dblExcise + dblSaleTax + Val(ctlOthers.Text) + dblAddVATamount + dblSurchargeVAT
            curRounoffDiff = dblTotalValue - System.Math.Round(dblTotalValue, intTotalInvoiceRoundOFFplace)
            dblTotalValue = System.Math.Round((Val(ctlSalesQuantity.Text) * Val(ctlBasicRate.Text)) + dblEcess + dblSEcess + dblExcise + dblSaleTax + Val(ctlOthers.Text) + dblAddVATamount + dblSurchargeVAT, intTotalInvoiceRoundOFFplace)
        End If
        ctlTotals.Text = CStr(dblTotalValue)
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        Select Case CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                strsqlHdr = "Insert into SupplementaryInv_hdr ("
                strsqlHdr = strsqlHdr & "Location_Code,Account_Code,Cust_name,Cust_Ref,Amendment_No,Doc_No,Invoice_DateFrom,Invoice_DateTo,Invoice_Date,Bill_Flag,Cancel_flag,Item_Code,"
                strsqlHdr = strsqlHdr & "Cust_Item_Code,Currency_Code,Rate,Basic_Amount,Accessible_amount,Excise_type,"
                strsqlHdr = strsqlHdr & "Excise_per,TotalExciseAmount,"
                strsqlHdr = strsqlHdr & "SalesTax_Type,SalesTax_Per,Sales_Tax_Amount,"
                strsqlHdr = strsqlHdr & "total_amount,"
                strsqlHdr = strsqlHdr & "SuppInv_Remarks,remarks,Ent_dt,Ent_UserId,Upd_dt,Upd_Userid"
                strsqlHdr = strsqlHdr & ",ECESS_Type,ECESS_Per,ECESS_Amount,supp_invdetail,sales_Quantity"
                strsqlHdr = strsqlHdr & ",SECESS_Type,SECESS_Per,SECESS_Amount,totalInvoiceAmtRoundoff_diff,ADDVAT_TYPE,ADDVAT_PER,ADDVAT_AMOUNT, SURCHARGE_SALESTAXTYPE, SURCHARGE_SALESTAX_PER, SURCHARGE_SALES_TAX_AMOUNT, Unit_Code"
                strsqlHdr = strsqlHdr & ") Values"
                strsqlHdr = strsqlHdr & " ('" & Trim(txtLocationCode.Text) & "','" & Trim(txtCustCode.Text) & "','" & Trim(lblCustCodeDes.Text) & "','"
                strsqlHdr = strsqlHdr & Trim(txtRefNo.Text) & "','" & Trim(txtAmendment.Text) & "'," & txtChallanNo.Text & ",'" & getDateForDB(GetServerDate()) & "','" & getDateForDB(GetServerDate()) & "','" & getDateForDB(GetServerDate()) & "',0,0,'" & Trim(lblItemCode.Text)
                strsqlHdr = strsqlHdr & "','" & Trim(txtCustPartCode.Text) & "','" & Trim(lblCurrencyDes.Text) & "'," & (ctlAccrate.Text) & ","
                strsqlHdr = strsqlHdr & (Val(ctlSalesQuantity.Text) * Val(ctlBasicRate.Text)) & "," & (Val(ctlSalesQuantity.Text) * Val(ctlAccrate.Text)) & ",'" & Trim(txtExciseTaxType.Text) & "',"
                strsqlHdr = strsqlHdr & Trim(lblExctax_Per.Text) & ","
                strsqlHdr = strsqlHdr & dblExcise & ","
                strsqlHdr = strsqlHdr & "'" & Trim(txtSaleTaxType.Text) & "'," & lblSaltax_Per.Text & "," & dblSaleTax
                strsqlHdr = strsqlHdr & "," & dblTotalValue & ",'" & Trim(txtCustRefRemarks.Text) & "','" & Trim(txtRemarks.Text) & "','" & getDateForDB(GetServerDate()) & "','"
                strsqlHdr = strsqlHdr & mP_User & "','" & getDateForDB(GetServerDate()) & "','" & mP_User & "',"
                strsqlHdr = strsqlHdr & "'" & Trim(Me.txtEcssCode.Text) & "'," & Trim(Me.lblEcssCode.Text) & "," & Val(CStr(dblEcess)) & ",'" & "O" & "'," & ctlSalesQuantity.Text & ",'"
                strsqlHdr = strsqlHdr & Trim(txtSEcssCode.Text) & "'," & Trim(lblSEcssCode.Text) & "," & Val(CStr(dblSEcess)) & "," & curRounoffDiff
                strsqlHdr = strsqlHdr & ",'" & txtAddVAT.Text.Trim & "'," & Val(lblAddVAT.Text) & "," & dblAddVATamount & ",'"
                strsqlHdr = strsqlHdr & txtSurcharge_VAT.Text.Trim & "'," & Val(lblSurcharge_VAT.Text) & "," & dblSurchargeVAT & ", '" & gstrUNITID & "')"
                strSqlDtl = "insert into SupplementaryInv_Dtl ("
                strSqlDtl = strSqlDtl & "Location_Code,Doc_No,SuppInvDate,Item_code,Cust_Item_Code,"
                strSqlDtl = strSqlDtl & "Rate_diff,Quantity,Basic_AmountDiff,Accessible_amountDiff,"
                strSqlDtl = strSqlDtl & "TotalExciseAmountDiff,Sales_Tax_AmountDiff,"
                strSqlDtl = strSqlDtl & "total_amountDiff,Ent_dt,Ent_UserID,Upd_Dt,Upd_UserID,ECESS_Amount_Diff,Refdoc_No,Unit_Code)"
                strSqlDtl = strSqlDtl & "Values ('" & Trim(txtLocationCode.Text) & "'," & Trim(txtChallanNo.Text) & ",'" & getDateForDB(GetServerDate()) & "','" & lblItemCode.Text & "','"
                strSqlDtl = strSqlDtl & Trim(txtCustPartCode.Text) & "'," & ctlAccrate.Text & "," & ctlSalesQuantity.Text & ","
                strSqlDtl = strSqlDtl & (Val(ctlSalesQuantity.Text) * Val(ctlBasicRate.Text)) & ","
                strSqlDtl = strSqlDtl & (Val(ctlSalesQuantity.Text) * Val(ctlAccrate.Text)) & "," & dblExcise & "," & dblSaleTax & "," & dblTotalValue & ",'"
                strSqlDtl = strSqlDtl & getDateForDB(GetServerDate()) & "','" & mP_User & "','" & getDateForDB(GetServerDate()) & "','" & mP_User & "'," & Val(CStr(dblEcess)) & ",0, '" & gstrUNITID & "')"
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                '        strsqlHdr = "update SupplementaryInv_hdr set Cust_ref = '" & Trim(txtRefNo.Text) & "',Amendment_No = '" & Trim(txtAmendment.Text) & "',"
                '        strsqlHdr = strsqlHdr & "Currency_Code ='" & Trim(lblCurrencyDes.Caption) & "' ,Rate = " & dblSummeryRate & ",Basic_Amount = " & dblSummeryBasic_Amount
                '        strsqlHdr = strsqlHdr & ",Accessible_amount = " & dblSummeryAccessible_amount & ",Excise_type = '" & Trim(txtExciseTaxType) & "',"
                '        strsqlHdr = strsqlHdr & " CVD_type = '" & Trim(txtCVDCode.Text) & "',SAD_type = '" & Trim(txtSADCode.Text) & "', Excise_per = " & val(lblExctax_Per.Caption)
                '        strsqlHdr = strsqlHdr & ",CVD_per = " & val(lblCVD_Per.Caption) & ",SVD_per = " & val(lblSAD_per.Caption)
                '        strsqlHdr = strsqlHdr & ",TotalExciseAmount = " & dblSummeryTotalExciseAmount & ",CustMtrl_Amount = " & dblSummeryCustMtrl_Amount & ","
                '        strsqlHdr = strsqlHdr & "ToolCost_amount = " & dblSummeryToolCost_amount & ",SalesTax_Type = '" & Trim(txtSaleTaxType.Text) & "',"
                '        strsqlHdr = strsqlHdr & "SalesTax_Per =" & val(lblSaltax_Per.Caption) & ",Sales_Tax_Amount = " & dblSummerySales_Tax_Amount
                '        strsqlHdr = strsqlHdr & ",Surcharge_salesTaxType = '" & Trim(txtSurchargeTaxType.Text) & "',Surcharge_SalesTax_Per = " & val(lblSurcharge_Per)
                '        strsqlHdr = strsqlHdr & ",Surcharge_Sales_Tax_Amount = " & dblSummerySurcharge_Sales_Tax_Amount & ",total_amount = " & dblSummerytotal_amount
                '        strsqlHdr = strsqlHdr & ", SuppInv_Remarks = '" & Trim(txtCustRefRemarks.Text) & "', remarks = '" & Trim(txtRemarks.Text) & "'"
                '        'Code add by Sourabh
                '        strsqlHdr = strsqlHdr & ",ECESS_Type = '" & Trim(Me.txtEcssCode.Text) & "',ECESS_Per = " & Trim(Me.lblEcssCode.Caption) & ",ECESS_Amount = " & dblSummeryEcss_amount
                '        '*******************
                '        strsqlHdr = strsqlHdr & ", Upd_dt = '" & GetServerDate & "',Upd_Userid = '" & mP_User & "' where Location_code = '"
                '        strsqlHdr = strsqlHdr & Trim(txtLocationCode.Text) & "' and Doc_no = " & Trim(txtChallanNo.Text)
                '
                '        '************String For Deletion
                '        strDeleteSuppDtl = "Delete from SupplementaryInv_Dtl Where Doc_no = '" & Trim(txtChallanNo.Text) & "'"
                '        strDeleteSuppCrAdvise = "Delete from SuppCreditAdvise_Dtl Where Doc_no = '" & Trim(txtChallanNo.Text) & "'"
        End Select
        If Len(Trim(strsqlHdr)) > 0 Then
            mP_Connection.BeginTrans()
            mP_Connection.Execute("set dateFormat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            mP_Connection.Execute(strsqlHdr, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            mP_Connection.Execute(strSqlDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            mP_Connection.CommitTrans()
        End If
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        MsgBox("Transaction Completed successfully!", MsgBoxStyle.Information, ResolveResString(100))
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SelectChallanNoFromSupplementatryInvHdr()
        '*****************************************************************************************
        'Function            - To Generate temperory No
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Dim strChallanNo As String
        Dim rsChallanNo As New ClsResultSetDB
        strChallanNo = "SELECT (CURRENT_NO + 1)CURRENT_NO FROM DOCUMENTTYPE_MST WHERE DOC_TYPE = 9999 AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE and Unit_Code ='" & gstrUNITID & "'"
        rsChallanNo.GetResult(strChallanNo, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsChallanNo.GetNoRows > 0 Then
            strChallanNo = rsChallanNo.GetValue("CURRENT_NO").ToString
            While Len(strChallanNo) < 6
                strChallanNo = "0" + strChallanNo
            End While
            strChallanNo = "99" + strChallanNo
            txtChallanNo.Text = strChallanNo
            strChallanNo = "UPDATE DOCUMENTTYPE_MST SET CURRENT_NO = CURRENT_NO + 1 WHERE DOC_TYPE = 9999 AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE and Unit_Code ='" & gstrUNITID & "'"
            mP_Connection.Execute(strChallanNo, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Else
            MsgBox("Temporary Invoice No. Series Not Define. Invoice Entry Can Not Be Saved.", MsgBoxStyle.Information, ResolveResString(100))
            txtChallanNo.Text = ""
        End If
        rsChallanNo.ResultSetClose()
        rsChallanNo = Nothing
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        rsChallanNo = Nothing
    End Sub
    Private Sub CalculateTotalIncoiceValue()
        '*****************************************************************************************
        'Revised By          - Davinder Singh
        'Revision Date       - 20 Apr 2007
        'Revision History    - To include the Secondary Ecess in Supplementary Invoice
        'Issue ID            - 19786
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Dim gobjDB As New ClsResultSetDB
        Dim dblTotalValue As Double
        Dim dblExcise As Double
        Dim dblEcess As Double
        Dim dblSEcess As Double
        Dim dblSaleTax As Double
        Dim dblAddVAT As Double
        Dim dblSurchargeVAT As Double
        gobjDB.GetResult("SELECT salestax_roundoff,Excise_Roundoff,ecess_roundoff,ecessroundoff_decimal,salestax_roundoff_decimal,excise_roundoff_decimal,totalinvoiceamount_roundoff,totalinvoiceamountroundoff_decimal, SST_Roundoff,SST_Roundoff_decimal FROM sales_parameter where Unit_Code ='" & gstrUNITID & "'")
        mblnExciseRoundOFFFlag = gobjDB.GetValue("Excise_Roundoff")
        mblnEcessRoundOFFFlag = gobjDB.GetValue("ecess_roundoff")
        mblnTotalInvoiceRoundOFFFlag = gobjDB.GetValue("totalinvoiceamount_roundoff")
        mblnsalestaxRoundOFFFlag = gobjDB.GetValue("salestax_roundoff")
        intExciseRoundOFFplace = gobjDB.GetValue("excise_roundoff_decimal")
        intEcessRoundOFFplace = gobjDB.GetValue("ecessroundoff_decimal")
        intTotalInvoiceRoundOFFplace = gobjDB.GetValue("totalinvoiceamountroundoff_decimal")
        intsalestaxRoundOFFplace = gobjDB.GetValue("salestax_roundoff_decimal")
        mblnISSurChargeTaxRoundOff = gobjDB.GetValue("SST_Roundoff")
        stSSTRoundOffDecimal = gobjDB.GetValue("SST_Roundoff_decimal")
        If mblnExciseRoundOFFFlag Then
            dblExcise = System.Math.Round((Val(ctlSalesQuantity.Text) * Val(ctlAccrate.Text)) * CDbl(lblExctax_Per.Text) / 100)
        Else
            dblExcise = System.Math.Round((Val(ctlSalesQuantity.Text) * Val(ctlAccrate.Text)) * CDbl(lblExctax_Per.Text) / 100, intEcessRoundOFFplace)
        End If
        If mblnEcessRoundOFFFlag Then
            dblEcess = System.Math.Round(Val(CStr(dblExcise * CDbl(lblEcssCode.Text) / 100)))
        Else
            dblEcess = System.Math.Round(Val(CStr(dblExcise * CDbl(lblEcssCode.Text) / 100)), intEcessRoundOFFplace)
        End If
        If mblnEcessRoundOFFFlag Then
            dblSEcess = System.Math.Round(Val(CStr(dblExcise * CDbl(lblSEcssCode.Text) / 100)))
        Else
            dblSEcess = System.Math.Round(Val(CStr(dblExcise * CDbl(lblSEcssCode.Text) / 100)), intEcessRoundOFFplace)
        End If
        If mblnsalestaxRoundOFFFlag Then
            dblSaleTax = System.Math.Round(((Val(ctlSalesQuantity.Text) * Val(ctlAccrate.Text)) + dblExcise + dblEcess + dblSEcess) * CDbl(lblSaltax_Per.Text) / 100)
            dblAddVAT = System.Math.Round(((Val(ctlSalesQuantity.Text) * Val(ctlAccrate.Text)) + dblExcise + dblEcess + dblSEcess) * CDbl(lblAddVAT.Text) / 100)
        Else
            dblSaleTax = System.Math.Round(((Val(ctlSalesQuantity.Text) * Val(ctlAccrate.Text)) + dblExcise + dblEcess + dblSEcess) * CDbl(lblSaltax_Per.Text) / 100, intsalestaxRoundOFFplace)
            dblAddVAT = System.Math.Round(((Val(ctlSalesQuantity.Text) * Val(ctlAccrate.Text)) + dblExcise + dblEcess + dblSEcess) * CDbl(lblAddVAT.Text) / 100, intsalestaxRoundOFFplace)
        End If
        If mblnISSurChargeTaxRoundOff Then
            dblSurchargeVAT = System.Math.Round((dblSaleTax * Convert.ToDouble(lblSurcharge_VAT.Text) / 100))
        Else
            dblSurchargeVAT = System.Math.Round((dblSaleTax * Convert.ToDouble(lblSurcharge_VAT.Text) / 100), stSSTRoundOffDecimal)
        End If
        If mblnTotalInvoiceRoundOFFFlag Then
            dblTotalValue = System.Math.Round((Val(ctlSalesQuantity.Text) * Val(ctlAccrate.Text)) + dblEcess + dblSEcess + dblExcise + dblSaleTax + Val(ctlOthers.Text) + dblAddVAT + dblSurchargeVAT)
        Else
            dblTotalValue = System.Math.Round((Val(ctlSalesQuantity.Text) * Val(ctlAccrate.Text)) + dblEcess + dblSEcess + dblExcise + dblSaleTax + Val(ctlOthers.Text) + dblAddVAT + dblSurchargeVAT, intTotalInvoiceRoundOFFplace)
        End If
        ctlTotals.Text = CStr(dblTotalValue)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function Checknumeric(ByVal KeyAscii As Short) As Short
        On Error GoTo ErrHandler
        If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> System.Windows.Forms.Keys.Delete And KeyAscii <> System.Windows.Forms.Keys.Back And KeyAscii <> System.Windows.Forms.Keys.Return) Then
            Checknumeric = 0
        Else
            Checknumeric = KeyAscii
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub RefreshCtrls()
        On Error GoTo ErrHandler
        txtCustCode.Text = ""
        txtCustPartCode.Text = ""
        lblCustItemDesc.Text = ""
        lblItemCode.Text = ""
        txtRefNo.Text = ""
        txtAmendment.Text = ""
        lblAddressDes.Text = ""
        txtCustRefRemarks.Text = ""
        txtRemarks.Text = ""
        ctlSalesQuantity.Text = "0.00"
        ctlBasicRate.Text = "0.00"
        ctlAccrate.Text = "0.00"
        txtExciseTaxType.Text = ""
        lblExctax_Per.Text = "0.00"
        txtEcssCode.Text = ""
        lblEcssCode.Text = "0.00"
        txtSEcssCode.Text = ""
        lblSEcssCode.Text = "0.00"
        txtSaleTaxType.Text = ""
        lblSaltax_Per.Text = "0.00"
        ctlTotals.Text = "0.00"
        lblCurrencyDes.Text = ""
        txtAddVAT.Text = ""
        lblAddVAT.Text = "0.00"
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSEcssCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSEcssCode.TextChanged
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Function            - To Refresh SEcss Label
        '*****************************************************************************************
        On Error GoTo ErrHandler
        txtSEcssCode.Text = Replace(txtSEcssCode.Text, "'", "")
        If Trim(txtSEcssCode.Text) = "" Then
            lblSEcssCode.Text = "0.00"
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSEcssCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSEcssCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Function            - To Call Help on F1 Click
        '*****************************************************************************************
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
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Function            - To call validation on pressing enter
        '*****************************************************************************************
        On Error GoTo ErrHandler
        If KeyAscii = Asc("'") Then KeyAscii = 0
        If KeyAscii = 13 Then txtSEcssCode_Validating(txtSEcssCode, New System.ComponentModel.CancelEventArgs((False)))
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtSEcssCode_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSEcssCode.Leave
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Function            - To Calculate total invoice value
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Call CalculateTotalIncoiceValue()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSEcssCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtSEcssCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*****************************************************************************************
        'Author              - Davinder Singh
        'Function            - To validate the data
        '*****************************************************************************************
        On Error GoTo ErrHandler
        If Trim(txtSEcssCode.Text) <> "" Then
            If CheckExistanceOfFieldData((txtSEcssCode.Text), "TxRt_Rate_No", "Gen_TaxRate", " Tx_TaxeID='ECSSH'") Then
                lblSEcssCode.Text = CStr(GetTaxRate((txtSEcssCode.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='ECSSH'"))
                CmdGrpChEnt.Focus()
            Else
                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtSEcssCode.Text = ""
                If txtSEcssCode.Enabled Then txtSEcssCode.Focus()
            End If
        Else
            CmdGrpChEnt.Focus()
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub ctlAccrate_Change(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlAccrate.Change
        On Error GoTo ErrHandler
        ctlAccrate.Text = Replace(Trim(ctlAccrate.Text), "'", "")
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlAccrate_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles ctlAccrate.KeyPress
        On Error GoTo ErrHandler
        e.KeyAscii = Checknumeric(e.KeyAscii)
        Select Case e.KeyAscii
            Case 39, 34, 96
                e.KeyAscii = 0
            Case System.Windows.Forms.Keys.Return
                Me.txtExciseTaxType.Focus()
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlAccrate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ctlAccrate.LostFocus
        On Error GoTo ErrHandler
        Call CalculateTotalIncoiceValue()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlBasicRate_Change(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlBasicRate.Change
        On Error GoTo ErrHandler
        ctlBasicRate.Text = Replace(Trim(ctlBasicRate.Text), "'", "")
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlBasicRate_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles ctlBasicRate.KeyPress
        On Error GoTo ErrHandler
        e.KeyAscii = Checknumeric(e.KeyAscii)
        Select Case e.KeyAscii
            Case 39, 34, 96
                e.KeyAscii = 0
            Case System.Windows.Forms.Keys.Return
                'txtCVDCode.SetFocus
                Me.ctlAccrate.Focus()
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlBasicRate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ctlBasicRate.LostFocus
        On Error GoTo ErrHandler
        ctlAccrate.Text = ctlBasicRate.Text
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlSalesQuantity_Change(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlSalesQuantity.Change
        On Error GoTo ErrHandler
        ctlSalesQuantity.Text = Replace(Trim(ctlSalesQuantity.Text), "'", "")
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlSalesQuantity_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles ctlSalesQuantity.KeyPress
        On Error GoTo ErrHandler
        e.KeyAscii = Checknumeric(e.KeyAscii)
        Select Case e.KeyAscii
            Case 39, 34, 96
                e.KeyAscii = 0
            Case System.Windows.Forms.Keys.Return
                'txtCVDCode.SetFocus
                ctlBasicRate.Focus()
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlSalesQuantity_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ctlSalesQuantity.LostFocus
        On Error GoTo ErrHandler
        Call CalculateTotalIncoiceValue()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlOthers_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles ctlOthers.KeyPress
        On Error GoTo ErrHandler
        Call CalculateTotalIncoiceValue()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlTotals_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ctlTotals.TextChanged
        On Error GoTo ErrHandler
        ctlTotals.Text = Replace(Trim(ctlTotals.Text), "'", "")
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdAddVAT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddVAT.Click
        '****************************************************
        'Created By     -  Manoj Vaish
        'Description    -  To Display Additonal VAT Help 
        '****************************************************
        Dim strHelp As String
        Dim strSTaxHelp() As String
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strHelp = "Select TxRT_Rate_No,TxRt_Percentage from Gen_taxRate where Tx_TaxeID in('ADVAT','ADCST') and Unit_Code ='" & gstrUNITID & "'"
                strSTaxHelp = Me.ctlEMPHelpSuppInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, "Add. VAT/CST Tax Help")
                If UBound(strSTaxHelp) <= 0 Then Exit Sub
                If strSTaxHelp(0) = "0" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtAddVAT.Text = "" : txtAddVAT.Focus() : Exit Sub
                Else
                    txtAddVAT.Text = strSTaxHelp(0)
                    lblAddVAT.Text = strSTaxHelp(1)
                    CalculateTotalIncoiceValue()
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtAddVAT_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAddVAT.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If Len(txtAddVAT.Text) > 0 Then
                            Call txtAddVAT_Validating(txtAddVAT, New System.ComponentModel.CancelEventArgs(False))
                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtAddVAT_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAddVAT.KeyUp
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If cmdAddVAT.Enabled Then Call cmdAddVAT_Click(cmdAddVAT, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtAddVAT_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAddVAT.Leave
        On Error GoTo ErrHandler
        Call CalculateTotalIncoiceValue()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtAddVAT_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAddVAT.TextChanged
        On Error GoTo ErrHandler
        If Len(txtAddVAT.Text) = 0 Then
            lblAddVAT.Text = "0.00"
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtAddVAT_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtAddVAT.Validating
        Dim Cancel As Boolean = e.Cancel
        On Error GoTo ErrHandler
        If Len(txtAddVAT.Text) > 0 Then
            If CheckExistanceOfFieldData((txtAddVAT.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='ADVAT' OR Tx_TaxeID='ADCST')") Then
                lblAddVAT.Text = CStr(GetTaxRate((txtAddVAT.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='ADVAT' OR Tx_TaxeID='ADCST')"))
            Else
                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtAddVAT.Text = ""
                If txtAddVAT.Enabled Then txtAddVAT.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        e.Cancel = Cancel
    End Sub
    Private Sub txtSurcharge_VAT_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSurcharge_VAT.KeyDown
        On Error GoTo ErrHandler
        If e.KeyCode = System.Windows.Forms.Keys.F1 Then
            Call cmdSurcharge_VAT_Click(cmdSurcharge_VAT, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSurcharge_VAT_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSurcharge_VAT.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtSurcharge_VAT_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSurcharge_VAT.Leave
        On Error GoTo ErrHandler
        Call CalculateTotalIncoiceValue()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSurcharge_VAT_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSurcharge_VAT.TextChanged
        If Trim(txtSurcharge_VAT.Text) = "" Then
            lblSurcharge_VAT.Text = "0.00"
        End If
    End Sub
    Private Sub txtSurcharge_VAT_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtSurcharge_VAT.Validating
        On Error GoTo ErrHandler
        If Trim(txtSurcharge_VAT.Text) <> "" Then
            If CheckExistanceOfFieldData((txtSurcharge_VAT.Text), "TxRt_Rate_No", "Gen_TaxRate", " Tx_TaxeID='SST'") Then
                lblSurcharge_VAT.Text = CStr(GetTaxRate((txtSurcharge_VAT.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SST'"))
            Else
                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                e.Cancel = True
                txtSurcharge_VAT.Text = ""
                If txtSurcharge_VAT.Enabled Then txtSurcharge_VAT.Focus()
            End If
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdSurcharge_VAT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSurcharge_VAT.Click
        On Error GoTo ErrHandler
        Dim strHelp As String
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Len(Me.txtSurcharge_VAT.Text) = 0 Then 'To check if There is No Text Then Show All Help
                    strHelp = ShowList(1, (txtSurcharge_VAT.MaxLength), "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='SST'")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtSurcharge_VAT.Text = strHelp
                        txtSurcharge_VAT.Focus()
                    End If
                Else
                    strHelp = ShowList(1, (txtSurcharge_VAT.MaxLength), txtSurcharge_VAT.Text, "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='SST'")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtSurcharge_VAT.Text = strHelp
                    End If
                End If
                Call txtSurcharge_VAT_Validating(txtSurcharge_VAT, New System.ComponentModel.CancelEventArgs(False))
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtCustPartCode_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCustPartCode.Validating
        'ISSUE ID : 10158545 New validation
        On Error GoTo ErrHandler
        If Trim(txtCustPartCode.Text) <> "" Then
            If Not CheckExistanceOfFieldData((txtCustPartCode.Text), "Cust_drgno", "Custitem_mst", " active=1 and account_code='" & txtCustCode.Text & "'") Then
                MsgBox("No Customer part Code Exists", MsgBoxStyle.Information, ResolveResString(100)) : txtCustPartCode.Text = "" : txtCustPartCode.Focus() : Exit Sub
                'Cancel = True
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
        'ISSUE ID End : 10158545 
    End Sub

  
  
End Class