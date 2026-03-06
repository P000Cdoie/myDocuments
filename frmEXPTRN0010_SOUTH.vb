Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.IO
Friend Class frmEXPTRN0010_SOUTH
    Inherits System.Windows.Forms.Form
    '===================================================================================
    ' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
    ' File Name         :   FRMEXPTRN0010.frm
    ' Function          :   Used to add sale details
    ' Created By        :   Nisha & Kapil
    ' Created On        :   15 May, 2001
    ' Revision History  :   Nisha Rai
    '21/09/2001 MARKED CHECKED BY BCs changed on version 3
    '03/10/2001 MARKED CHECKED BY BCs  for jobwork invoice changed on version 7
    '09/10/2001  changed on version 8 for schedule Status
    '09/01/2002 changed fof Smiel Chennei to add CVD_PER,SVD_Per,Insurance
    '25/01/2002 changed for decimal 4 places on Chacked Out Form No = 4019
    '28/01/2002 changed for decimal 4 places on Chacked Out Form No = 4033
    'in ChangeCellTypeStaticText()
    '02/02/2002 Add Export Challan Entry
    '15/01/2002 CHANGED FOR DOCUMENT NO. ON FORM NO. 4068
    '22/03/2002 INCREASED SIZE OF CONTAINER NO.
    '27/06/2002 DatePicker Added ,So that Dates of Export Invoice can be set.   -
    '                    - NITIN SOOD
    'changed by nisha on 21/03/2003 for financial rollover & temp no will start from 99000001
    'changed by nisha on 01/05/2003 for back date entry check
    'Changed by nisha 0n 25/07/203
    'Three new feilds added 1.Service type invoice Check Box
    '                        2.Bank id
    '                        3.Remarks
    '===================================================================================
    'Revised By         : Arul on
    'Revised on         : 26-08-2005
    'Revised For        : To captupe the Unt_CodeID,Doc_No,Advance_lice_No,Pallet_Length,Pallet_width,Pallet_Height,Pallet_Total,ARE_NO,Net_Weight,Gross_Weight and Export_Type
    '===================================================================================
    'Revised By         : Davinder Singh
    'Revised on         : 12-Feb-2007
    'Revision History   : 1) To Make INvoice Agst Despatch Advice
    '                     2) Two new Functions CheckMKTSchedules and UpdateMKTSchedules
    '                        are Intoduced in which two Stored Procedures are used for
    '                        Checking and Updation of Schedules
    '                     3) New function FillDataAgstDespatchAdvise is addedin which
    '                        stored procedure INVOICE_AGST_DISPATCHADVICE  is used to fetch data from database
    '===================================================================================
    'Revised By         : Manoj Kr. Vaish
    'Revised on         : 07-Sep-2007 Issue ID:21054
    'Revised For        : To validate the invoice extry only for dispatch advice.
    '===================================================================================
    'Revised By         : Manoj Kr. Vaish
    'Revised on         : 23-Sep-2008 Issue ID:eMpro-20080923-21892
    'Revised For        : Changes has been reverted for Export Invoice entry through sales order
    '===================================================================================
    'Revised By         : Manoj Vaish
    'Revision Date      : 02 Mar 2009
    'Issue ID           : eMpro-20090227-27987
    'Revision History   : Changes for commercial invoice at Mate Units
    '-----------------------------------------------------------------------------
    'Revised By         : Manoj Kr. Vaish
    'Issue ID           : eMpro-20090611-32362
    'Revision Date      : 16 Jun 2009
    'History            : Export Invoice for Nissan---HILEX
    '                     To add the Bin Quantity field and populate From Box
    '                     and To Box according to Bin Quantity
    '-----------------------------------------------------------------------------
    'Revised By        -    Vinod Singh
    'Revision Date     -    19/05/2011
    'Revision History  -    Changes for Multi Unit
    '-----------------------------------------------------------------------------
    'Revised By         : Prashant Rajpal
    'Issue ID           : 10107074
    'Revision Date      : 22 Jun 2011
    'History            : Changes for ASN related Customer - More than one item Invoice not picked, now picked
    '****************************************************************************************
    'Modified by Jyolsna VN on 12-Oct-2011 for MultiUnit Change Management
    '----------------------------------------------------------------------------------------
    '****************************************************************************************
    'Revised By         : Prashant Rajpal
    'Issue ID           : 10127477
    'Revision Date      : 18 Aug 2011
    'History            : Changes for ASN related Customer - All Mode of Transport appeared now
    '****************************************************************************************
    'Modified by Roshan Singh on 09-Nov-2011 for MultiUnit Change Management
    '****************************************************************************************
    'Revised By         : Prashant Rajpal
    'Issue ID           : 10227422 
    'Revision Date      : 22 May 2012
    'History            : In Export invoice , Total Value is calculated wrong for 6 decimal Rate for MTL sharjah.
    '****************************************************************************************
    'Revised By         :   Prashant Rajpal
    'Revision Date      :   27-Aug 2012
    'Issue ID           :   10266201  
    'Revision History   :   Incoroprate the validation of Sales Order (Despatch quantity shoul  d not be greater than Schedule Qty )
    '**********************************************************************************************************************
    'Revised By         :   Prashant Rajpal
    'Revision Date      :   27-Aug 2012
    'Issue ID           :   10274457   
    'Revision History   :  For mtl sharjah , no need for validation of sales order schedule vs disaptch
    '**********************************************************************************************************************
    'Revised By         :   Prashant Rajpal
    'Revision Date      :   18-Mar-2013
    'Issue ID           :   10354980   
    'Revision History   :   Woco migration changes
    '**********************************************************************************************************************
    'Revised By         :   Prashant Rajpal
    'Revision Date      :   11-Apr-2014
    'Issue ID           :   10549878 
    'Revision History   :   Mode of Transport not appearing sometimes in export invoice 
    '**********************************************************************************************************************
    'REVISED BY        -    PRASHANT RAJPAL
    'REVISION DATE     -    19/03/2015
    'REVISION HISTORY  -    CHANGES FOR LOGIC CHANGE FOR GENERATING TEMPORARY INVOICE: DUE TO DOUBLE ENTRY DATA
    'ISSUE ID           -     10777177  
    '**********************************************************************************************************************
    'REVISED BY        -    PRASHANT RAJPAL
    'REVISION DATE     -    24/04/2015
    'REVISION HISTORY  -    Changes for MTL Sharjah: Dr is not cr issue
    'ISSUE ID          -    10729758   
    '**********************************************************************************************************************
    'CREATED BY       : Parveen Kumar
    'CREATED ON       : 09 JUN 2015
    'DESCRIPTION      : eMPro- Declaration No. in Export Invoice
    'AGAINST ISSUE ID : 10826755
    '----------------------------------------------------------------
    'REVISED BY        -    PRASHANT RAJPAL
    'REVISION DATE     -    29/10/2015
    'REVISION HISTORY  -    Changes for MTL Sharjah: Dr is not cr issue
    'ISSUE ID          -    10871426   
    '**********************************************************************************************************************
    'REVISED BY        -    PRASHANT RAJPAL
    'REVISION DATE     -    25/11/2015
    'REVISION HISTORY  -    Credit term picked wrong 
    'ISSUE ID          -    10895403    
    '**********************************************************************************************************************

    Dim mintIndex As Short 'Declared To Hold The Form CountKD
    Dim mdblPrevQty() As Object 'to store prev quantity in edit mode
    Dim mdblToolCost() As Object 'to insert tool cost item wise
    Dim ArrExpDetails() As String
    Dim strExpDetails As String
    Dim strExpEditDetails As String
    Public mstrItemCode As String 'To Get The Value Of Item Code
    Dim mstrInvoiceType As String 'To Get The Value Of Invoice Type
    Dim mstrInvoiceSubType As String 'To Get The Value Of Invoice Sub Type
    Dim mstrAmmendmentNo As String 'To Get The Value Of Ammendment No.
    Dim mstrInvType As String 'To Get Value Of Inv Type From SalesChallan_Dtl
    Dim mstrInvSubType As String 'To Get Value Of Inv SubType From SalesChallan_Dtl
    Dim mstrUpdDispatchSql As String 'To Make Update Query For Dispatch_Qty From Daily/Monthly Mkt Schedule
    Dim mstrAmmNo As String
    Dim mstrRefNo As String
    Dim strupSalechallan As String
    Dim strupSaleDtl As String
    Dim strInvType As String
    Dim strInvSubType As String
    Dim msubTotal, mInvNo, mExDuty, mBasicAmt, mOtherAmt As Double
    Dim mstrupdateASNdtl As String
    Dim mstrupdateASNCumFig As String
    Dim STREXPDET As String
    Dim mInvAgstDispAdv As Boolean
    Dim mSchTypeArr() As String
    Dim mblnInvocieforMTL As Boolean
    Dim mstrCreditTermId As String
    Dim mblnInvoicelike_MTLsharjah As Boolean
    Dim mblncustomer_agstdispatchadvice As Boolean
    Dim mstrexportsotype As String = String.Empty



    Private Sub chkServiceInvFormat_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles chkServiceInvFormat.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                txtBankAc.Focus()
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
    Private Sub CmbInvSubType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbInvSubType.SelectedIndexChanged
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Validate invoice sub type as per invoice type Selected.
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Call SelectInvTypeSubTypeFromSaleConf((CmbInvType.Text), (CmbInvSubType.Text))
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmbInvSubType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles CmbInvSubType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                txtLocationCode.Focus()
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub CmbInvSubType_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbInvSubType.Leave
        '*******************************************************************************
        'Author             :   Manoj Kr. Vaish
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To set controls enabled/disabled Condition According to
        '                       Invoice type Selected
        'Comments           :   NA
        'Creation Date      :   07-Sep-2007 Issue ID 21054
        '*******************************************************************************
        Select Case CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                Select Case UCase(CmbInvSubType.Text)
                    Case "EXPORTS"
                        If InvAgstDispAdvise() = True And mblncustomer_agstdispatchadvice = True Then
                            txtRefNo.Enabled = False : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            CmdRefNoHelp.Enabled = False
                            txtDispAdvNo.Enabled = True : txtDispAdvNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            cmdDispAdvNo.Enabled = True
                        Else
                            txtRefNo.Enabled = True : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            CmdRefNoHelp.Enabled = True
                            txtDispAdvNo.Enabled = False : txtDispAdvNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            cmdDispAdvNo.Enabled = False
                        End If
                End Select
        End Select
    End Sub
    Private Sub CmbInvType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbInvType.SelectedIndexChanged
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Refresh invoice Sub type
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Call SelectInvoiceSubTypeFromSaleConf((CmbInvType.Text))
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmbInvType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles CmbInvType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        CmbInvSubType.Focus()
                End Select
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub CmbInvType_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbInvType.Leave
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To set controls enabled/disabled Condition According to
        '                       Invoice type Selected
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        Select Case CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                Select Case UCase(CmbInvType.Text)
                    Case "EXPORT INVOICE"
                        txtRefNo.Enabled = True : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdRefNoHelp.Enabled = True
                        txtAnnex.Enabled = False : txtAnnex.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdAnnexHelp.Enabled = False
                        txtExciseDuty.Enabled = False
                        txtExciseDuty.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        txtAddExciseDuty.Enabled = False
                        txtAddExciseDuty.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        ctlSVD.Enabled = False
                        ctlSVD.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        txtSalesTax.Enabled = False
                        Me.ctlInsurance.Enabled = False
                        Me.ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        Me.txtFreight.Enabled = False
                        Me.txtFreight.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        Me.txtSurcharge.Enabled = False
                        Me.txtSurcharge.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        txtSalesTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtSaleTaxType.Enabled = False : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        CmdSaleTaxType.Enabled = False
                        If InvAgstDispAdvise() = True And mblncustomer_agstdispatchadvice = True Then
                            txtRefNo.Enabled = False : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            CmdRefNoHelp.Enabled = False
                            txtDispAdvNo.Enabled = True : txtDispAdvNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            cmdDispAdvNo.Enabled = True
                        Else
                            txtRefNo.Enabled = True : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            CmdRefNoHelp.Enabled = True
                            txtDispAdvNo.Enabled = False : txtDispAdvNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            cmdDispAdvNo.Enabled = False
                        End If
                End Select
        End Select
    End Sub
    Private Sub CmbTransType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles CmbTransType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        txtVehNo.Focus()
                End Select
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub cmdAcCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAcCode.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Display Help From SaleTax Master
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strHelpString As String
        Dim strBankNo() As String

        Dim strSql As String = ""
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT

                strHelpString = "Select Bnk_BankId,Bnk_accNo from Gen_bankMaster where unit_code='" & gstrUNITID & "' and USE_IN_EXPORT_INVOICE = 1"


                strBankNo = ctlExportChallanEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelpString, "Bank Codes")
                If UBound(strBankNo) <= 0 Then Exit Sub
                Dim aa As String = strBankNo(0).ToString
                If strBankNo(0) = "0" Then
                    MsgBox("No Bank Code Available To Display.", MsgBoxStyle.Information, "empower") : txtBankAc.Text = "" : txtBankAc.Focus() : Exit Sub
                Else
                    txtBankAc.Text = strBankNo(0)
                    lblAcCodeDes.Text = strBankNo(1)
                End If
        End Select
        txtBankAc.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub

    Private Sub CmdChallanNo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdChallanNo.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Display Help For Invoice No.
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strHelpString As String
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Trim(txtLocationCode.Text) = "" Then
                    Call ConfirmWindow(10239, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO, 100)
                    txtLocationCode.Focus()
                    Exit Sub
                End If
                If Len(Trim(txtChallanNo.Text)) = 0 Then
                    strHelpString = ShowList(1, (txtChallanNo.MaxLength), "", "Doc_No", DateColumnNameInShowList("Invoice_Date", 1) & " As Invoice_Date", "SalesChallan_Dtl ", "and Invoice_Type ='EXP' AND Location_Code='" & Trim(txtLocationCode.Text) & "'")
                    If strHelpString = "-1" Then 'If No Record Found
                        Call ConfirmWindow(10253, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        txtChallanNo.Focus()
                    Else
                        txtChallanNo.Text = strHelpString
                    End If
                Else
                    strHelpString = ShowList(1, (txtChallanNo.MaxLength), txtChallanNo.Text, "Doc_No", DateColumnNameInShowList("Invoice_Date", 1) & " As Invoice_Date", "SalesChallan_Dtl ", "and Invoice_Type ='EXP' AND Location_Code='" & Trim(txtLocationCode.Text) & "'")
                    If strHelpString = "-1" Then 'If No Record Found
                        Call ConfirmWindow(10253, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        txtChallanNo.Focus()
                    Else
                        txtChallanNo.Text = strHelpString
                    End If
                End If
        End Select
        txtChallanNo.Focus()
        If Val(txtChallanNo.Text) > 99000000 Then
            Cmditems.Enabled = True
        Else
            Cmditems.Enabled = False
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdConsCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdConsCodeHelp.Click
        '*******************************************************************************
        'Author             :   Ashutosh Verma, Issue Id:19661
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Display Consignee code's Help
        'Comments           :   NA
        'Creation Date      :   19 Mar 2007
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strHelpString As String
        If Len(Trim(txtLocationCode.Text)) = 0 Then
            Call ConfirmWindow(10116, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            txtLocationCode.Focus()
            Exit Sub
        End If
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If Trim(mstrInvoiceType) = "EXP" Then
                    If Len(Trim(txtConsCode.Text)) = 0 Then
                        strHelpString = ShowList(1, (txtConsCode.MaxLength), "", "Customer_Code", "Cust_Name", "Customer_Mst", , "Consignee Code Help")
                        If strHelpString = "-1" Then 'If No Record Found
                            Call ConfirmWindow(10225, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Else
                            txtConsCode.Text = strHelpString
                        End If
                    Else
                        strHelpString = ShowList(1, (txtConsCode.MaxLength), txtConsCode.Text, "customer_code", "cust_name", "Customer_Mst", , "Consignee Code Help")
                        If strHelpString = "-1" Then 'If No Record Found
                            Call ConfirmWindow(10225, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Else
                            txtConsCode.Text = strHelpString
                        End If
                    End If
                End If
        End Select
        Call SelectDescriptionForField("Cust_Name", "Customer_Code", "Customer_Mst", lblConsCodeDes, (txtConsCode.Text))
        If txtConsCode.Enabled = True Then txtConsCode.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdCustCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdCustCodeHelp.Click
        'Function           :   To Display Customer code's Help
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strHelpString As String
        If Len(Trim(txtLocationCode.Text)) = 0 Then
            Call ConfirmWindow(10116, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            txtLocationCode.Focus()
            Exit Sub
        End If
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If Trim(mstrInvoiceType) = "EXP" Then
                    If Len(Trim(txtCustCode.Text)) = 0 Then
                        strHelpString = ShowList(1, (txtCustCode.MaxLength), "", "Customer_Code", "Cust_Name", "Customer_Mst", , "Customer Code Help")
                        If strHelpString = "-1" Then 'If No Record Found
                            Call ConfirmWindow(10225, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Else
                            txtCustCode.Text = strHelpString
                            'Issue ID 10127477
                            Call SetControlsforASNDetails(Trim(txtCustCode.Text))
                            'Issue ID 10127477 END 
                        End If
                    Else
                        strHelpString = ShowList(1, (txtCustCode.MaxLength), txtCustCode.Text, "customer_code", "cust_name", "Customer_Mst", , "Customer Code Help")
                        If strHelpString = "-1" Then 'If No Record Found
                            Call ConfirmWindow(10225, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Else
                            txtCustCode.Text = strHelpString
                            Call SetControlsforASNDetails(Trim(txtCustCode.Text))
                        End If
                    End If
                End If
        End Select
        Call SelectDescriptionForField("Cust_Name", "Customer_Code", "Customer_Mst", lblCustCodeDes, (txtCustCode.Text))
        Call SelectDescriptionForField("Cust_Name", "Customer_Code", "Customer_Mst", lblConsCodeDes, (txtConsCode.Text))
        mblncustomer_agstdispatchadvice = CBool(Find_Value("SELECT ENABLE_AGSTDISPADVICE FROM customer_mst WHERE UNIT_CODE='" & gstrUNITID & "' and customer_code='" & txtCustCode.Text & "'"))
        txtCustCode.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmdDispAdvNo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDispAdvNo.Click
        '-------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Function      : To Display help for Despatch Advise
        ' Datetime      : 12-Feb-2007
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim StrData() As String
        Dim strQry As String
        strQry = "SELECT cast(DOCNO as varchar(15)) as DocNo ,DOCDATE,CUST_REF,ISNULL(AMENDMENT_NO,'') AS AMENDMENT_NO"
        strQry = strQry & " FROM BAR_DISPATCHADVICE_HDR"
        strQry = strQry & " WHERE UNIT_CODE='" & gstrUNITID & "' AND ISNULL(INVOICENO,0)=0  AND CUSTOMERCODE= '" & Trim(txtCustCode.Text) & "' And Consignee_Code='" & Trim(txtConsCode.Text) & "' And Status = 0 "
        StrData = ctlExportChallanEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry, "Dispatch Advice Help", 1)
        If Not (UBound(StrData) <= 0) Then
            If (Len(StrData(0)) >= 1) And StrData(0) = "0" Then
                MsgBox("No Pending Dispatch Advise exist for selected Customer !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                Exit Sub
            Else
                'If mblnInvocieforMTL = False Or mblnInvoicelike_MTLsharjah = False Then
                If mblnInvocieforMTL = False Then
                    mstrRefNo = Trim(StrData(2))
                    mstrAmmNo = Trim(StrData(3))
                    txtRefNo.Text = mstrRefNo
                End If
                txtDispAdvNo.Text = Trim(StrData(0))
                Call txtDispAdvNo_Validating(txtDispAdvNo, New System.ComponentModel.CancelEventArgs(False))

            End If
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub CmdGrpChEnt_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles CmdGrpChEnt.ButtonClick
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   Code for ADD/EDIT/UPDATE/CANCEL/CLOSE
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        'Revised By         : Davinder Singh
        'Revised on         : 12-Feb-2007
        'Revision History   : Schedule Checking and Updation is done through Stored procedures
        '                     Concept of invoice agst Despatch Advise is introduced
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strSalesChallan As String
        Dim strSalesDtl As String
        Dim Description As String
        Dim intLoopcount As Short
        Dim varQuantity As Object
        Dim varDrgNo As Object
        Dim varItemCode As Object
        Dim varRate As Object
        Dim varCustMtrl As Object
        Dim varPacking As Object
        Dim varOthers As Object
        Dim varFromBox As Object
        Dim VarToBox As Object
        Dim PresQty As Object
        Dim rsCustItemMst As ClsResultSetDB
        Dim rsSaleConf As ClsResultSetDB
        Dim rsItemMst As ClsResultSetDB
        Dim rsSalesChallandtl As ClsResultSetDB
        Dim rsExternalsalesorder As ClsResultSetDB
        Dim strDispathAdvice As String
        Dim strMktScheduleCheck As String
        Dim strupdatebarmst As String
        Dim strInsertUpdateASNdtl As String
        STREXPDET = ""
        Dim varBinQty As Object
        Dim intLoop As Short
        Dim strStock_Location As String
        Dim updatearr() As String
        Dim strsql As String
        Dim varCustRef As Object
        Dim varAmendmentNo As Object
        Dim dbltotaltaxamount As Double
        Dim blnISSalesTaxRoundOff As Boolean
        Dim ldblTotalSaleTaxAmount As Double
        Dim intSaleTaxRoundOffDecimal As Short
        Dim strexportsotype As String

        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                ldblTotalSaleTaxAmount = 0
                strValues = ""
                strExpDetails = ""
                Call EnableControls(True, Me, True)
                Call SelectChallanNoFromSalesChallanDtl()
                txtChallanNo.Enabled = False : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                CmdChallanNo.Enabled = False : txtChallanNo.Enabled = False
                txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : CmdRefNoHelp.Enabled = False
                lblLocCodeDes.Text = "" : lblCustCodeDes.Text = ""
                Me.SpChEntry.Enabled = True
                CmbInvType.SelectedIndex = 0 : CmbInvSubType.SelectedIndex = 0 : CmbTransType.SelectedIndex = 0
                With Me.SpChEntry
                    .MaxRows = 1
                    .Row = 1 : .Row2 = 1 : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .Lock = False : .BlockMode = False
                End With
                If (Trim(CmbInvType.Text) = "EXPORT INVOICE") And (InvAgstDispAdvise() = True And mblncustomer_agstdispatchadvice = True) Then
                    txtRefNo.Enabled = False
                    txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    CmdRefNoHelp.Enabled = False
                    txtDispAdvNo.Enabled = True : txtDispAdvNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    cmdDispAdvNo.Enabled = True
                Else
                    txtRefNo.Enabled = True
                    txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    CmdRefNoHelp.Enabled = True
                    txtDispAdvNo.Enabled = False : txtDispAdvNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    cmdDispAdvNo.Enabled = False
                End If
                If GetMktCodeExecutionZone() = "NORTH" And GetPlantName() <> "HILEX" Then
                    txtConsCode.Enabled = False
                    txtConsCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    CmdConsCodeHelp.Enabled = False
                Else
                    txtConsCode.Enabled = True
                    txtConsCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    CmdConsCodeHelp.Enabled = True
                End If
                ctlInsurance.Enabled = True : ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtvesselnumber.Enabled = True : txtvesselnumber.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                CmbInvType.Visible = True : CmbInvSubType.Visible = True
                lblInvSubType.Visible = True : lblInvType.Visible = True
                lblDateDes.Text = VB6.Format(GetServerDate(), gstrDateFormat)
                With dtpDateDesc
                    .Value = ConvertToDate(lblDateDes.Text)
                    .Visible = True
                End With
                Call SetMaxLengthInSpread()
                Call ChangeCellTypeStaticText()
                dtpDateDesc.Focus()
                txtshippingaddcode.Text = ""
                txtboxShipadddesc.Text = ""

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                Call EnableControls(False, Me)
                rsSalesChallandtl = New ClsResultSetDB
                rsSalesChallandtl.GetResult("select Invoice_type from Saleschallan_dtl where unit_code='" & gstrUNITID & "' and doc_no = " & txtChallanNo.Text, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                SpChEntry.Enabled = True
                SpChEntry.Row = 1 : SpChEntry.Row2 = SpChEntry.MaxRows : SpChEntry.Col = 0 : SpChEntry.Col2 = SpChEntry.MaxCols
                SpChEntry.BlockMode = True : SpChEntry.Lock = False : SpChEntry.BlockMode = False
                If Len(Trim(txtDispAdvNo.Text)) > 0 Then
                    With SpChEntry
                        .Row = 1
                        .Row2 = .MaxRows
                        .Col = 1
                        .Col2 = 5
                        .BlockMode = True
                        .Lock = True
                        .BlockMode = False
                    End With
                End If
                Call SetMaxLengthInSpread()
                Call ChangeCellTypeStaticText()
                ReDim mdblPrevQty(SpChEntry.MaxRows - 1) ' To get value of Quantity in Arrey for updation in despatch
                For intLoop = 1 To SpChEntry.MaxRows
                    Call SpChEntry.GetText(5, intLoop, mdblPrevQty(intLoop - 1))
                Next
                chkServiceInvFormat.Enabled = True : txtBankAc.Enabled = True : txtBankAc.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdAcCode.Enabled = True
                txtRemarks.Enabled = True : txtRemarks.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtLorryNo.Enabled = True : txtLorryNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtOTLNo.Enabled = True : txtOTLNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                ctlInsurance.Enabled = True : ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtvesselnumber.Enabled = True : txtvesselnumber.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                'Samiksha shipaddress code changes
                txtshippingaddcode.Enabled = True : txtshippingaddcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtboxShipadddesc.Enabled = True : txtboxShipadddesc.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                CmdhelpShipAddCode.Enabled = True
                cmexport.Enabled = True : cmexport.Focus()
                rsSalesChallandtl.ResultSetClose()
                rsSalesChallandtl = Nothing
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                Select Case CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        blnISSalesTaxRoundOff = Find_Value("select SalesTax_Roundoff from sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'")
                        intSaleTaxRoundOffDecimal = Find_Value("select SalesTax_Roundoff_decimal from sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'")

                        If Not ValidateBeforeSave("ADD") Then
                            gblnCancelUnload = True
                            gblnFormAddEdit = True
                            Exit Sub
                        End If
                        lblDateDes.Text = VB6.Format(dtpDateDesc.Value, gstrDateFormat)
                        If QuantityCheck() = True Then
                            Exit Sub
                        End If
                        rsSaleConf = New ClsResultSetDB
                        rsSaleConf.GetResult("Select Invoice_Type,Sub_Type,stock_location from SaleConf where unit_code='" & gstrUNITID & "' and Description ='" & CmbInvType.Text & "'and Sub_type_Description ='" & Trim(CmbInvSubType.Text) & "' and datediff(dd,'" & getDateForDB(dtpDateDesc.Value) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(dtpDateDesc.Value) & "')<=0")
                        If rsSaleConf.GetNoRows > 0 Then
                            strStock_Location = rsSaleConf.GetValue("stock_location")
                        Else
                            strStock_Location = ""
                        End If
                        strSalesChallan = ""
                        Call SelectChallanNoFromSalesChallanDtl()
                        'strexportsotype = ""
                        'If UCase(CmbInvType.Text) = "EXPORT INVOICE" And txtRefNo.Text <> "" Then
                        '    strexportsotype = Find_Value("Select exportsotype from cust_ord_hdr where UNIT_CODE='" + gstrUNITID + "' AND  account_code='" & Trim(txtCustCode.Text) & "' and cust_ref='" & txtRefNo.Text.Trim & "' and amendment_no='" & mstrAmmNo & "' and active_flag='a'")
                        'End If
                        strSalesChallan = Trim(strSalesChallan) & "Insert into SalesChallan_dtl (Unit_Code,Location_Code,from_location,Doc_No,Suffix,Consignee_Code,Lorry_No,OTL_No,RefChallan,"
                        strSalesChallan = strSalesChallan & "Transport_Type,Vehicle_No,Vessel_Flight_number,"
                        strSalesChallan = strSalesChallan & "From_Station,To_Station,Account_Code,Cust_Ref,"
                        strSalesChallan = strSalesChallan & "Amendment_No,Bill_Flag,Form3,Carriage_Name,"
                        strSalesChallan = strSalesChallan & "Year,Insurance,"
                        strSalesChallan = strSalesChallan & "Frieght_Tax,invoice_Type,Ref_Doc_No,Cust_Name,"
                        strSalesChallan = strSalesChallan & "Sub_Category,"
                        strSalesChallan = strSalesChallan & "Annex_no,invoice_Date,Invoice_Time,Ent_dt,"
                        'Samiksha shipadress code changes
                        If UCase(CmbInvType.Text) = "EXPORT INVOICE" And txtRefNo.Text <> "" Then
                            strSalesChallan = strSalesChallan & "Ent_UserId,Upd_dt,Upd_UserId,print_flag,SalesTax_type,Sales_Tax_Amount ,salestax_per, total_amount,ServiceInvoiceformatExport,CustBankID,Remarks, PrintExciseFormat, FreshCrRecd,ExportSotype,ShipAddress_Code ) Values ('" & gstrUNITID & "','" & Trim(txtLocationCode.Text)
                        Else
                            strSalesChallan = strSalesChallan & "Ent_UserId,Upd_dt,Upd_UserId,print_flag,SalesTax_type,Sales_Tax_Amount ,salestax_per, total_amount,ServiceInvoiceformatExport,CustBankID,Remarks, PrintExciseFormat, FreshCrRecd) Values ('" & gstrUNITID & "','" & Trim(txtLocationCode.Text)
                        End If

                        strSalesChallan = strSalesChallan & "', '" & Trim(strStock_Location) & "'," & Trim(txtChallanNo.Text) & ",'', '" & Trim(txtConsCode.Text) & "','" & Trim(Me.txtLorryNo.Text) & "','" & Trim(Me.txtOTLNo.Text) & "','" & Trim(txtRefChallanNo.Text) & "' "
                        strSalesChallan = strSalesChallan & ",'" & Mid(Trim(CmbTransType.Text), 1, 1) & "', '" & Trim(txtVehNo.Text) & "','"
                        strSalesChallan = strSalesChallan & Trim(txtvesselnumber.Text) & "','"
                        strSalesChallan = strSalesChallan & "','','" & Trim(txtCustCode.Text)
                        strSalesChallan = strSalesChallan & "','" & Trim(txtRefNo.Text) & "','" & Trim(mstrAmmNo) & "',0"
                        strSalesChallan = strSalesChallan & ",'','" & Trim(txtCarrServices.Text)
                        strSalesChallan = strSalesChallan & "','" & Trim(CStr(Year(dtpDateDesc.Value))) & "',"
                        strSalesChallan = strSalesChallan & "IsNull(" & Val(ctlInsurance.Text) & ", 0)"
                        strSalesChallan = strSalesChallan & ",IsNull(" & Val(txtFreight.Text) & ", 0),'" & Trim(rsSaleConf.GetValue("Invoice_type")) & "','"
                        strSalesChallan = strSalesChallan & Trim(txtAnnex.Text) & "','" & Trim(lblCustCodeDes.Text) & "',"
                        strSalesChallan = strSalesChallan & "'" & Trim(rsSaleConf.GetValue("Sub_Type")) & "',"
                        strSalesChallan = strSalesChallan & "0,'" & getDateForDB(lblDateDes.Text) & "',substring(convert(varchar(20),Getdate()),13,len(getdate())),getdate(),'" & mP_User & "',getdate(),'" & mP_User & "',0"
                        strSalesChallan = strSalesChallan & ",'" & txtSaleTaxType.Text & "','"
                        dbltotaltaxamount = 0
                        If mblnInvocieforMTL = True Then
                            If blnISSalesTaxRoundOff Then
                                dbltotaltaxamount = System.Math.Round((CalculateTotalInvoiceAmount() * GetTaxRate(txtSaleTaxType.Text, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='VAT'")) / 100)
                            Else
                                dbltotaltaxamount = System.Math.Round((CalculateTotalInvoiceAmount() * GetTaxRate(txtSaleTaxType.Text, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='VAT'")) / 100, intSaleTaxRoundOffDecimal)
                            End If
                        Else
                            dbltotaltaxamount = 0
                        End If

                        ldblTotalSaleTaxAmount = dbltotaltaxamount
                        txtSalesTax.Text = ldblTotalSaleTaxAmount
                        strSalesChallan = strSalesChallan & ldblTotalSaleTaxAmount & "'," & GetTaxRate(txtSaleTaxType.Text, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='VAT'") & "," & CalculateTotalInvoiceAmount()

                        If chkServiceInvFormat.CheckState = System.Windows.Forms.CheckState.Checked Then
                            strSalesChallan = strSalesChallan & ",1"
                        Else
                            strSalesChallan = strSalesChallan & ",0"
                        End If
                        'Samiksha ship address code changes
                        If UCase(CmbInvType.Text) = "EXPORT INVOICE" And txtRefNo.Text <> "" Then
                            strSalesChallan = strSalesChallan & ",'" & Trim(txtBankAc.Text) & "','" & Trim(txtRemarks.Text) & "', 0 ,0,'" & mstrexportsotype & "','" & Trim(txtshippingaddcode.Text) & "')"
                        Else
                            strSalesChallan = strSalesChallan & ",'" & Trim(txtBankAc.Text) & "','" & Trim(txtRemarks.Text) & "', 0 ,0 )"
                        End If

                        If Len(Trim(strValues)) = 0 Then
                            MsgBox("Please Select Export Details", MsgBoxStyle.Information, "empower")
                            cmexport.Focus()
                            Exit Sub
                        ElseIf Len(Trim(strValues)) = 36 Then
                            MsgBox("Please Select Export Details", MsgBoxStyle.Information, "empower")
                            cmexport.Focus()
                            Exit Sub
                        Else
                            updatearr = Split(strValues, "§")
                            strsql = strsql & "update saleschallan_dtl set "
                            strsql = strsql & "Frieght_Amount=" & updatearr(16) & ","
                            strsql = strsql & "Currency_Code='" & updatearr(0) & "' ,"
                            strsql = strsql & "Nature_of_Contract='" & updatearr(7) & "' ,"
                            strsql = strsql & "OriginStatus='" & updatearr(1) & "' ,"
                            strsql = strsql & "Ctry_destination_goods='" & updatearr(2) & "' ,"
                            strsql = strsql & "Delivery_Terms='" & updatearr(11) & "' ,"
                            strsql = strsql & "Payment_Terms='" & updatearr(12) & "' ,"
                            strsql = strsql & "Pre_carriage_by='" & updatearr(3) & "' ,"
                            strsql = strsql & "Receipt_Precarriage_at='" & updatearr(4) & "' ,"
                            strsql = strsql & "Vessel_flight_number= '" & Trim(txtvesselnumber.Text) & "',"
                            strsql = strsql & "Port_of_loading='" & updatearr(5) & "' ,"
                            strsql = strsql & "Port_of_discharge='" & updatearr(6) & "' ,"
                            strsql = strsql & "Final_destination='" & updatearr(8) & "' ,"
                            strsql = strsql & "Mode_of_Shipment='" & updatearr(9) & "' ,"
                            strsql = strsql & "DISPATCH_MODE='" & updatearr(10) & "' ,"
                            strsql = strsql & "Buyer_Description_of_Goods='" & updatearr(13) & "' ,"
                            strsql = strsql & "Invoice_Description_of_EPC='" & updatearr(14) & "' ,"
                            strsql = strsql & "Exchange_Rate='" & updatearr(15) & "',"
                            strsql = strsql & "Exchange_Date='" & updatearr(17) & "',"
                            strsql = strsql & "Other_ref ='" & updatearr(18) & "',"
                            strsql = strsql & "buyer_id ='" & updatearr(19) & "'"
                            If gstrUNITID = "WCS" Then
                                strsql = strsql & ",total_amount =" & CalculateTotalInvoiceAmount() + updatearr(16)
                            Else
                                strsql = strsql & ",total_amount =" & CalculateTotalInvoiceAmount()
                            End If

                            'strsql = strsql & ",total_amount =" & CalculateTotalInvoiceAmount()
                        End If
                        strsql = strsql & " where unit_code='" & gstrUNITID & "' and "
                        strsql = strsql & " doc_no=" & txtChallanNo.Text & " and "
                        strsql = strsql & " suffix='' "
                        STREXPDET = STREXPDET & " Insert into EXPORT_SALES_EXTRA_DETAIL(UNIT_CODE,Unt_CodeID,Doc_No,Advance_lice_No,Pallet_Length,Pallet_width,Pallet_Height,Pallet_Total,ARE_NO,Net_Weight,Gross_Weight,Export_Type,Volume_weight,DRAWBACK_TYPE,HS_CODE,CommodityType)"
                        STREXPDET = STREXPDET & " Values('" & gstrUNITID & "','" & Trim(txtLocationCode.Text) & "'," & Trim(txtChallanNo.Text) & ",'" & updatearr(20) & "'," & Val(updatearr(21)) & "," & Val(updatearr(22)) & "," & Val(updatearr(23)) & ","
                        STREXPDET = STREXPDET & Val(updatearr(24)) & ",'" & updatearr(25) & "'," & Val(updatearr(26)) & ","
                        'Samiksha commodity changes
                        STREXPDET = STREXPDET & Val(updatearr(27)) & ",'" & updatearr(28) & "'," & Val(updatearr(29)) & ",'" & updatearr(30) & "','" & updatearr(31) & "','" & updatearr(34) & "')"
                        strSalesDtl = ""
                        For intLoopcount = 1 To SpChEntry.MaxRows
                            varItemCode = Nothing
                            varFromBox = Nothing
                            VarToBox = Nothing
                            varDrgNo = Nothing
                            varRate = Nothing
                            varCustMtrl = Nothing
                            varQuantity = Nothing
                            varPacking = Nothing
                            varOthers = Nothing
                            varBinQty = Nothing
                            varCustRef = Nothing
                            varAmendmentNo = Nothing
                            Call SpChEntry.GetText(1, intLoopcount, varItemCode)
                            Call SpChEntry.GetText(2, intLoopcount, varDrgNo)
                            Call SpChEntry.GetText(3, intLoopcount, varRate)
                            Call SpChEntry.GetText(4, intLoopcount, varCustMtrl)
                            Call SpChEntry.GetText(5, intLoopcount, varQuantity)
                            Call SpChEntry.GetText(6, intLoopcount, varPacking)
                            Call SpChEntry.GetText(7, intLoopcount, varOthers)
                            Call SpChEntry.GetText(8, intLoopcount, varFromBox)
                            Call SpChEntry.GetText(9, intLoopcount, VarToBox)

                            'If mblnInvocieforMTL = True Or (mblnInvoicelike_MTLsharjah = True And mblncustomer_like_MTLsharjah = True) Then
                            If mblnInvocieforMTL = True Or (InvAgstDispAdvise() = True And mblncustomer_agstdispatchadvice = True) Then
                                Call SpChEntry.GetText(10, intLoopcount, varBinQty)
                                Call SpChEntry.GetText(11, intLoopcount, varCustRef)
                                Call SpChEntry.GetText(12, intLoopcount, varAmendmentNo)
                            Else
                                Call SpChEntry.GetText(10, intLoopcount, varBinQty)
                            End If
                            rsCustItemMst = New ClsResultSetDB
                            rsItemMst = New ClsResultSetDB
                            rsExternalsalesorder = New ClsResultSetDB

                            rsItemMst.GetResult("Select Description,Cons_measure_code from Item_Mst where UNIT_CODE='" & gstrUNITID & "' and Item_Code ='" & Trim(varItemCode) & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

                            rsCustItemMst.GetResult("Select Drg_desc,Decl_No from CustItem_Mst where UNIT_CODE='" & gstrUNITID & "' and Account_code ='" & Trim(txtCustCode.Text) & "'and Cust_DrgNo='" & varDrgNo & "'and Item_code ='" & varItemCode & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            rsExternalsalesorder.GetResult("SELECT external_salesorder_no FROM cust_ord_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_code ='" & Trim(txtCustCode.Text) & "'and Cust_DrgNo='" & varDrgNo & "'and Item_code ='" & varItemCode & "'and active_flag='A' and cust_ref='" & Me.txtRefNo.Text & "' ", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

                            strSalesDtl = Trim(strSalesDtl) & "Insert into sales_Dtl(UNIT_CODE,Location_Code,Doc_No,Suffix,Item_Code,Sales_Quantity,"
                            strSalesDtl = strSalesDtl & "From_Box,To_Box,Rate,Packing,Others,Cust_Mtrl,"
                            strSalesDtl = strSalesDtl & "Year,Cust_Item_Code,Cust_Item_Desc,Sales_Decl_No,Tool_Cost,Measure_Code,"
                            'If mblnInvocieforMTL = True Or (mblnInvoicelike_MTLsharjah = True And mblncustomer_like_MTLsharjah = True) Then
                            If mblnInvocieforMTL = True Or (InvAgstDispAdvise() = True And mblncustomer_agstdispatchadvice = True) Then
                                strSalesDtl = strSalesDtl & "Cust_ref,amendment_no,"
                            End If
                            '10826755--Starts
                            'If DataExist("SELECT SOUPLD_FOREXCURRENCY FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & Trim(txtCustCode.Text) & "'") = True Then
                            If (DataExist("SELECT SOUPLD_FOREXCURRENCY FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & Trim(txtCustCode.Text) & "'") = False) Or (gstrUNITID = "MS1") Then
                                strSalesDtl = strSalesDtl & "Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,Basic_Amount,Accessible_amount,BinQuantity) values ('" & gstrUNITID & "','" & Trim(txtLocationCode.Text) & "',"

                            Else
                                strSalesDtl = strSalesDtl & "Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,Basic_Amount,Accessible_amount,BinQuantity,EXTERNAL_SALESORDER_NO ) values ('" & gstrUNITID & "','" & Trim(txtLocationCode.Text) & "',"
                            End If

                            strSalesDtl = strSalesDtl & Trim(txtChallanNo.Text) & ",'','" & Trim(varItemCode) & "'," & Val(varQuantity) & ","
                            strSalesDtl = strSalesDtl & Val(varFromBox) & "," & Val(VarToBox) & "," & Val(varRate) & ","
                            strSalesDtl = strSalesDtl & Val(varPacking) & "," & Val(varOthers) & "," & Val(varCustMtrl) & ","
                            strSalesDtl = strSalesDtl & Trim(CStr(Year(dtpDateDesc.Value))) & ",'" & Trim(varDrgNo) & "','" & IIf((Len(Trim(rsCustItemMst.GetValue("Drg_Desc"))) <= 0 Or Trim(CStr(rsCustItemMst.GetValue("dRG_DESC") = "Unknown"))), Trim(rsItemMst.GetValue("Description")), Trim(rsCustItemMst.GetValue("Drg_Desc"))) & "','" & Trim(rsCustItemMst.GetValue("Decl_No")) & "', "
                            '10826755--Ends
                            If CmbInvType.Text = "NORMAL INVOICE" Then
                                strSalesDtl = strSalesDtl & mdblToolCost(intLoopcount - 1) & ",',getdate(),'"
                            Else
                                'If Not (mblnInvocieforMTL = True Or mblnInvoicelike_MTLsharjah = True) Then
                                If Not ((mblnInvocieforMTL = True) Or (InvAgstDispAdvise() = True And mblncustomer_agstdispatchadvice = True)) Then
                                    strSalesDtl = strSalesDtl & "0,'" & Trim(rsItemMst.GetValue("Cons_Measure_Code")) & "',getdate(),'"
                                Else
                                    strSalesDtl = strSalesDtl & "0,'" & Trim(rsItemMst.GetValue("Cons_Measure_Code")) & "','" & Trim(varCustRef) & "','" & Trim(varAmendmentNo) & "',getdate() ,'"
                                End If
                            End If
                            If (DataExist("SELECT SOUPLD_FOREXCURRENCY FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & Trim(txtCustCode.Text) & "'") = False) Or (gstrUNITID = "MS1") Then
                                strSalesDtl = strSalesDtl & Trim(mP_User) & "',getdate(),'" & Trim(mP_User) & "'," & (Val(varRate) * Val(varQuantity)) & "," & (Val(varRate) * Val(varQuantity)) & "," & Val(varBinQty) & ")" & vbCrLf

                            Else
                                strSalesDtl = strSalesDtl & Trim(mP_User) & "',getdate(),'" & Trim(mP_User) & "'," & (Val(varRate) * Val(varQuantity)) & "," & (Val(varRate) * Val(varQuantity)) & "," & Val(varBinQty) & ",'" & rsExternalsalesorder.GetValue("External_salesorder_NO").ToString & "' )" & vbCrLf
                            End If

                            rsCustItemMst.ResultSetClose()
                            rsCustItemMst = Nothing
                            rsItemMst.ResultSetClose()
                            rsItemMst = Nothing

                            If AllowASNTextFileGeneration(Trim(txtCustCode.Text)) = True Then
                                strInsertUpdateASNdtl = Trim(strInsertUpdateASNdtl) & "INSERT INTO MKT_ASN_INVDTL(unit_code,Doc_no,Cust_PlantCode,ARL_Code,ASN_Status,Cust_Part_Code,Cummulative_Qty)values('" & gstrUNITID & "' ,'"
                                strInsertUpdateASNdtl = strInsertUpdateASNdtl & Trim(txtChallanNo.Text) & "','" & Trim(txtPlantCode.Text) & "','" & Trim(txtActualReceivingLoc.Text) & "',0,'"
                                strInsertUpdateASNdtl = strInsertUpdateASNdtl & Trim(varDrgNo) & "','" & Val(CStr(varQuantity)) & "')" & vbCrLf

                                If CheckExistanceOfFieldData(Trim(varDrgNo), "cust_part_code", "MKT_ASN_CUMFIG", "(unit_code='" & gstrUNITID & "' and cust_part_code='" & Trim(varDrgNo) & "' and cust_PlantCode='" & Trim(txtPlantCode.Text) & "')") = False Then
                                    strInsertUpdateASNdtl = strInsertUpdateASNdtl & "INSERT INTO MKT_ASN_CUMFIG(Unit_code,Cust_Part_Code,Cust_PlantCode,Cummulative_Qty) VALUES('" & gstrUNITID & "','" & Trim(varDrgNo) & "','" & Trim(txtPlantCode.Text) & "',0)" & vbCrLf
                                End If

                            End If

                        Next

                        'If AllowASNTextFileGeneration(Trim(txtCustCode.Text)) = True Then
                        '    strInsertUpdateASNdtl = Trim(strInsertUpdateASNdtl) & "INSERT INTO MKT_ASN_INVDTL(unit_code,Doc_no,Cust_PlantCode,ARL_Code,ASN_Status,Cust_Part_Code,Cummulative_Qty)values('" & gstrUNITID & "' ,'"
                        '    strInsertUpdateASNdtl = strInsertUpdateASNdtl & Trim(txtChallanNo.Text) & "','" & Trim(txtPlantCode.Text) & "','" & Trim(txtActualReceivingLoc.Text) & "',0,'"
                        '    strInsertUpdateASNdtl = strInsertUpdateASNdtl & Trim(varDrgNo) & "','" & Val(CStr(varQuantity)) & "')" & vbCrLf

                        '    If CheckExistanceOfFieldData(Trim(varDrgNo), "cust_part_code", "MKT_ASN_CUMFIG", " (unit_code='" & gstrUNITID & "' and cust_part_code='" & Trim(varDrgNo) & "' and cust_PlantCode='" & Trim(txtPlantCode.Text) & "')") = False Then
                        '        strInsertUpdateASNdtl = strInsertUpdateASNdtl & "INSERT INTO MKT_ASN_CUMFIG(Unit_code,Cust_Part_Code,Cust_PlantCode,Cummulative_Qty) VALUES('" & gstrUNITID & "','" & Trim(varDrgNo) & "','" & Trim(txtPlantCode.Text) & "',0)" & vbCrLf
                        '    End If
                        'End If
                        If Len(Trim(txtDispAdvNo.Text)) > 0 Then
                            strDispathAdvice = "UPDATE BAR_DISPATCHADVICE_HDR"
                            strDispathAdvice = strDispathAdvice & " SET INVOICENO = '" & (txtChallanNo.Text) & "'"
                            strDispathAdvice = strDispathAdvice & " WHERE UNIT_CODE='" & gstrUNITID & "' AND DOCNO='" & Trim(txtDispAdvNo.Text) & "'"
                            strDispathAdvice = strDispathAdvice & " AND CUSTOMERCODE= '" & Trim(txtCustCode.Text) & "'"

                            strupdatebarmst = "UPDATE bar_palette_mst"
                            strupdatebarmst = strupdatebarmst & " SET INVOICE_NO = '" & (txtChallanNo.Text) & "'"
                            strupdatebarmst = strupdatebarmst & " WHERE UNIT_CODE='" & gstrUNITID & "' and DISPATCH_ADVICE='" & Trim(txtDispAdvNo.Text) & "'"
                            strupdatebarmst = strupdatebarmst & " AND CUSTOMER_CODE= '" & Trim(txtCustCode.Text) & "'"
                        End If
                        rsSaleConf.ResultSetClose()
                        rsSaleConf = Nothing
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If Not ValidateBeforeSave("EDIT") Then
                            gblnCancelUnload = True
                            gblnFormAddEdit = True
                            Exit Sub
                        End If
                        If QuantityCheck() Then
                            Exit Sub
                        End If
                        If Len(Trim(strValues)) = 0 Then
                            strValues = strExpDetails
                        End If
                        updatearr = Split(strValues, "§")
                        strSalesChallan = ""
                        strSalesChallan = "Update SalesChallan_Dtl Set "
                        strSalesChallan = strSalesChallan & "Insurance = " & Val(ctlInsurance.Text)
                        strSalesChallan = strSalesChallan & Val(txtSurcharge.Text) & ",Frieght_Tax=" & Val(txtFreight.Text)
                        strSalesChallan = strSalesChallan & ",Frieght_Amount=" & updatearr(16) & ","
                        strSalesChallan = strSalesChallan & " Consignee_Code='" & Trim(txtConsCode.Text) & "',"
                        strSalesChallan = strSalesChallan & " Lorry_No='" & Trim(txtLorryNo.Text) & "',"
                        strSalesChallan = strSalesChallan & " OTL_No='" & Trim(txtOTLNo.Text) & "',"
                        strSalesChallan = strSalesChallan & " RefChallan='" & Trim(txtRefChallanNo.Text) & "',"
                        strSalesChallan = strSalesChallan & "Currency_Code='" & updatearr(0) & "' ,"
                        strSalesChallan = strSalesChallan & "Nature_of_Contract='" & updatearr(7) & "' ,"
                        strSalesChallan = strSalesChallan & "OriginStatus='" & updatearr(1) & "' ,"
                        strSalesChallan = strSalesChallan & "Ctry_destination_goods='" & updatearr(2) & "' ,"
                        strSalesChallan = strSalesChallan & "Delivery_Terms='" & updatearr(11) & "' ,"
                        strSalesChallan = strSalesChallan & "Payment_Terms='" & updatearr(12) & "' ,"
                        strSalesChallan = strSalesChallan & "Pre_carriage_by='" & QuoteString(updatearr(3)) & "' ,"
                        strSalesChallan = strSalesChallan & "Receipt_Precarriage_at='" & updatearr(4) & "' ,"
                        strSalesChallan = strSalesChallan & "Vessel_flight_number= '" & Trim(txtvesselnumber.Text) & "',"
                        strSalesChallan = strSalesChallan & "Port_of_loading='" & updatearr(5) & "' ,"
                        strSalesChallan = strSalesChallan & "Port_of_discharge='" & updatearr(6) & "' ,"
                        strSalesChallan = strSalesChallan & "Final_destination='" & updatearr(8) & "' ,"
                        strSalesChallan = strSalesChallan & "Mode_of_Shipment='" & updatearr(9) & "' ,"
                        strSalesChallan = strSalesChallan & "DISPATCH_MODE='" & updatearr(10) & "' ,"
                        strSalesChallan = strSalesChallan & "Buyer_Description_of_Goods='" & updatearr(13) & "' ,"
                        strSalesChallan = strSalesChallan & "Invoice_Description_of_EPC='" & updatearr(14) & "' ,"
                        strSalesChallan = strSalesChallan & "Exchange_Date='" & updatearr(17) & "' ,"
                        strSalesChallan = strSalesChallan & "Exchange_Rate='" & Val(updatearr(15)) & "', "
                        strSalesChallan = strSalesChallan & "other_ref ='" & updatearr(18) & "', "
                        strSalesChallan = strSalesChallan & "buyer_id ='" & updatearr(19) & "' "
                        If gstrUNITID = "WCS" Then
                            strSalesChallan = strSalesChallan & ",total_amount =" & CalculateTotalInvoiceAmount() + updatearr(16)
                        Else
                            strSalesChallan = strSalesChallan & ",total_amount =" & CalculateTotalInvoiceAmount()
                        End If
                        ' strSalesChallan = strSalesChallan & ",total_amount =" & CalculateTotalInvoiceAmount()
                        If chkServiceInvFormat.CheckState = System.Windows.Forms.CheckState.Checked Then
                            strSalesChallan = strSalesChallan & ",ServiceInvoiceformatExport = 1"
                        Else
                            strSalesChallan = strSalesChallan & ",ServiceInvoiceformatExport = 0"
                        End If
                        strSalesChallan = strSalesChallan & ",CustBankID = '" & Trim(txtBankAc.Text) & "'"
                        strSalesChallan = strSalesChallan & ",Remarks = '" & Trim(txtRemarks.Text) & "'"
                        strSalesChallan = strSalesChallan & ",ShipAddress_Code = '" & Trim(txtshippingaddcode.Text) & "'"
                        strSalesChallan = strSalesChallan & " where UNIT_CODE='" & gstrUNITID & "' and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
                        strSalesChallan = strSalesChallan & " and Doc_No =" & Val(txtChallanNo.Text)
                        If updatearr.GetUpperBound(0).ToString <= 19 Then
                            MsgBox("Please Select Export Details", MsgBoxStyle.Information, ResolveResString(100))
                            cmexport.Focus()
                            Exit Sub
                        Else
                            STREXPDET = STREXPDET & " update EXPORT_SALES_EXTRA_DETAIL set Advance_lice_No = '" & updatearr(20) & "',Pallet_Length = " & updatearr(21) & ","
                            STREXPDET = STREXPDET & " Pallet_width = " & updatearr(22) & ",Pallet_Height = " & updatearr(23) & ",Pallet_Total = " & updatearr(24) & ","
                            STREXPDET = STREXPDET & " ARE_NO = '" & updatearr(25) & "' ,Net_Weight = " & updatearr(26) & ",Gross_Weight = " & updatearr(27) & ",Export_Type = '" & updatearr(28) & "'"
                            STREXPDET = STREXPDET & " ,Volume_weight = " & updatearr(29) & " "
                            STREXPDET = STREXPDET & " ,DRAWBACK_TYPE = '" & updatearr(30) & "',HS_CODE='" & updatearr(31) & "',CommodityType='" & updatearr(34) & "'"
                            STREXPDET = STREXPDET & " where UNIT_CODE='" & gstrUNITID & "' and Unt_CodeID = '" & Trim(txtLocationCode.Text) & "'"
                            STREXPDET = STREXPDET & " and Doc_No = " & Val(txtChallanNo.Text)
                        End If
                        strSalesDtl = ""
                        For intLoopcount = 1 To SpChEntry.MaxRows
                            varQuantity = Nothing
                            varDrgNo = Nothing
                            varRate = Nothing
                            Call SpChEntry.GetText(5, intLoopcount, varQuantity)
                            Call SpChEntry.GetText(2, intLoopcount, varDrgNo)
                            Call SpChEntry.GetText(3, intLoopcount, varRate)
                            strSalesDtl = Trim(strSalesDtl) & "Update Sales_dtl set Sales_Quantity = " & Val(varQuantity) & ","
                            strSalesDtl = Trim(strSalesDtl) & "basic_amount=" & (Val(varQuantity) * Val(varRate)) & ","
                            strSalesDtl = Trim(strSalesDtl) & "Accessible_amount=" & (Val(varQuantity) * Val(varRate))
                            strSalesDtl = Trim(strSalesDtl) & " where UNIT_CODE='" & gstrUNITID & "' and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
                            strSalesDtl = Trim(strSalesDtl) & " and Doc_No =" & Val(txtChallanNo.Text) & " and Cust_Item_Code='"
                            strSalesDtl = Trim(strSalesDtl) & Trim(varDrgNo) & "'" & vbCrLf
                        Next
                        If AllowASNTextFileGeneration(Trim(txtCustCode.Text)) = True Then
                            strInsertUpdateASNdtl = Trim(strInsertUpdateASNdtl) & "UPDATE MKT_ASN_INVDTL SET Cummulative_Qty=" & Val(CStr(varQuantity)) & " WHERE UNIT_CODE='" & gstrUNITID & "' AND DOC_NO='" & Val(txtChallanNo.Text) & "' and Cust_part_Code='" & Trim(varDrgNo) & "'" & vbCrLf
                        End If
                End Select

                With mP_Connection
                    .BeginTrans()
                    .Execute(strSalesChallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    .Execute(strSalesDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                        .Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    .Execute(STREXPDET, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    If Len(Trim(strDispathAdvice)) > 0 Then .Execute(strDispathAdvice, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    If Len(Trim(strupdatebarmst)) > 0 Then .Execute(strupdatebarmst, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    If Len(strInsertUpdateASNdtl) > 0 Then
                        .Execute(strInsertUpdateASNdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    If GetPlantName() <> "HILEX" Then
                        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                            If UpdateMktSchedules("+") = False Then
                                .RollbackTrans()
                                GoTo ErrHandler
                            End If
                        ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                            If UpdateMktSchedules("-") = False Then
                                .RollbackTrans()
                                GoTo ErrHandler
                            End If
                            If UpdateMktSchedules("+") = False Then
                                .RollbackTrans()
                                GoTo ErrHandler
                            End If
                        End If
                    End If
                    .CommitTrans()
                    '' Add by priti for Hilex to add HSN Code on 13 May 2025
                    If GetPlantName() = "HILEX" Then
                        SqlConnectionclass.ExecuteNonQuery("update DTL SET  DTL.HSNSACCODE=SO.HSNSACCODE " &
                        "from saleschallan_dtl HDR (NOLOCK) INNER JOIN sales_dtl DTL  (NOLOCK) ON " &
                        "HDR.doc_no = DTL.doc_no And " &
                        "HDR.unit_code = DTL.unit_code " &
                        "INNER JOIN CUST_ORD_DTL SO (NOLOCK) ON " &
                        "HDR.UNIT_CODE=SO.UNIT_CODE AND " &
                        "HDR.Account_Code=SO.Account_Code AND " &
                        "HDR.Cust_Ref=SO.Cust_Ref AND  " &
                        "HDR.Amendment_No=SO.Amendment_No AND " &
                        "DTL.ITEM_CODE=SO.ITEM_CODE AND " &
                        "DTL.Cust_Item_Code=SO.Cust_DrgNo " &
                        "WHERE HDR.UNIT_CODE='" & gstrUNITID & "' and " &
                        "HDR.invoice_type='EXP' and isnull(DTL.HSNSACCODE,'')='' " &
                        "And SO.HSNSACCODE <> '' and HDR.Doc_No='" & (txtChallanNo.Text) & "'")
                        '' Ends here by priti for Hilex to add HSN Code on 13 May 2025
                    End If
                End With
                Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                CmdGrpChEnt.Revert()
                Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                gblnCancelUnload = False : gblnFormAddEdit = False
                Call EnableControls(False, Me)
                SpChEntry.Enabled = True
                SpChEntry.Row = 1 : SpChEntry.Row2 = SpChEntry.MaxRows : SpChEntry.Col = 0 : SpChEntry.Col2 = SpChEntry.MaxCols
                SpChEntry.BlockMode = True : SpChEntry.Lock = True : SpChEntry.BlockMode = False
                CmbInvType.Enabled = True : CmbInvSubType.Enabled = True
                txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                CmdLocCodeHelp.Enabled = True : CmdChallanNo.Enabled = True
                With dtpDateDesc
                    lblDateDes.Text = VB6.Format(.Value, gstrDateFormat)
                    .Visible = False
                End With
                txtLocationCode.Focus()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                Call frmEXPTRN0010_SOUTH_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
                If ConfirmWindow(10054, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    strDispathAdvice = "UPDATE BAR_DISPATCHADVICE_HDR"
                    strDispathAdvice = strDispathAdvice & " SET InvoiceNo = Null"
                    strDispathAdvice = strDispathAdvice & " WHERE UNIT_CODE='" & gstrUNITID & "' AND DOCNO ='" & Trim(txtDispAdvNo.Text) & "'"
                    strDispathAdvice = strDispathAdvice & " AND CUSTOMERCODE = '" & Trim(txtCustCode.Text) & "'"

                    strupdatebarmst = "UPDATE bar_palette_mst"
                    strupdatebarmst = strupdatebarmst & " SET INVOICE_NO = Null"
                    strupdatebarmst = strupdatebarmst & " WHERE UNIT_CODE='" & gstrUNITID & "' AND DISPATCH_ADVICE='" & Trim(txtDispAdvNo.Text) & "'"
                    strupdatebarmst = strupdatebarmst & " AND CUSTOMER_CODE= '" & Trim(txtCustCode.Text) & "'"
                    With SpChEntry
                        For intLoopcount = 1 To .MaxRows Step 1
                            .Row = intLoopcount
                            .Col = 5

                            mdblPrevQty(intLoopcount - 1) = Val(Trim(.Text))
                        Next intLoopcount
                    End With
                    strMktScheduleCheck = CheckMktSchedules()
                    If Len(Trim(strMktScheduleCheck)) > 0 Then
                        If strMktScheduleCheck = "Error" Then GoTo ErrHandler
                    End If
                    Call DeleteRecords()

                    mP_Connection.BeginTrans()

                    If AllowASNTextFileGeneration(Trim(txtCustCode.Text)) = True Then
                        mP_Connection.Execute("Delete from  MKT_ASN_INVDTL where UNIT_CODE='" & gstrUNITID & "' and Doc_No ='" & Trim(txtChallanNo.Text) & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If

                    mP_Connection.Execute(strupSaleDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    mP_Connection.Execute(strupSalechallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                    mP_Connection.Execute(STREXPDET, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    If Len(Trim(txtDispAdvNo.Text)) > 0 Then
                        mP_Connection.Execute(strDispathAdvice, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute(strupdatebarmst, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    If UpdateMktSchedules("-") = False Then
                        mP_Connection.RollbackTrans()
                        GoTo ErrHandler
                    End If
                    mP_Connection.CommitTrans()
                    txtChallanNo.Text = ""
                    Call EnableControls(False, Me, True)
                    txtLocationCode.Enabled = True
                    txtLocationCode.BackColor = System.Drawing.Color.White
                    CmdLocCodeHelp.Enabled = True
                    txtChallanNo.Enabled = True
                    txtChallanNo.BackColor = System.Drawing.Color.White
                    CmdChallanNo.Enabled = True
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
        End Select

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub

    End Sub
    Private Sub Cmditems_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cmditems.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   Display Another Form for User To Select Item Code >From CustOrd_Dtl
        '                       And After Selecting Item Code Select Data From Sales_Dtl and Display
        '                       That Details In The Spread
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim rssalechallan As ClsResultSetDB
        Dim salechallan As String
        Dim rsSaleConf As ClsResultSetDB
        Dim strStockLocation As String

        If Len(Trim(txtDispAdvNo.Text)) > 0 Then
            MsgBox("Item(s) Can't be selected explicitly while making Invoice Against Dispatch Advice !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            Exit Sub
        End If
        With Me.SpChEntry
            .MaxRows = 1
            .Row = 1 : .Row2 = .MaxRows : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
        End With

        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                rssalechallan = New ClsResultSetDB
                salechallan = ""
                salechallan = "Select Invoice_type,SUB_CATEGORY from saleschallan_dtl where UNIT_CODE='" & gstrUNITID & "' AND doc_No = "
                salechallan = salechallan & Val(txtChallanNo.Text)
                rssalechallan.GetResult(salechallan, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rssalechallan.GetNoRows > 0 Then
                    rssalechallan.MoveFirst()
                    strInvType = rssalechallan.GetValue("Invoice_type")
                    strInvSubType = rssalechallan.GetValue("sub_category")
                End If
                rssalechallan.ResultSetClose()
                rssalechallan = Nothing
                If (strInvType = "EXP") Then
                    strStockLocation = StockLocationSalesConf(strInvType, strInvSubType, "TYPE", " datediff(dd,'" & getDateForDB(lblDateDes.Text) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(lblDateDes.Text) & "')<=0")
                    If Len(Trim(strStockLocation)) > 0 Then
                        mstrItemCode = frmMKTTRN0021_SOUTH.SelectDatafromsaleDtl(Trim(txtChallanNo.Text))
                        If Len(Trim(mstrItemCode)) = 0 Then SpChEntry.MaxRows = 0 : frmMKTTRN0021_SOUTH.Close()
                    Else
                        MsgBox("Please Define Stock Location in Sales Conf")
                        Exit Sub
                    End If
                Else
                    mstrItemCode = frmMKTTRN0021_SOUTH.SelectDatafromsaleDtl(Trim(txtChallanNo.Text))
                    If Len(Trim(mstrItemCode)) = 0 Then SpChEntry.MaxRows = 0 : frmMKTTRN0021_SOUTH.Close()
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If (Trim(CmbInvType.Text) = "EXPORT INVOICE") And (UCase(Trim(CmbInvSubType.Text)) = "EXPORTS") Then
                    If InvAgstDispAdvise() = True And mblncustomer_agstdispatchadvice = True Then
                        If Len(Trim(txtDispAdvNo.Text)) = 0 Then
                            MsgBox("Select Dispatch Advice no.", MsgBoxStyle.Information, ResolveResString(100))
                            txtDispAdvNo.Focus()
                            Exit Sub
                        End If
                    Else
                        If Len(Trim(txtRefNo.Text)) = 0 Then
                            Call ConfirmWindow(10240, ConfirmWindowButtonsEnum.BUTTON_OK)
                            txtRefNo.Focus()
                            Exit Sub
                        End If
                    End If
                End If
                If (Trim(CmbInvType.Text) = "EXPORT INVOICE") Then
                    strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION", " datediff(dd,'" & getDateForDB(dtpDateDesc.Value) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(dtpDateDesc.Value) & "')<=0  ")
                    If Len(Trim(strStockLocation)) > 0 Then
                        mstrItemCode = frmMKTTRN0021_SOUTH.SelectDataFromCustOrd_Dtl(Trim(txtCustCode.Text), Trim(txtRefNo.Text), mstrAmmNo, Trim(CmbInvSubType.Text), Trim(CmbInvType.Text), strStockLocation, , , Trim(txtConsCode.Text))
                    Else
                        MsgBox("Please Define Stock Location in Sales Conf")
                        Exit Sub
                    End If
                    If Len(Trim(mstrItemCode)) = 0 Then SpChEntry.MaxRows = 0
                End If
        End Select

        If Len(mstrItemCode) > 0 Then
            mstrItemCode = Mid(mstrItemCode, 1, Len(mstrItemCode) - 1)
            Select Case Me.CmdGrpChEnt.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                    Call DisplayDetailsInSpread() 'Procedure Call To Select Data >From Sales_Dtl
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    Call displayDeatilsfromCustOrdHdrandDtl()
            End Select

            Me.CmdGrpChEnt.Focus()
        Else
            frmMKTTRN0021_SOUTH.Close()
        End If
        Call ChangeCellTypeStaticText()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdLocCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdLocCodeHelp.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Display Help From Location Master
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strHelp As String
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                If Len(Me.txtLocationCode.Text) = 0 Then 'To check if There is No Text Then Show All Help
                    strHelp = ShowList(1, (txtLocationCode.MaxLength), "", "s.Location_Code", "l.Description", "Location_mst l,SaleConf s", "and s.unit_code = l.unit_code and s.Location_Code=l.Location_Code and datediff(dd,GETDATE(),S.fin_start_date)<=0  and datediff(dd,S.fin_end_date,GETDATE())<=0", , , , , , "S.UNIT_CODE")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtLocationCode.Text = strHelp
                    End If
                Else
                    strHelp = ShowList(1, (txtLocationCode.MaxLength), txtLocationCode.Text, "s.Location_Code", "l.Description", "Location_mst l,SaleConf s", "and s.unit_code=l.unit_code and s.Location_Code=l.Location_Code and datediff(dd,GETDATE(),S.fin_start_date)<=0  and datediff(dd,S.fin_end_date,GETDATE())<=0", , , , , , "S.UNIT_CODE")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtLocationCode.Text = strHelp
                    End If
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
        End Select
        Call SelectDescriptionForField("Description", "Location_Code", "Location_Mst", lblLocCodeDes, (txtLocationCode.Text))
        txtLocationCode.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdRefNoHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdRefNoHelp.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Display Details Of Customer Order
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If Len(txtCustCode.Text) = 0 Then
            Call ConfirmWindow(10416, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            txtCustCode.Focus()
            Exit Sub
        End If
        Dim strRefAmm As String
        Dim intPos As Short
        strRefAmm = frmMKTTRN0020_SOUTH.SelectDataFromCustOrd_Dtl((txtCustCode.Text), (CmbInvType.Text), Trim(txtConsCode.Text))
        If Len(strRefAmm) > 0 Then
            intPos = InStr(1, Trim(strRefAmm), ",", CompareMethod.Text)
            mstrRefNo = Mid(Trim(strRefAmm), 2, intPos - 3)
            mstrAmmNo = Mid(strRefAmm, intPos + 2, ((Len(Trim(strRefAmm))) - intPos) - 2)
            txtRefNo.Text = Trim(mstrRefNo)
            If CmbInvType.Text.ToUpper = "EXPORT INVOICE" Then
                mstrexportsotype = Find_Value("SELECT EXPORTSOTYPE FROM CUST_ORD_HDR WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" & txtCustCode.Text & "' AND cust_ref='" & mstrRefNo & "' and amendment_no='" & mstrAmmNo & "'")
                lblexportsodetails.Text = mstrexportsotype
            Else
                lblexportsodetails.Text = ""
            End If
            If txtCarrServices.Enabled = True Then txtCarrServices.Focus()
        Else
            If txtCarrServices.Enabled = True Then txtCarrServices.Focus()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdSaleTaxType_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdSaleTaxType.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Display Help From SaleTax Master
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strHelp As String
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Len(Me.txtSaleTaxType.Text) = 0 Then 'To check if There is No Text Then Show All Help
                    strHelp = ShowList(1, (txtSaleTaxType.MaxLength), "", "SaleTax_Code", "SaleTax_Type", "SaleTax_Mst")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtSaleTaxType.Text = LTrim(RTrim(strHelp))
                    End If
                Else
                    strHelp = ShowList(1, (txtSaleTaxType.MaxLength), txtSaleTaxType.Text, "SaleTax_Code", "SaleTax_Type", "SaleTax_Mst")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtSaleTaxType.Text = LTrim(RTrim(strHelp))
                    End If
                End If
        End Select
        txtSaleTaxType.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmexport_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmexport.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Disply Export details Form
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim arrintExpDetails() As String
        Dim strMode As String
        frmMKTTRN0022_SOUTH.SetDocumentDate = CStr(dtpDateDesc.Value)
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            frmMKTTRN0022_SOUTH.SetCurrencyID = GetCurrencyINSO("ADD")
            '09/12/15
            mstrCreditTermId = Find_Value("select term_payment from cust_ord_hdr where UNIT_CODE='" & gstrUNITID & "' AND Account_code='" & Trim(txtCustCode.Text) & "' AND cust_ref='" & txtRefNo.Text & "'")
            frmMKTTRN0022_SOUTH.SetCreditTermID = mstrCreditTermId
        Else
            frmMKTTRN0022_SOUTH.SetCurrencyID = GetCurrencyINSO("EDIT")
            frmMKTTRN0022_SOUTH.SetCreditTermID = mstrCreditTermId
        End If
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                strMode = "MODE_VIEW"

                If Len(Trim(strExpDetails)) Then
                    strExpDetails = frmMKTTRN0022_SOUTH.ShowValuestoString(strExpDetails, strMode)
                Else 'STREXPdETAIL IS lOCAL VARIABLE THEN TO ASSIGN VALUES OF STRVALUES
                    strExpDetails = strValues
                    strValues = ""
                    strExpDetails = frmMKTTRN0022_SOUTH.ShowValuestoString(strExpDetails, strMode)
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strMode = "MODE_EDIT"

                frmMKTTRN0022_SOUTH.mstrInvSubType = UCase(CmbInvSubType.Text)

                If Len(Trim(strExpDetails)) Then
                    strExpDetails = frmMKTTRN0022_SOUTH.ShowValuestoString(strExpDetails, strMode)
                Else 'STREXPdETAIL IS lOCAL VARIABLE THEN TO ASSIGN VALUES OF STRVALUES
                    strExpDetails = strValues
                    strValues = ""
                    strExpDetails = frmMKTTRN0022_SOUTH.ShowValuestoString(strExpDetails, strMode)
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                strMode = "MODE_ADD"
                If (Trim(CmbInvType.Text) = "EXPORT INVOICE") Then
                    'Changed for Issue ID 21054 Starts
                    'If InvAgstDispAdvise() = False Then
                    '    If Len(Trim(txtRefNo.Text)) = 0 Then
                    '        Call ConfirmWindow(10240, ConfirmWindowButtonsEnum.BUTTON_OK)
                    '        txtRefNo.Focus()
                    '        Exit Sub
                    '    End If
                    'Else
                    '    If Len(Trim(txtDispAdvNo.Text)) = 0 Then
                    '        MsgBox("Select Dispatch Advice no.", vbInformation, ResolveResString(100))
                    '        txtDispAdvNo.Focus()
                    '        Exit Sub
                    '    End If
                    'End If
                    If InvAgstDispAdvise() = True And mblncustomer_agstdispatchadvice = True Then
                        If Len(Trim(txtDispAdvNo.Text)) = 0 Then
                            MsgBox("Select Dispatch Advice no.", vbInformation, ResolveResString(100))
                            txtDispAdvNo.Focus()
                            Exit Sub
                        End If
                    Else
                        If Len(Trim(txtRefNo.Text)) = 0 Then
                            Call ConfirmWindow(10240, ConfirmWindowButtonsEnum.BUTTON_OK)
                            txtRefNo.Focus()
                            Exit Sub
                        End If
                    End If

                    frmMKTTRN0022_SOUTH.mstrInvSubType = UCase(CmbInvSubType.Text)
                    If Len(Trim(strValues)) = 0 Then
                        strExpDetails = frmMKTTRN0022_SOUTH.ShowValuestoString(strExpDetails, strMode)
                    Else
                        strExpDetails = frmMKTTRN0022_SOUTH.ShowValuestoString(strValues, strMode)
                    End If
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlFormHeader1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        On Error GoTo ErrHandler
        Call ShowHelp("HLPMKTTRN0005.HTM")
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub dtpDateDesc_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComCtl2.DDTPickerEvents_KeyDownEvent)
        On Error GoTo Err_Handler
        If eventArgs.keyCode = System.Windows.Forms.Keys.Return And eventArgs.shift = 0 Then
            CmbInvType.Focus()
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmEXPTRN0010_SOUTH_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Disply empoer .hlp help
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo Err_Handler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlFormHeader1_ClickEvent(ctlFormHeader1, New System.EventArgs())
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmEXPTRN0010_SOUTH_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                If Me.CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call Me.CmdGrpChEnt.Revert()
                        Call EnableControls(False, Me, True)
                        CmbInvType.Enabled = True : CmbInvSubType.Enabled = True
                        txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : lblLocCodeDes.Text = ""
                        txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        CmdLocCodeHelp.Enabled = True : CmdChallanNo.Enabled = True : Me.SpChEntry.Enabled = False
                        Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                        Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                        Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                        CmbInvType.SelectedIndex = 0 : CmbInvSubType.SelectedIndex = 0
                        gblnCancelUnload = False
                        gblnFormAddEdit = False
                        With Me.SpChEntry
                            .MaxRows = 1 : .set_RowHeight(1, 300)
                            .Row = 1 : .Row2 = 1 : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
                        End With
                        lblDateDes.Text = VB6.Format(GetServerDate(), gstrDateFormat)
                        dtpDateDesc.Visible = False
                        txtLocationCode.Focus()
                    Else
                        Me.ActiveControl.Focus()
                    End If
                End If
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub frmEXPTRN0010_SOUTH_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To initilise Values on Form Load
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************

        On Error GoTo ErrHandler
        Dim blnyes As Boolean
        strValues = ""
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FillLabelFromResFile(Me) 'Fill Labels >From Resource File
        Call FitToClient(Me, FraChEnt, ctlFormHeader1, CmdGrpChEnt)
        CmdLocCodeHelp.Image = My.Resources.ico111.ToBitmap
        CmdChallanNo.Image = My.Resources.ico111.ToBitmap
        CmdCustCodeHelp.Image = My.Resources.ico111.ToBitmap
        CmdSaleTaxType.Image = My.Resources.ico111.ToBitmap
        CmdRefNoHelp.Image = My.Resources.ico111.ToBitmap
        Call EnableControls(False, Me, True)
        dtpDateDesc.Format = DateTimePickerFormat.Custom
        dtpDateDesc.CustomFormat = gstrDateFormat
        lblDateDes.Text = VB6.Format(GetServerDate(), gstrDateFormat)
        With dtpDateDesc
            .Value = ConvertToDate(lblDateDes.Text)
            .Visible = False
        End With
        Call AddTransPortTypeToCombo()
        txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        CmdLocCodeHelp.Enabled = True : CmdChallanNo.Enabled = True : Me.SpChEntry.Enabled = False
        Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
        Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
        Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        mblnInvocieforMTL = Find_Value("select isnull(InvoiceForMTLSharjah,0)as InvoiceForMTLSharjah from sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'")
        mblnInvoicelike_MTLsharjah = Find_Value("select isnull(Invoicelike_MTLsharjah,0)as Invoicelike_MTLsharjah from sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'")

        With Me.SpChEntry
            .Row = 0 : .Col = 1 : .Text = "Item Code" : .set_ColWidth(1, 2000)
            .Row = 0 : .Col = 2 : .Text = "Drawing No." : .set_ColWidth(2, 2000)
            .Row = 0 : .Col = 3 : .Text = "Rate" : .set_ColWidth(3, 1100)
            .Row = 0 : .Col = 4 : .Text = "Cust Material"
            .Row = 0 : .Col = 5 : .Text = "Quantity"
            .Row = 0 : .Col = 6 : .Text = "Packing"
            .Row = 0 : .Col = 7 : .Text = "Others"
            .Row = 0 : .Col = 8 : .Text = "From Box"
            .Row = 0 : .Col = 9 : .Text = "To Box"
            If GetPlantName() = "HILEX" Then
                .Row = 0 : .Col = 10 : .Text = "Bin Qty" : .set_ColWidth(10, 1000)
                .MaxCols = 10
                'ElseIf mblnInvocieforMTL = True Or (mblnInvoicelike_MTLsharjah = True And mblncustomer_like_MTLsharjah = True) Then
            ElseIf mblnInvocieforMTL = True Or (InvAgstDispAdvise() = True And mblncustomer_agstdispatchadvice = True) Then
                .MaxCols = 12
                .Row = 0 : .Col = 10 : .Text = "Bin Qty" : .set_ColWidth(10, 1000)
                .Row = 0 : .Col = 11 : .Text = "Reference No" : .set_ColWidth(11, 1000)
                .Row = 0 : .Col = 12 : .Text = "Amendment No." : .set_ColWidth(12, 1000)
            Else
                .MaxCols = 9
            End If
        End With
        Call SelectInvoiceTypeFromSaleConf()
        mInvAgstDispAdv = InvAgstDispAdvise()
        Call addRowAtEnterKeyPress(1)


        'strsql = "Select Bnk_BankId,Bnk_accNo from Gen_bankMaster where unit_code='" & gstrUNITID & "' and EXPORT_INVOICE_Default_Bank = 1"
        'If IsRecordExists(strsql) Then
        '    DT = SqlConnectionclass.GetDataTable(strsql)
        '    If DT.Rows.Count = 1 Then
        '        txtBankAc.Text = DT.Rows(0).Item("Bnk_BankId").ToString.Trim
        '        lblAcCodeDes.Text = DT.Rows(0).Item("Bnk_accNo").ToString.Trim
        '    Else
        '        txtBankAc.Text = ""
        '        lblAcCodeDes.Text = ""
        '    End If

        'End If


        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmEXPTRN0010_SOUTH_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Set some Values on Activation of form
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If CmbInvType.Items.Count <= 0 Then
            MsgBox("No Data Defined for this financial Year For Export Invoice in Sales Conf.", MsgBoxStyle.Information, "empower")
        End If
        mdifrmMain.CheckFormName = mintIndex
        If txtLocationCode.Enabled = True Then txtLocationCode.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub

    Private Sub frmEXPTRN0010_SOUTH_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmEXPTRN0010_SOUTH_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        On Error GoTo errHandler
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
            If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then 'If not View Mode
                enmValue = ConfirmWindow(10055, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) 'Confirm before unloading the FORM
                If enmValue <> eMPowerFunctions.ConfirmWindowReturnEnum.VAL_CANCEL Then 'If  'YES' or 'NO'
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then 'If YES
                        Call CmdGrpChEnt_ButtonClick(CmdGrpChEnt, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE))
                        eventArgs.Cancel = True
                    ElseIf enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Then  'If NO than Unload the Form
                        gblnCancelUnload = False 'Variable used in MDI Form before unloading
                        gblnFormAddEdit = False
                        Me.CmdGrpChEnt.Focus()
                    Else
                        gblnCancelUnload = True : gblnFormAddEdit = True ' If Cancel than Focus will be set on in the first field.
                        Me.CmdGrpChEnt.Focus()
                    End If
                Else
                    If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then ' if Mode is add than focus will be fixed on Plant Code
                        Me.txtLocationCode.Focus()
                    Else
                        Me.txtLocationCode.Focus() ' Else Focus will be set on Plant Name.
                    End If
                    gblnCancelUnload = True
                    gblnFormAddEdit = True
                End If
            Else
                gblnCancelUnload = False
            End If
        End If
        If gblnCancelUnload = True Then eventArgs.Cancel = True 'Do not unload FORM, if the value of gblncancelUnload is False
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmEXPTRN0010_SOUTH_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        frmMKTTRN0020_SOUTH = Nothing 'Assign form to nothing
        frmMKTTRN0021_SOUTH = Nothing 'Assign form to nothing
        frmMKTTRN0022_SOUTH = Nothing 'Assign form to nothing
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Me.Dispose()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub addRowAtEnterKeyPress(ByRef pintRows As Short)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   Add Row At Enter Key Press Of Last Column Of Spread
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim intRowHeight As Short
        With Me.SpChEntry
            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            .ColsFrozen = 1 : .ColsFrozen = 2
            For intRowHeight = 1 To pintRows
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .set_RowHeight(.Row, 300)
                .Col = 4 ''Cust Matt
                .ColHidden = True
                .Col = 6 ''Packing
                .ColHidden = True
                .Col = 7 ''Others
                .ColHidden = True
            Next intRowHeight
            If .MaxRows > 4 Then .ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsBoth
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Function SelectInvoiceTypeFromSaleConf() As Object
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   Select Invoice Type,Invoice SubTypeDescription From SaleConf
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strSaleConfSql As String
        Dim rsSaleConf As ClsResultSetDB
        Dim intRecCount As Short
        Dim intLoopCounter As Short
        strSaleConfSql = "Select Distinct(Description) from SaleConf where UNIT_CODE='" & gstrUNITID & "' AND Invoice_Type in('EXP') and datediff(dd,GETDATE(),fin_start_date)<=0  and datediff(dd,fin_end_date,GETDATE())<=0"
        rsSaleConf = New ClsResultSetDB
        rsSaleConf.GetResult(strSaleConfSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSaleConf.GetNoRows > 0 Then
            intRecCount = rsSaleConf.GetNoRows
            rsSaleConf.MoveFirst()
            For intLoopCounter = 0 To intRecCount - 1
                VB6.SetItemString(CmbInvType, intLoopCounter, rsSaleConf.GetValue("Description"))
                rsSaleConf.MoveNext()
            Next intLoopCounter
        End If
        rsSaleConf.ResultSetClose()

        rsSaleConf = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Private Sub SelectInvoiceSubTypeFromSaleConf(ByRef pstrInvType As String)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   Select Invoice SubTypeDescription From SaleConf Acc. to Inv. Type
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strSaleConfSql As String
        Dim rsSaleConf As ClsResultSetDB
        Dim intRecCount As Short
        Dim intLoopCounter As Short
        strSaleConfSql = "Select Distinct(Sub_Type_Description) from SaleConf where UNIT_CODE='" & gstrUNITID & "' AND Description='" & Trim(pstrInvType) & "'and datediff(dd,GETDATE(),fin_start_date)<=0  and datediff(dd,fin_end_date,GETDATE())<=0"
        rsSaleConf = New ClsResultSetDB
        rsSaleConf.GetResult(strSaleConfSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSaleConf.GetNoRows > 0 Then
            intRecCount = rsSaleConf.GetNoRows
            rsSaleConf.MoveFirst()
            CmbInvSubType.Items.Clear()
            For intLoopCounter = 0 To intRecCount - 1
                VB6.SetItemString(CmbInvSubType, intLoopCounter, rsSaleConf.GetValue("Sub_Type_Description"))
                rsSaleConf.MoveNext()
            Next intLoopCounter
            CmbInvSubType.SelectedIndex = 0
        End If
        rsSaleConf.ResultSetClose()

        rsSaleConf = Nothing
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SelectDescriptionForField(ByRef pstrFieldName1 As String, ByRef pstrFieldName2 As String, ByRef pstrTableName As String, ByRef pContrName As System.Windows.Forms.Control, ByRef pstrControlText As String)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   pstrFieldName1 - Field Name1,pstrFieldName2 - Field Name2,pstrTableName - Table Name
        '                       pContName - Name Of The Control where Caption Is To Be Set
        '                       pstrControlText - Field Text
        'Return Value       :   NA
        'Function           :   To Select The Field Description In The Description Labels
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strDesSql As String 'Declared to make Select Query
        Dim rsDescription As ClsResultSetDB
        strDesSql = "Select " & Trim(pstrFieldName1) & " from " & Trim(pstrTableName) & " where " & Trim(pstrFieldName2) & "='" & Trim(pstrControlText) & "' and unit_code='" & gstrUNITID & "'"
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
    Private Sub txtAnnex_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAnnex.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        txtCarrServices.Focus()
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtBankAc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBankAc.KeyDown
        If e.KeyCode = Keys.F1 Then
            Me.cmdAcCode.PerformClick()
        End If
    End Sub

    Private Sub txtBankAc_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBankAc.TextChanged
        lblAcCodeDes.Text = ""
    End Sub

    Private Sub txtBankAc_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBankAc.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Creation Date      :   15/05/2001
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Call txtBankAc_Validating(txtBankAc, New System.ComponentModel.CancelEventArgs(False))
                chkServiceInvFormat.Focus()
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
    Private Sub txtBankAc_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtBankAc.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim clsBankMster As ClsResultSetDB
        On Error GoTo ErrHandler
        If Len(txtBankAc.Text) > 0 Then
            clsBankMster = New ClsResultSetDB
            clsBankMster.GetResult("Select Bnk_Bankid,Bnk_accNo from Gen_bankMaster where UNIT_CODE='" & gstrUNITID & "' AND bnk_Bankid ='" & Trim(txtBankAc.Text) & "'")
            If clsBankMster.GetNoRows > 0 Then
                clsBankMster.MoveFirst()
                lblAcCodeDes.Text = clsBankMster.GetValue("Bnk_accNo")
                If Cmditems.Enabled Then Cmditems.Focus()
            Else
                MsgBox("Invalid Bank Account Code.", MsgBoxStyle.Information, ResolveResString(100))
                Cancel = True
                txtBankAc.Text = ""
                If txtBankAc.Enabled Then txtBankAc.Focus()
            End If
            clsBankMster.ResultSetClose()
            clsBankMster = Nothing
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred

EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtCarrServices_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCarrServices.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        CmbTransType.Focus()
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
    Private Sub txtChallanNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtChallanNo.TextChanged
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Refresh Value of other controls when Challan No changes
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If Len(Trim(txtChallanNo.Text)) = 0 Then
            Select Case Me.CmdGrpChEnt.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    Call RefreshForm("CHALLAN")
            End Select
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtChallanNo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtChallanNo.Enter
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Show Data Selected.
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Me.txtChallanNo.SelectionStart = 0
        Me.txtChallanNo.SelectionLength = Len(Me.txtChallanNo.Text)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtChallanNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtChallanNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                        If Len(txtChallanNo.Text) > 0 Then
                            Call txtChallanNo_Validating(txtChallanNo, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            Me.CmdGrpChEnt.Focus()
                        End If
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        txtCustCode.Focus()
                End Select
        End Select
        If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then
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
    Private Sub txtChallanNo_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtChallanNo.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   If F1 Key Press Then Display Help From SalesChallan_Dtl
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdChallanNo.Enabled Then Call CmdChallanNo_Click(CmdChallanNo, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub txtChallanNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtChallanNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   Check Validity Of Challan No. In SalesChallan_Dtl
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                If Len(txtChallanNo.Text) > 0 Then
                    If CheckExistanceOfFieldData((txtChallanNo.Text), "Doc_No", "SalesChallan_Dtl", "UNIT_CODE='" & gstrUNITID & "'") Then
                        If Len(txtLocationCode.Text) > 0 Then
                            'Samiksha shipadress code changes in GetDatInViewMode
                            If GetDataInViewMode() Then 'if record found
                                Call FillDataAgstDispAdv()
                                Cmditems.Enabled = True
                                cmexport.Enabled = True
                                Cmditems.Focus()
                            Else 'if no record found then display message
                                Call ConfirmWindow(10414, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                                Cmditems.Enabled = False
                                txtLocationCode.Focus()
                            End If
                        Else 'if location code field is blank
                            Call ConfirmWindow(10239, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            txtLocationCode.Focus()
                        End If
                    Else 'If Doc_No Is Invalid
                        Call ConfirmWindow(10404, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        txtChallanNo.Text = ""
                    End If


                End If
        End Select
        If Val(txtChallanNo.Text) > 99000000 Then
            Cmditems.Enabled = True
        Else
            CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
            CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
        End If


        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub


    Private Sub txtConsCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConsCode.TextChanged
        '*******************************************************************************
        'Author             :   Ashutosh Verma
        'Argument(s)if any  :   Issue Id: 19661
        'Return Value       :   NA
        'Function           :   To Refresh the Values of Other Controls when Customer code changes.
        'Comments           :   NA
        'Creation Date      :   19 Mar 2007
        '*******************************************************************************
        On Error GoTo ErrHandler
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            lblConsCodeDes.Text = ""
        End If
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            If txtConsCode.Enabled = True Then txtConsCode.Focus()
        ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            CmdGrpChEnt.Focus()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub txtConsCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtConsCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Ashutosh Verma
        'Argument(s)if any  :   Issue Id: 19661
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   19 Mar 2007
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Len(txtConsCode.Text) > 0 Then
                            Call txtConsCode_Validating(txtConsCode, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            If txtRefNo.Enabled Then txtRefNo.Focus()
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
    Private Sub txtConsCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtConsCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*******************************************************************************
        'Author             :   Ashutosh Verma
        'Argument(s)if any  :   Issue Id: 19661
        'Return Value       :   NA
        'Function           :   If F1 Key Press Then Display Help From Customer Master/Vendor Master
        'Comments           :   NA
        'Creation Date      :   19 Mar 2007
        '*******************************************************************************
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdCustCodeHelp.Enabled Then Call CmdCustCodeHelp_Click(CmdCustCodeHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtConsCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtConsCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*******************************************************************************
        'Author             :   Ashutosh Verma
        'Argument(s)if any  :   Issue Id: 19661
        'Return Value       :   NA
        'Function           :   To Validate Cons Code Entered by User
        'Comments           :   NA
        'Creation Date      :   19 Mar 2007
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If Len(txtConsCode.Text) > 0 Then
                    If Trim(mstrInvoiceType) = "EXP" Then
                        If CheckExistanceOfFieldData((txtConsCode.Text), "Customer_Code", "Customer_Mst", "UNIT_CODE='" & gstrUNITID & "'") Then
                            Call SelectDescriptionForField("Cust_Name", "Customer_Code", "Customer_Mst", lblConsCodeDes, (txtConsCode.Text))
                            If (CmbInvType.Text = "EXPORT INVOICE") Then
                                txtDispAdvNo.Focus()
                            Else
                                txtCarrServices.Focus()
                            End If
                        Else
                            Cancel = True
                            MsgBox("Invalid Consignee Code.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                            txtConsCode.Text = ""
                            txtConsCode.Focus()
                        End If
                    End If
                End If
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtCustCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustCode.TextChanged
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Refresh the Values of Other Controls when Customer code changes.
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            lblCustCodeDes.Text = ""
            txtRefNo.Text = ""
            txtDispAdvNo.Text = ""
            SpChEntry.MaxRows = 0
            mstrItemCode = ""
        End If
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            txtCustCode.Focus()
        ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            CmdGrpChEnt.Focus()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtCustCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Len(txtCustCode.Text) > 0 Then
                            Call txtCustCode_Validating(txtCustCode, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            If txtRefNo.Enabled Then txtRefNo.Focus()
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
    Private Sub txtCustCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   If F1 Key Press Then Display Help From Customer Master/Vendor Master
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdCustCodeHelp.Enabled Then Call CmdCustCodeHelp_Click(CmdCustCodeHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtCustCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        'Function           :   To Validate Customer Code Entered by User
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        Dim strSql As String = ""
        Dim DT As DataTable

        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If Len(txtCustCode.Text) > 0 Then
                    If Trim(mstrInvoiceType) = "EXP" Then

                        If CheckExistanceOfFieldData((txtCustCode.Text), "Customer_Code", "Customer_Mst", "UNIT_CODE='" & gstrUNITID & "'") Then
                            mblncustomer_agstdispatchadvice = CBool(Find_Value("SELECT ENABLE_AGSTDISPADVICE FROM customer_mst WHERE UNIT_CODE='" & gstrUNITID & "' and customer_code='" & txtCustCode.Text & "'"))
                            Call SelectDescriptionForField("Cust_Name", "Customer_Code", "Customer_Mst", lblCustCodeDes, (txtCustCode.Text))
                            txtConsCode.Text = txtCustCode.Text
                            Call SelectDescriptionForField("Cust_Name", "Customer_Code", "Customer_Mst", lblConsCodeDes, (txtCustCode.Text))
                            Call SetControlsforASNDetails(Trim(txtCustCode.Text))
                            If (CmbInvType.Text = "EXPORT INVOICE") Then
                                txtDispAdvNo.Focus()
                            Else
                                txtCarrServices.Focus()
                            End If
                            txtBankAc.Text = ""
                            lblAcCodeDes.Text = ""
                            strSql = "SELECT ISNULL(EXPORT_ENTRY_BANK_CODE, '') EXPORT_ENTRY_BANK_CODE, ISNULL(BNK_ACCNO, '') BNK_ACCNO FROM CUSTOMER_MST A " & _
                                    " INNER JOIN GEN_BANKMASTER B " & _
                                    " ON A.UNIT_CODE = B.UNIT_CODE AND A.EXPORT_ENTRY_BANK_CODE = B.BNK_BANKID " & _
                                    " WHERE A.UNIT_CODE = '" & gstrUNITID & "' AND A.Customer_Code = '" & txtCustCode.Text.Trim & "'"

                            If IsRecordExists(strSql) Then
                                DT = SqlConnectionclass.GetDataTable(strSql)
                                If DT.Rows.Count >= 0 Then
                                    txtBankAc.Text = DT.Rows(0).Item("EXPORT_ENTRY_BANK_CODE").ToString.Trim
                                    lblAcCodeDes.Text = DT.Rows(0).Item("BNK_ACCNO").ToString.Trim
                                Else
                                    txtBankAc.Text = ""
                                    lblAcCodeDes.Text = ""
                                End If

                            End If


                        Else
                            Cancel = True
                            Call ConfirmWindow(10417, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            txtCustCode.Text = ""
                            txtCustCode.Focus()
                        End If
                    End If
                End If
        End Select
        If InvAgstDispAdvise() = True And mblncustomer_agstdispatchadvice = True Then
            txtRefNo.Enabled = False : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            CmdRefNoHelp.Enabled = False
            txtDispAdvNo.Enabled = True : txtDispAdvNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            cmdDispAdvNo.Enabled = True
        Else
            txtRefNo.Enabled = True : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            CmdRefNoHelp.Enabled = True
            txtDispAdvNo.Enabled = False : txtDispAdvNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            cmdDispAdvNo.Enabled = False
        End If
        With SpChEntry
            If mblnInvocieforMTL = True Or (InvAgstDispAdvise() = True And mblncustomer_agstdispatchadvice = True) Then
                .MaxCols = 12
                .Row = 0 : .Col = 10 : .Text = "Bin Qty" : .set_ColWidth(10, 1000)
                .Row = 0 : .Col = 11 : .Text = "Reference No" : .set_ColWidth(11, 1000)
                .Row = 0 : .Col = 12 : .Text = "Amendment No." : .set_ColWidth(12, 1000)
            Else
                .MaxCols = 9
            End If

        End With
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtDispAdvNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDispAdvNo.TextChanged
        On Error GoTo ErrHandler
        If Len(Trim(txtDispAdvNo.Text)) = 0 Then
            txtRefNo.Text = ""
            mstrRefNo = ""
            mstrAmmNo = ""
            SpChEntry.MaxRows = 0
            If txtDispAdvNo.Enabled Then txtDispAdvNo.Focus()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtDispAdvNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtDispAdvNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then cmdDispAdvNo.PerformClick()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtDispAdvNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDispAdvNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                If Len(Trim(txtDispAdvNo.Text)) > 0 Then
                    Call txtDispAdvNo_Validating(txtDispAdvNo, New System.ComponentModel.CancelEventArgs(False))
                Else
                    If txtCarrServices.Enabled Then txtCarrServices.Focus()
                End If
            Case Asc("0") To Asc("9")
            Case System.Windows.Forms.Keys.Back
            Case System.Windows.Forms.Keys.Delete
            Case Else
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtDispAdvNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtDispAdvNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '-------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Function      : To Validate Entered Despatch Advise No.
        ' Datetime      : 12-Feb-2007
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim Rs As ClsResultSetDB
        If Len(Trim(txtDispAdvNo.Text)) > 0 Then
            Rs = New ClsResultSetDB
            strsql = "Select Top 1 1 from BAR_DISPATCHADVICE_HDR where UNIT_CODE='" & gstrUNITID & "' AND CustomerCode = '" & Trim(txtCustCode.Text) & "' AND Consignee_Code='" & Trim(txtConsCode.Text) & "' AND DocNo='" & Trim(txtDispAdvNo.Text) & "' and isnull(InvoiceNo,0) = 0 And Status = 0"
            If Rs.GetResult(strsql) = False Then GoTo ErrHandler
            If Rs.GetNoRows = 0 Then
                MsgBox("Entered Dispatch Advice is Invalid", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                txtDispAdvNo.Text = ""
                Cancel = True
                GoTo EventExitSub
            ElseIf CheckScannedPallete(Trim(txtDispAdvNo.Text)) = False Then
                MsgBox("All the Palletes are not scanned against this Dispatch Advice No.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                txtDispAdvNo.Text = ""
                Cancel = True
                GoTo EventExitSub
            ElseIf Rs.GetNoRows > 0 Then
                Call FillDataAgstDispAdv()
            End If
            txtCarrServices.Focus()
            Rs.ResultSetClose()
            Rs = Nothing
        End If
        GoTo EventExitSub
ErrHandler:

        Rs = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtLocationCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLocationCode.TextChanged
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Set The Values of Related control on Change of Location Code
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If Len(Trim(txtLocationCode.Text)) = 0 Then
            Select Case Me.CmdGrpChEnt.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    Call RefreshForm("LOCATION")
            End Select
        End If
        txtCustCode.Text = ""
        txtConsCode.Text = ""
        lblConsCodeDes.Text = ""
        lblCustCodeDes.Text = ""
        txtRefNo.Text = ""
        SpChEntry.MaxRows = 0
        mstrItemCode = ""
        lblLocCodeDes.Text = ""
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLocationCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLocationCode.Enter
        On Error GoTo ErrHandler
        Me.txtLocationCode.SelectionStart = 0
        Me.txtLocationCode.SelectionLength = Len(Me.txtLocationCode.Text)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLocationCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLocationCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                        If Len(txtLocationCode.Text) > 0 Then
                            Call txtLocationCode_Validating(txtLocationCode, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            Me.CmdGrpChEnt.Focus()
                        End If
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Len(txtLocationCode.Text) > 0 Then
                            Call txtLocationCode_Validating(txtLocationCode, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            txtCustCode.Focus()
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
    Private Sub txtLocationCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtLocationCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdLocCodeHelp.Enabled Then Call CmdLocCodeHelp_Click(CmdLocCodeHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SelectInvTypeSubTypeFromSaleConf(ByRef pstrInvType As String, ByRef pstrInvSubtype As String)
        On Error GoTo ErrHandler
        Dim strSaleConfSql As String
        Dim rsSaleConf As ClsResultSetDB
        strSaleConfSql = "Select Invoice_Type,Sub_Type from SaleConf where UNIT_CODE='" & gstrUNITID & "' AND Description='" & Trim(pstrInvType) & "'"
        strSaleConfSql = strSaleConfSql & " and Sub_Type_Description='" & Trim(pstrInvSubtype) & "' and datediff(dd,GETDATE(),fin_start_date)<=0  and datediff(dd,fin_end_date,GETDATE())<=0"
        rsSaleConf = New ClsResultSetDB
        rsSaleConf.GetResult(strSaleConfSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSaleConf.GetNoRows > 0 Then
            mstrInvoiceType = rsSaleConf.GetValue("Invoice_Type")
            mstrInvoiceSubType = rsSaleConf.GetValue("Sub_Type")
        End If
        rsSaleConf.ResultSetClose()
        rsSaleConf = Nothing
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLocationCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtLocationCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   Check Validity Of Location Code In The Location_Mst
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Len(txtLocationCode.Text) > 0 Then
                    If CheckExistanceOfFieldData((txtLocationCode.Text), "Location_Code", "SalesChallan_Dtl", "Unit_code = '" & gstrUNITID & "'") Then
                        If txtChallanNo.Enabled Then
                            txtChallanNo.Focus()
                        Else
                            txtCustCode.Focus()
                        End If
                    Else
                        Call ConfirmWindow(10411, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Cancel = True
                        txtLocationCode.Text = ""
                        txtLocationCode.Focus()
                    End If
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If Len(txtLocationCode.Text) > 0 Then
                    If CheckExistanceOfFieldData((txtLocationCode.Text), "Location_Code", "SaleConf", "UNIT_CODE='" & gstrUNITID & "' AND datediff(dd,'" & getDateForDB(dtpDateDesc.Value) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(dtpDateDesc.Value) & "')<=0") Then
                        If txtChallanNo.Enabled Then
                            txtChallanNo.Focus()
                        Else
                            txtCustCode.Focus()
                        End If
                    Else
                        Call ConfirmWindow(10411, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Cancel = True
                        txtLocationCode.Text = ""
                        txtLocationCode.Focus()
                    End If
                End If
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtLorryNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLorryNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Ashutosh Verma
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :
        'Comments           :   NA
        'Creation Date      :   26 Mar 2007
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                If txtOTLNo.Enabled Then txtOTLNo.Focus()
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
    Private Sub txtOTLNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtOTLNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                If Cmditems.Enabled = True Then Cmditems.Focus()
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
    Private Sub txtRefNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRefNo.TextChanged
        If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD And Len(Trim(txtRefNo.Text)) = 0 Then
            SpChEntry.MaxRows = 0
            txtDispAdvNo.Text = ""
            mstrItemCode = ""
            If txtRefNo.Enabled = True Then
                txtRefNo.Focus()
            End If
        End If
    End Sub
    Private Sub txtRefNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRefNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Len(txtRefNo.Text) > 0 Then
                            Call txtRefNo_Validating(txtRefNo, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            If txtDispAdvNo.Enabled Then txtDispAdvNo.Focus()
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
    Private Sub txtRefNo_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtRefNo.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdRefNoHelp.Enabled Then Call CmdRefNoHelp_Click(CmdRefNoHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtRefNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtRefNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If Len(txtLocationCode.Text) > 0 Then
            If Len(txtRefNo.Text) > 0 Then
                If CheckExistanceOfFieldData((txtRefNo.Text), "Cust_ref", "Cust_ord_Hdr", "UNIT_CODE='" & gstrUNITID & "'") Then
                    If CmbInvType.Text = "NORMAL INVOICE" Then
                        txtCarrServices.Focus()
                    Else
                    End If
                Else
                    Call ConfirmWindow(10436, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    Cancel = True
                    txtRefNo.Text = ""
                    txtRefNo.Focus()
                End If
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtRemarks_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRemarks.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Len(Trim(txtDispAdvNo.Text)) > 0 Then
                            With SpChEntry
                                If .MaxRows > 0 Then
                                    .Row = 1
                                    .Col = 8
                                    .Focus()
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                End If
                            End With
                        Else
                            If txtBankAc.Enabled = True Then txtBankAc.Focus()
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
    Private Sub txtSaleTaxType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSaleTaxType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If Len(txtSaleTaxType.Text) > 0 Then
                            Call txtSaleTaxType_Validating(txtSaleTaxType, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            txtSalesTax.Focus()
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
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Show Help on Sales Tax Type.
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
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
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Validate Data of Sales Tax Type entered by user
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If Len(txtSaleTaxType.Text) > 0 Then
            If CheckExistanceOfFieldData((txtSaleTaxType.Text), "SaleTax_Code", "SaleTax_Mst", "UNIT_CODE='" & gstrUNITID & "'") Then
                txtSalesTax.Focus()
            Else
                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtSaleTaxType.Text = ""
                txtSaleTaxType.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
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

    Private Sub txtVehNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtVehNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                If ctlInsurance.Enabled = True Then ctlInsurance.Focus()
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



    Private Function CheckExistanceOfFieldData(ByRef pstrFieldText As String, ByRef pstrColumnName As String, ByRef pstrTableName As String, Optional ByRef pstrCondition As String = "") As Boolean
        On Error GoTo ErrHandler
        CheckExistanceOfFieldData = False
        Dim strTableSql As String 'Declared To Make Select Query
        Dim rsExistData As ClsResultSetDB
        strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "'"
        If Len(Trim(pstrCondition)) > 0 Then
            strTableSql = strTableSql & " AND " & pstrCondition
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
    Private Function GetDataInViewMode() As Boolean
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To display data in view mode from SalasChallan_Dtl,Sales_Dtl acc.to LacationCode & Challan_No.
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        GetDataInViewMode = False
        Dim strGetData As String
        Dim rsGetData As ClsResultSetDB
        Dim rsBankMaster As ClsResultSetDB
        Dim rsCustName As ClsResultSetDB
        Dim rsShipadddesc As ClsResultSetDB
        Dim strSalesChallanDtl As String
        Dim STREXPDET As String
        Dim rsCustMst As ClsResultSetDB
        Dim strCustMst As String

        'Samiksha Ship address code changes
        strSalesChallanDtl = "SELECT Transport_type,Vehicle_No,Account_code,Cust_ref,salesTax_type,"
        strSalesChallanDtl = strSalesChallanDtl & "Insurance,Invoice_Date,"
        strSalesChallanDtl = strSalesChallanDtl & "Invoice_Type,Sub_Category,Cust_Name,Carriage_Name,Frieght_Tax, "
        strSalesChallanDtl = strSalesChallanDtl & "Amendment_No,ref_doc_no,"
        strSalesChallanDtl = strSalesChallanDtl & "Currency_Code,Originstatus,ctry_destination_goods,Pre_Carriage_by,"
        strSalesChallanDtl = strSalesChallanDtl & "Receipt_PreCarriage_at,Port_of_loading,Port_of_Discharge,"
        strSalesChallanDtl = strSalesChallanDtl & "nature_of_contract,Final_destination,Mode_of_shipment,Dispatch_Mode,"
        strSalesChallanDtl = strSalesChallanDtl & "Delivery_terms,Payment_terms,Buyer_Description_of_goods,Invoice_Description_of_EPC,"
        strSalesChallanDtl = strSalesChallanDtl & "Exchange_Rate,Frieght_amount,Exchange_Date,other_ref,buyer_id,ServiceInvoiceformatExport,CustBankID,Remarks,Consignee_Code,Lorry_no,OTL_No,RefChallan,Vessel_flight_number,exportsotype,ISNULL(ShipAddress_Code,'') ShipAddress_Code"
        strSalesChallanDtl = strSalesChallanDtl & " From Saleschallan_dtl where UNIT_CODE='" & gstrUNITID & "' AND Location_Code ='"
        strSalesChallanDtl = strSalesChallanDtl & Trim(txtLocationCode.Text) & "' and Doc_No = " & Val(txtChallanNo.Text)
        rsGetData = New ClsResultSetDB
        rsGetData.GetResult(strSalesChallanDtl, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsGetData.GetNoRows > 0 Then
            GetDataInViewMode = True
            txtCustCode.Text = rsGetData.GetValue("Account_code")
            txtSaleTaxType.Text = rsGetData.GetValue("salesTax_type")
            txtConsCode.Text = rsGetData.GetValue("Consignee_Code")
            'Samiksha Ship address code changes
            txtshippingaddcode.Text = rsGetData.GetValue("ShipAddress_Code")
            If Val(Trim(CStr(Len(rsGetData.GetValue("ShipAddress_Code"))))) > 0 Then
                rsShipadddesc = New ClsResultSetDB
                rsShipadddesc.GetResult("select Distinct Shipping_Desc from Customer_Shipping_Dtl where unit_code='" & gstrUNITID & "' and InActive_Flag=0 and customer_code='" & Trim(txtCustCode.Text) & "'")
                If rsShipadddesc.GetNoRows > 0 Then
                    txtboxShipadddesc.Text = rsShipadddesc.GetValue("Shipping_Desc")
                Else
                    txtboxShipadddesc.Text = ""
                End If
                rsShipadddesc.ResultSetClose()
                rsShipadddesc = Nothing
            End If

            If Val(Trim(CStr(Len(rsGetData.GetValue("Consignee_Code"))))) > 0 Then
                    rsCustName = New ClsResultSetDB
                    rsCustName.GetResult("Select cust_name from customer_mst where UNIT_CODE='" & gstrUNITID & "' AND customer_code='" & Trim(rsGetData.GetValue("Consignee_Code")) & "'")
                    If rsCustName.GetNoRows > 0 Then
                        lblConsCodeDes.Text = rsCustName.GetValue("Cust_name")
                    Else
                        lblConsCodeDes.Text = ""
                    End If
                    rsCustName.ResultSetClose()
                    rsCustName = Nothing
                End If
                Me.txtVehNo.Text = rsGetData.GetValue("Vehicle_No")
                txtLorryNo.Text = rsGetData.GetValue("Lorry_No")
                txtOTLNo.Text = rsGetData.GetValue("OTL_No")
                txtRefChallanNo.Text = rsGetData.GetValue("RefChallan")
                txtRefNo.Text = rsGetData.GetValue("Cust_ref")
                txtCarrServices.Text = rsGetData.GetValue("Carriage_Name")
                ctlInsurance.Text = rsGetData.GetValue("Insurance")
                txtFreight.Text = rsGetData.GetValue("Frieght_tax")
                lblCustCodeDes.Text = rsGetData.GetValue("Cust_Name")
                mstrAmmendmentNo = rsGetData.GetValue("Amendment_No")
                lblDateDes.Text = VB6.Format(rsGetData.GetValue("Invoice_Date"), gstrDateFormat)
                mstrInvType = rsGetData.GetValue("Invoice_Type")
                mstrInvSubType = rsGetData.GetValue("Sub_Category")
                lblexportsodetails.Text = rsGetData.GetValue("exportsotype")
                If rsGetData.GetValue("ServiceInvoiceformatExport") = True Then
                    chkServiceInvFormat.CheckState = System.Windows.Forms.CheckState.Checked
                Else
                    chkServiceInvFormat.CheckState = System.Windows.Forms.CheckState.Unchecked
                End If
                txtBankAc.Text = Trim(rsGetData.GetValue("CustBankID"))
                rsBankMaster = New ClsResultSetDB
                rsBankMaster.GetResult("Select bnk_accno from gen_bankMaster where UNIT_CODE='" & gstrUNITID & "' AND bnk_Bankid = '" & Trim(txtBankAc.Text) & "'")
                If rsBankMaster.GetNoRows > 0 Then
                    lblAcCodeDes.Text = rsBankMaster.GetValue("bnk_accno")
                End If
                txtRemarks.Text = Trim(rsGetData.GetValue("Remarks"))
                txtvesselnumber.Text = Trim(rsGetData.GetValue("Vessel_Flight_number"))
                CmbInvType.SelectedIndex = 0
                CmbInvSubType.SelectedIndex = 0
                CmbTransType.Text = Nothing
                'issue id 10549878
                If AllowASNTextFileGeneration(Trim(txtCustCode.Text)) = True Then
                    CmbTransType.Items.Add("A - Air")
                    CmbTransType.Items.Add("S - Ocean")
                    CmbTransType.Items.Add("M - Motor")
                    CmbTransType.Items.Add("W - Inland Waterway")
                    CmbTransType.Items.Add("H - Customer Pickup")
                    CmbTransType.Items.Add("R - Rail")
                    CmbTransType.Items.Add("O - Containerized Ocean")
                    CmbTransType.Items.Add("C - Consolidation")
                    CmbTransType.Items.Add("U - UPS")
                    CmbTransType.Items.Add("E - Expedited Truck")
                End If
                'issue id 10549878
                Dim i As Integer
                For i = 0 To Me.CmbTransType.Items.Count - 1
                    If Mid(CStr(CmbTransType.Items.Item(i)), 1, 1) = Mid(rsGetData.GetValue("Transport_type"), 1, 1) Then
                        CmbTransType.SelectedIndex = i
                        Exit For
                    End If
                Next
                strExpDetails = ""
                strExpDetails = rsGetData.GetValue("Currency_Code") & "§" & rsGetData.GetValue("Originstatus") & "§" & rsGetData.GetValue("ctry_destination_goods") & "§" & rsGetData.GetValue("Pre_Carriage_by") & "§" & rsGetData.GetValue("Receipt_PreCarriage_at") & "§" & rsGetData.GetValue("Port_of_loading") & "§" & rsGetData.GetValue(" Port_of_Discharge") & "§" & rsGetData.GetValue("nature_of_contract") & "§" & rsGetData.GetValue("Final_destination") & "§" & rsGetData.GetValue("Mode_of_shipment") & "§" & rsGetData.GetValue("Dispatch_Mode") & "§" & rsGetData.GetValue("Delivery_terms") & "§" & rsGetData.GetValue("Payment_terms") & "§" & rsGetData.GetValue("Buyer_Description_of_goods") & "§" & rsGetData.GetValue("Invoice_Description_of_EPC") & "§" & rsGetData.GetValue("Exchange_Rate") & "§" & rsGetData.GetValue("Frieght_amount") & "§" & rsGetData.GetValue("Exchange_Date") & "§" & rsGetData.GetValue("other_ref") & "§" & rsGetData.GetValue("buyer_id")
                rsBankMaster.ResultSetClose()
                rsBankMaster = Nothing
            Else
                GetDataInViewMode = False
        End If

        rsGetData.ResultSetClose()
        rsGetData = Nothing
        'Samiksha commodity type changes

        STREXPDET = "select Advance_lice_No,Pallet_Length,Pallet_width,Pallet_Height,Pallet_Total,ARE_NO,Net_Weight,Gross_Weight,Export_Type,Volume_Weight,DRAWBACK_TYPE,HS_CODE,CommodityType from EXPORT_SALES_EXTRA_DETAIL where UNIT_CODE='" & gstrUNITID & "' AND Unt_CodeID = '" & Trim(txtLocationCode.Text) & "' and Doc_No = " & Val(txtChallanNo.Text) & ""
        rsGetData = New ClsResultSetDB
        rsGetData.GetResult(STREXPDET, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        'Samiksha commodity type changes
        If rsGetData.GetNoRows > 0 Then
            strExpDetails = strExpDetails & "§" & rsGetData.GetValue("Advance_lice_No") & "§" & rsGetData.GetValue("Pallet_Length") & "§" & rsGetData.GetValue("Pallet_width") & "§" & rsGetData.GetValue("Pallet_Height") & "§" & rsGetData.GetValue("Pallet_Total") & "§" & rsGetData.GetValue("ARE_NO") & "§" & rsGetData.GetValue("Net_Weight") & "§" & rsGetData.GetValue("Gross_Weight") & "§" & rsGetData.GetValue("Export_Type") & "§" & rsGetData.GetValue("Volume_Weight") & "§" & rsGetData.GetValue("DRAWBACK_TYPE") & "§" & rsGetData.GetValue("hs_code") & "§" & rsGetData.GetValue("CommodityType")
        End If
        rsGetData.ResultSetClose()
        rsGetData = Nothing
        If Len(Trim(txtCustCode.Text)) > 0 Then
            rsCustMst = New ClsResultSetDB
            strCustMst = "select cust_plantcode,arl_code from mkt_asn_invdtl where UNIT_CODE='" & gstrUNITID & "' AND doc_no=" & Val(txtChallanNo.Text)
            rsCustMst.GetResult(strCustMst, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rsCustMst.GetNoRows > 0 Then
                txtPlantCode.Text = rsCustMst.GetValue("cust_plantcode")
                txtActualReceivingLoc.Text = rsCustMst.GetValue("arl_code")
            End If
            rsCustMst.ResultSetClose()
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function DisplayDetailsInSpread() As Boolean
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To display Details From Sales_Dtl Acc To Location Code,Challan No and Drawing No
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim intLoopCounter As Short
        Dim intRecordCount As Short
        Dim strsaledtl As String
        Dim rsSalesDtl As ClsResultSetDB
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strsaledtl = ""
                strsaledtl = "Select Location_Code,Doc_No,Suffix,Item_Code,Sales_Quantity,From_Box,To_Box,Rate,Sales_Tax,Excise_Tax,Packing,Others,Cust_Mtrl,Year,Cust_Item_Code,Cust_Item_Desc,Tool_Cost,Measure_Code,Excise_type,SalesTax_type,CVD_type,SAD_type,GL_code,SL_code,Basic_Amount,Accessible_amount,CVD_Amount,SVD_amount,Excise_per,CVD_per,SVD_per,CustMtrl_Amount,ToolCost_amount,pervalue,TotalExciseAmount,SupplementaryInvoiceFlag,To_Location,Discount_type,Discount_amt,Discount_perc,From_Location,Cust_ref,Amendment_No,SRVDINO,SRVLocation,USLOC,SchTime,BinQuantity,Packing_Type,ItemPacking_Amount,Item_remark,pkg_amount,csiexcise_amount,ADD_EXCISE_TYPE,ADD_EXCISE_PER,ADD_EXCISE_AMOUNT from Sales_Dtl where UNIT_CODE='" & gstrUNITID & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "'"
                strsaledtl = strsaledtl & " and Doc_No=" & Val(txtChallanNo.Text) & " and Cust_Item_Code in(" & Trim(mstrItemCode) & ")"
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If (Trim(CmbInvType.Text) = "EXPORT INVOICE") Then
                    strsaledtl = ""
                    strsaledtl = "Select a.Item_Code,a.Cust_DrgNo,Rate,a.Cust_Mtrl,Packing,Others,a.tool_Cost,b.BinQuantity"
                    strsaledtl = strsaledtl & " from Cust_ord_dtl a Inner Join Custitem_mst b on a.unit_code = b.unit_code and a.item_code=b.item_code and"
                    strsaledtl = strsaledtl & " a.account_code=b.account_code and a.cust_drgno=b.cust_drgno where "
                    strsaledtl = strsaledtl & " a.UNIT_CODE='" & gstrUNITID & "' AND a.Account_Code ='" & txtCustCode.Text & "' and a.Cust_ref ='"
                    strsaledtl = strsaledtl & Trim(txtRefNo.Text) & "' and a.Amendment_No = '" & mstrAmmNo & "'and "
                    strsaledtl = strsaledtl & " Active_flag ='A' and a.Cust_DrgNo in(" & mstrItemCode & ")"
                Else
                    strsaledtl = ""
                    strsaledtl = "SELECT Item_Code,standard_Rate from Item_Mst where UNIT_CODE='" & gstrUNITID & "' AND "
                    strsaledtl = strsaledtl & " Status = 'A' and Hold_flag <> 1 and Item_Code in (" & mstrItemCode & ")"
                End If
        End Select
        rsSalesDtl = New ClsResultSetDB
        rsSalesDtl.GetResult(strsaledtl, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim intLoopcount As Short
        If rsSalesDtl.GetNoRows > 0 Then
            intRecordCount = rsSalesDtl.GetNoRows
            ReDim mdblPrevQty(intRecordCount - 1) ' To get value of Quantity in Arrey for updation in despatch
            ReDim mdblToolCost(intRecordCount - 1) ' To get value of Quantity i
            If intRecordCount = 1 Then
                SpChEntry.MaxRows = 0
                Call addRowAtEnterKeyPress(intRecordCount)
            Else
                Call addRowAtEnterKeyPress(intRecordCount - 1)
            End If
            rsSalesDtl.MoveFirst()

            For intLoopCounter = 1 To intRecordCount
                With Me.SpChEntry
                    Select Case Me.CmdGrpChEnt.Mode
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                            .Row = 1 : .Row2 = .MaxRows : .Col = 0 : .Col2 = .MaxCols
                            .Enabled = True : .BlockMode = True
                            If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                                .Lock = True
                            Else
                                .Lock = False
                            End If
                            .BlockMode = False
                            Call .SetText(1, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
                            Call .SetText(2, intLoopCounter, rsSalesDtl.GetValue("Cust_Item_Code"))
                            Call .SetText(3, intLoopCounter, rsSalesDtl.GetValue("Rate"))
                            Call .SetText(4, intLoopCounter, rsSalesDtl.GetValue("Cust_Mtrl"))
                            Call .SetText(5, intLoopCounter, rsSalesDtl.GetValue("Sales_Quantity"))
                            Call .GetText(5, intLoopCounter, mdblPrevQty(intLoopCounter - 1))
                            Call .SetText(6, intLoopCounter, rsSalesDtl.GetValue("Packing"))
                            Call .SetText(7, intLoopCounter, rsSalesDtl.GetValue("Others"))
                            Call .SetText(8, intLoopCounter, rsSalesDtl.GetValue("From_Box"))
                            Call .SetText(9, intLoopCounter, rsSalesDtl.GetValue("To_Box"))
                            If GetPlantName() = "HILEX" Then
                                Call .SetText(10, intLoopCounter, rsSalesDtl.GetValue("BinQuantity"))
                            End If
                            'If mblnInvocieforMTL = True Or (mblnInvoicelike_MTLsharjah = True And mblncustomer_like_MTLsharjah = True) Then
                            If mblnInvocieforMTL = True Or (InvAgstDispAdvise() = True And mblncustomer_agstdispatchadvice = True) Then
                                Call .SetText(10, intLoopCounter, rsSalesDtl.GetValue("BinQuantity"))
                                Call .SetText(11, intLoopCounter, rsSalesDtl.GetValue("Cust_ref"))
                                Call .SetText(12, intLoopCounter, rsSalesDtl.GetValue("Amendment_no"))
                            End If

                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                            .Enabled = True
                            .Row = 1 : .Row2 = .MaxRows : .Col = 0 : .Col2 = .MaxCols
                            .BlockMode = True : .Lock = False : .BlockMode = False
                            If (Trim(CmbInvType.Text) = "EXPORT INVOICE") Then
                                Call .SetText(1, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
                                Call .SetText(2, intLoopCounter, rsSalesDtl.GetValue("Cust_DrgNo"))
                                Call .SetText(3, intLoopCounter, rsSalesDtl.GetValue("Rate"))
                                Call .SetText(4, intLoopCounter, rsSalesDtl.GetValue("Cust_Mtrl"))
                                Call .SetText(6, intLoopCounter, rsSalesDtl.GetValue("Packing"))
                                Call .SetText(7, intLoopCounter, rsSalesDtl.GetValue("Others"))
                                If GetPlantName() = "HILEX" Then
                                    Call .SetText(10, intLoopCounter, rsSalesDtl.GetValue("BinQuantity"))
                                End If
                            Else
                                Call .SetText(1, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
                                Call .SetText(2, intLoopCounter, rsSalesDtl.GetValue("Item_code"))
                                Call .SetText(3, intLoopCounter, rsSalesDtl.GetValue("Standard_Rate"))
                            End If
                    End Select
                End With
                rsSalesDtl.MoveNext()
            Next intLoopCounter
        End If
        rsSalesDtl.ResultSetClose()
        rsSalesDtl = Nothing

        If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If CDbl(Trim(txtChallanNo.Text)) > 99000000 Then
                Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
            End If
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function ValidateBeforeSave(ByRef pstrMode As String) As Boolean
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Check the Blank Fields In The Form
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim lstrControls As String
        Dim lNo As Integer
        Dim lctrFocus As System.Windows.Forms.Control
        ValidateBeforeSave = True
        lNo = 1
        lstrControls = ResolveResString(10059)
        Select Case UCase(Trim(pstrMode))
            Case "ADD"
                If (Len(txtLocationCode.Text) = 0) Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Location Code."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtLocationCode
                    End If
                    ValidateBeforeSave = False
                End If
                If (Len(txtCustCode.Text) = 0) Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Customer Code."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtCustCode
                    End If
                    ValidateBeforeSave = False
                End If
                If (Len(txtConsCode.Text) = 0) Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Consignee Code."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtConsCode
                    End If
                    ValidateBeforeSave = False
                End If
                If (Len(txtBankAc.Text) = 0) Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Bank Account Code."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtBankAc
                    End If
                    ValidateBeforeSave = False
                End If
                If Not DateIsAppropriate() Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Date specified either falls Before the LAST Invoice Date or is Greater than Todays Date."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtCustCode
                    End If
                    ValidateBeforeSave = False
                End If
                If SpChEntry.MaxRows = 0 Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Select Items"
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Cmditems
                    End If
                    ValidateBeforeSave = False
                End If
                If AllowASNTextFileGeneration(Trim(txtCustCode.Text)) = True Then
                    If txtPlantCode.Text = "" Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". Plant Code."
                        lNo = lNo + 1
                        If lctrFocus Is Nothing Then
                            lctrFocus = Me.txtPlantCode
                        End If
                        ValidateBeforeSave = False
                    End If
                End If
                If (Len(txtAddExciseDuty.Text) = 0) Then
                    txtAddExciseDuty.Text = "0.00"
                End If

                If (Len(txtFreight.Text) = 0) Then
                    txtFreight.Text = "0.00"
                End If

                If (Len(Me.txtSurcharge.Text) = 0) Then
                    txtSurcharge.Text = "0.00"
                End If

                If (Len(Me.ctlSVD.Text) = 0) Then
                    ctlSVD.Text = "0.00"
                End If

                If (Len(Me.ctlInsurance.Text) = 0) Then
                    ctlInsurance.Text = "0.00"
                End If
            Case "EDIT"
                '*****

                If (Len(Me.txtAddExciseDuty.Text) = 0) Then
                    txtAddExciseDuty.Text = "0.00"
                End If

                If (Len(Me.txtFreight.Text) = 0) Then
                    txtFreight.Text = "0.00"
                End If

                If (Len(Me.txtSurcharge.Text) = 0) Then
                    txtSurcharge.Text = "0.00"
                End If

        End Select
        If Not ValidateBeforeSave Then
            MsgBox(lstrControls, MsgBoxStyle.Information, ResolveResString(10059))
            lctrFocus.Focus()
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Function
    End Function
    Private Sub ChangeCellTypeStaticText()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Change The Cell Type In Spread Control to Cell Type Static Text to
        '                       Make Cell Type UnEditable
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim intRow As Short
        Dim intcol As Short
        With Me.SpChEntry
            Select Case Me.CmdGrpChEnt.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    If (Trim(CmbInvType.Text) = "EXPORT INVOICE") Then
                        For intRow = 1 To .MaxRows
                            .Row = intRow
                            For intcol = 1 To .MaxCols
                                .Col = intcol
                                If intcol = 5 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = 8 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                                    .TypeIntegerMin = 0.0#
                                    .TypeIntegerMax = CInt("99999999")
                                    .TypeMaxEditLen = 6
                                ElseIf intcol = 9 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                                    .TypeIntegerMin = 0.0#
                                    .TypeIntegerMax = CInt("99999999")
                                    .TypeMaxEditLen = 6
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
                                If intcol = 5 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = 8 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                                    .TypeIntegerMin = 0.0#
                                    .TypeIntegerMax = CInt("99999999")
                                    .TypeMaxEditLen = 6
                                ElseIf intcol = 3 Then
                                    'issue id 10227422
                                    'If mblnInvocieforMTL = True Or (mblnInvoicelike_MTLsharjah = True And mblncustomer_like_MTLsharjah = True) Then
                                    If mblnInvocieforMTL = True Or (InvAgstDispAdvise() = True And mblncustomer_agstdispatchadvice = True) Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 6 : .TypeFloatMin = CDbl("0.000000") : .TypeFloatMax = CDbl("99999999999999.999999")
                                    Else
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 4 : .TypeFloatMin = CDbl("0.0000") : .TypeFloatMax = CDbl("99999999999999.99999")
                                    End If
                                    'issue id 10227422 done 
                                ElseIf intcol = 9 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                                    .TypeIntegerMin = 0.0#
                                    .TypeIntegerMax = CInt("99999999")
                                    .TypeMaxEditLen = 6
                                Else
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                End If
                            Next intcol
                        Next intRow
                    End If
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                    If (UCase(strInvType) = "EXP") Then
                        For intRow = 1 To .MaxRows
                            .Row = intRow
                            For intcol = 1 To .MaxCols
                                .Col = intcol
                                If intcol = 5 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = 8 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                                    .TypeIntegerMin = 0.0#
                                    .TypeIntegerMax = CInt("99999999")
                                    .TypeMaxEditLen = 6
                                ElseIf intcol = 9 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                                    .TypeIntegerMin = 0.0#
                                    .TypeIntegerMax = CInt("99999999")
                                    .TypeMaxEditLen = 6
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
                                If intcol = 5 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = 8 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                                    .TypeIntegerMin = 0.0#
                                    .TypeIntegerMax = CInt("99999999")
                                    .TypeMaxEditLen = 6
                                ElseIf intcol = 3 Then
                                    'issue id 10227422
                                    'If mblnInvocieforMTL = True Or (mblnInvoicelike_MTLsharjah = True And mblncustomer_like_MTLsharjah = True) Then
                                    If mblnInvocieforMTL = True Or (InvAgstDispAdvise() = True And mblncustomer_agstdispatchadvice = True) Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 6 : .TypeFloatMin = CDbl(Val("0.000000")) : .TypeFloatMax = CDbl(Val("99999999999999.999999"))
                                    Else
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 4 : .TypeFloatMin = CDbl(Val("0.0000")) : .TypeFloatMax = CDbl(Val("99999999999999.99999"))
                                    End If
                                    'issue id 10227422 done 
                                ElseIf intcol = 9 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                                    .TypeIntegerMin = 0.0#
                                    .TypeIntegerMax = CInt("99999999")
                                    .TypeMaxEditLen = 6
                                Else
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                End If
                            Next intcol
                        Next intRow
                    End If
            End Select
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Function QuantityCheck() As Boolean
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Check Schedule Quantity From DailyMktSchedule/MonthlyMktSchedule
        'Revision History   : Code of checking Schedules is commented and Stored Procedures are used for
        '                     checking Schedules (Daily/Monthly)
        '*******************************************************************************
        On Error GoTo ErrHandler
        QuantityCheck = False
        Dim strScheduleSql As String
        Dim rsMktSchedule As ClsResultSetDB
        Dim rsSaleConf As ClsResultSetDB
        Dim strQuantity As String
        Dim intRwCount As Short 'To Count No. Of Rows
        Dim intLoopcount As Short
        Dim varItemQty As Object 'To Get Quantity Acc. To Drawing No and Item Code
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim strItembal As String
        Dim PresQty As Object
        Dim intcol As Short
        Dim intFromBox As Long
        Dim strScheduleCheck As String
        rsMktSchedule = New ClsResultSetDB
        For intRwCount = 1 To SpChEntry.MaxRows
            For intcol = 1 To SpChEntry.MaxCols
                SpChEntry.Col = intcol
                If (SpChEntry.Col = 5) Or (SpChEntry.Col = 3) Or (SpChEntry.Col = 9) Or (SpChEntry.Col = 8) Then
                    SpChEntry.Row = intRwCount
                    If (Val(Trim(SpChEntry.Text)) = 0) Then
                        QuantityCheck = True
                        Call ConfirmWindow(10419, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        SpChEntry.Row = intRwCount : SpChEntry.Col = intcol : SpChEntry.Action = 0 : SpChEntry.Focus()
                        Exit Function
                    End If
                    If (SpChEntry.Col = 9) Then
                        SpChEntry.Row = intRwCount : SpChEntry.Col = 8 : intFromBox = Val(Trim(SpChEntry.Text))
                        SpChEntry.Row = intRwCount : SpChEntry.Col = 9
                        If Val(Trim(SpChEntry.Text)) < intFromBox Then
                            QuantityCheck = True
                            Call ConfirmWindow(10235, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            SpChEntry.Row = intRwCount : SpChEntry.Col = 9 : SpChEntry.Action = 0 : SpChEntry.Focus()
                            Exit Function
                        End If
                    End If
                End If
            Next intcol
        Next intRwCount
        For intRwCount = 1 To SpChEntry.MaxRows
            varItemCode = Nothing
            varItemQty = Nothing
            varDrgNo = Nothing
            Call SpChEntry.GetText(1, intRwCount, varItemCode)
            Call SpChEntry.GetText(5, intRwCount, varItemQty)
            Call SpChEntry.GetText(2, intRwCount, varDrgNo)
            'ISSUE ID 10266201 
            'If Not (mblnInvocieforMTL = True Or mblnInvoicelike_MTLsharjah = True) Then
            If Not ((mblnInvocieforMTL = True) Or (InvAgstDispAdvise() = True And mblncustomer_agstdispatchadvice = True)) Then
                If CheckcustorddtlQty("ADD", CStr(varItemCode), CStr(varDrgNo), CDbl(varItemQty)) = False Then
                    QuantityCheck = True
                    Exit Function
                End If
            End If
            'ISSUE ID 10266201
            If CheckMeasurmentUnit(varItemCode, varItemQty, intRwCount) = False Then
                QuantityCheck = True
                Exit Function
            End If
        Next
        Dim var_ItemCode As Object
        If ((Trim(CmbInvType.Text) = "NORMAL INVOICE") And ((Trim(CmbInvSubType.Text) = "FINISHED GOODS") Or (Trim(CmbInvSubType.Text) = "TRADING GOODS"))) Or (Trim(CmbInvType.Text) = "JOBWORK INVOICE") Or (Trim(CmbInvType.Text) = "EXPORT INVOICE") Then
            strScheduleCheck = CheckMktSchedules()
            If Len(Trim(strScheduleCheck)) > 0 Then
                If strScheduleCheck = "Error" Then
                    GoTo ErrHandler
                Else
                    QuantityCheck = True
                    MsgBox(Space(20) & "Schedule Status:" & vbCrLf & strScheduleCheck, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Exit Function
                End If
                QuantityCheck = True
            End If

        End If

        'To check Current Balance from Itembal_Mst
        'If Quantity Entered Is Greater Then Cur_Bal In The ItemBal_Mst
        'Then Restrict User To Change The Entered Quantity
        '******************************************
        'To Get Item Code From Spread
        Dim strItCode As String 'To Make Item Code String
        rsSaleConf = New ClsResultSetDB
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                rsSaleConf.GetResult("select Stock_Location From saleconf where UNIT_CODE='" & gstrUNITID & "' AND invoice_type ='" & Trim(mstrInvType) & "' and sub_type ='" & Trim(mstrInvSubType) & "' AND Location_Code ='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & getDateForDB(lblDateDes.Text) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(lblDateDes.Text) & "')<=0")
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                rsSaleConf.GetResult("select Stock_Location From saleconf where UNIT_CODE='" & gstrUNITID & "' AND Description ='" & Trim(CmbInvType.Text) & "' and sub_type_Description ='" & Trim(CmbInvSubType.Text) & "' AND Location_Code ='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & getDateForDB(dtpDateDesc.Value) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(dtpDateDesc.Value) & "')<=0")
        End Select

        If Len(Trim(rsSaleConf.GetValue("Stock_Location"))) = 0 Then
            MsgBox("Please Define Stock Location in Sales Conf First", MsgBoxStyle.OkOnly, "empower")
            QuantityCheck = True
            rsSaleConf.ResultSetClose()
            rsSaleConf = Nothing
            Exit Function
        End If
        If mblnInvocieforMTL = False Then 'not executed for MTL Sharjah
            For intRwCount = 1 To Me.SpChEntry.MaxRows
                varItemCode = Nothing
                Call Me.SpChEntry.GetText(1, intRwCount, varItemCode)
                strItembal = "Select IsNull(Cur_Bal,0) as Cur_bal From ItemBal_Mst where UNIT_CODE='" & gstrUNITID & "' AND Location_Code ='" & Trim(rsSaleConf.GetValue("Stock_Location")) & "' and item_Code in('" & Trim(varItemCode) & "')"
                rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rsMktSchedule.GetNoRows > 0 Then
                    rsMktSchedule.MoveFirst()
                    strQuantity = rsMktSchedule.GetValue("Cur_Bal")
                    varItemQty = Nothing
                    Call Me.SpChEntry.GetText(5, intRwCount, varItemQty)
                    If Val(varItemQty) > Val(strQuantity) Then
                        QuantityCheck = True
                        If strQuantity = 0 Then
                            MsgBox("No Balance Available for this Item." & strQuantity, vbOKOnly, "empower")
                        Else
                            MsgBox("Quantity for Item Code : " & varItemCode & " should not be Greater than its Current Balance : " & strQuantity & " at location : " & rsSaleConf.GetValue("Stock_Location"), vbOKOnly, "empower")
                        End If
                        With Me.SpChEntry
                            .Row = intRwCount : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                        End With
                        Exit Function
                    Else
                        QuantityCheck = False
                    End If
                    rsMktSchedule.MoveNext()
                End If
            Next
        End If
        rsSaleConf.ResultSetClose()
        rsSaleConf = Nothing
        rsMktSchedule.ResultSetClose()
        rsMktSchedule = Nothing

        If UCase(Trim(mstrInvoiceType)) = "JOB" Then
            If BomCheck() = False Then
                QuantityCheck = True
                Exit Function
            Else
                QuantityCheck = False
            End If
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub RefreshForm(ByRef pstrType As String)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Refresh All The Fields
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case UCase(pstrType)
            Case "LOCATION"
                txtLocationCode.Text = "" : lblLocCodeDes.Text = ""
                txtChallanNo.Text = "" : txtCustCode.Text = "" : lblCustCodeDes.Text = ""
                txtConsCode.Text = "" : lblConsCodeDes.Text = ""
                txtCarrServices.Text = "" : txtVehNo.Text = ""
                txtExciseDuty.Text = ""
                txtAddExciseDuty.Text = ""
                txtFreight.Text = "" : txtSaleTaxType.Text = ""
                txtSalesTax.Text = ""
                txtSurcharge.Text = ""
                Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                txtLorryNo.Text = "" : txtOTLNo.Text = "" : txtBankAc.Text = "" : lblAcCodeDes.Text = "" : txtRefChallanNo.Text = ""
                'Samiksha shipaddress code changes
                txtshippingaddcode.Text = ""
                txtboxShipadddesc.Text = ""
            Case "CHALLAN"
                txtChallanNo.Text = "" : txtCustCode.Text = "" : lblCustCodeDes.Text = ""
                txtConsCode.Text = "" : lblConsCodeDes.Text = ""
                txtCarrServices.Text = "" : txtVehNo.Text = ""
                txtExciseDuty.Text = ""
                txtAddExciseDuty.Text = ""
                txtFreight.Text = "" : txtSaleTaxType.Text = ""
                txtSalesTax.Text = ""
                txtSurcharge.Text = "" : txtDispAdvNo.Text = "" : txtRefNo.Text = ""
                If CmbInvType.Items.Count > 0 Then
                    CmbInvType.SelectedIndex = 0 : CmbInvSubType.SelectedIndex = 0
                End If
                Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                txtLorryNo.Text = "" : txtOTLNo.Text = "" : txtRefChallanNo.Text = "" : txtBankAc.Text = "" : lblAcCodeDes.Text = ""
                'Samiksha shipaddress code changes
                txtshippingaddcode.Text = ""
                txtboxShipadddesc.Text = ""
        End Select
        With Me.SpChEntry
            .MaxRows = 1
            .Row = 1 : .Row2 = 1 : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub AddTransPortTypeToCombo()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Transport Type in Combo
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        VB6.SetItemString(CmbTransType, 0, "R - Road") 'Road
        VB6.SetItemString(CmbTransType, 1, "L - Rail") 'Rail
        VB6.SetItemString(CmbTransType, 2, "S - Sea") 'Sea
        VB6.SetItemString(CmbTransType, 3, "A - Air") 'Air
        VB6.SetItemString(CmbTransType, 4, "H - Hand") 'Hand
        VB6.SetItemString(CmbTransType, 5, "C - Courier") 'Courier
        CmbTransType.SelectedIndex = 0
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SelectChallanNoFromSalesChallanDtl()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Select Max.  Challan No. From SalesChallan_Dtl
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        'Revised by         :   Prashant rajpal
        'Revised Date       :   19/03/2015
        'REVISED ISSUE ID   :   10777177
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strChallanNo As String
        Dim rsChallanNo As ClsResultSetDB
        Dim strUpdateSQL As String
        '10777177 
        'strChallanNo = "Select max(Doc_No) as Doc_No from SalesChallan_Dtl where UNIT_CODE='" & gstrUNITID & "' AND Doc_No>" & 99000000
        strChallanNo = "Select Current_No From  DocumentType_Mst (nolock) WHERE UNIT_CODE='" + gstrUNITID + "' AND  " & _
                        " Doc_Type = 9999  AND fin_start_date <= CONVERT(DateTime,'" & dtpDateDesc.Value & "',103) " & _
                        " And Fin_End_date >= Convert(datetime,'" & dtpDateDesc.Value & "',103)"
        rsChallanNo = New ClsResultSetDB
        rsChallanNo.GetResult(strChallanNo, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

        If rsChallanNo.GetNoRows > 0 Then
            strChallanNo = (rsChallanNo.GetValue("Current_No") + 1).ToString

            strUpdateSQL = "UPDATE DocumentType_Mst with (ROWLOCK) Set Current_No = " & CLng(strChallanNo) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
            strUpdateSQL = strUpdateSQL + " Doc_Type = 9999 AND fin_start_date <= CONVERT(DateTime,'" & dtpDateDesc.Value & "',103) "
            strUpdateSQL = strUpdateSQL + " And Fin_End_date >= Convert(datetime,'" & dtpDateDesc.Value & "',103) "
            mP_Connection.Execute(strUpdateSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

            While Len(strChallanNo) < 6
                strChallanNo = "0" + strChallanNo
            End While
            strChallanNo = "99" + strChallanNo
            txtChallanNo.Text = strChallanNo
        Else
            MsgBox("Temporary Invoice No. Series Not Define. Invoice Entry Can Not Be Saved.", MsgBoxStyle.Information, ResolveResString(100))
            txtChallanNo.Text = ""
        End If

        rsChallanNo.ResultSetClose()
        rsChallanNo = Nothing
        '10777177 
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Sub displayDeatilsfromCustOrdHdrandDtl()

        'Return Value       :   NA
        'Function           :   To Display Sales order details on Selection of Customer Refrance

        On Error GoTo ErrHandler
        Dim strCustOrdHdr As String
        Dim rsCustOrdHdr As ClsResultSetDB
        Select Case CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                strCustOrdHdr = "Select Max(Order_date),Excise_Duty,Extra_Excise_Duty,Sales_Tax,"
                strCustOrdHdr = strCustOrdHdr & "Surcharge_Sales_Tax ,SalesTax_Type from Cust_ord_hdr"
                strCustOrdHdr = strCustOrdHdr & " Where UNIT_CODE='" & gstrUNITID & "' AND Account_code='" & txtCustCode.Text & "' and Cust_Ref ='"
                strCustOrdHdr = strCustOrdHdr & mstrRefNo & "'and Amendment_No ='" & mstrAmmNo & "'"
                strCustOrdHdr = strCustOrdHdr & " group by Excise_Duty,Extra_Excise_Duty,Sales_Tax,"
                strCustOrdHdr = strCustOrdHdr & "Surcharge_Sales_Tax,SalesTax_Type"
                rsCustOrdHdr = New ClsResultSetDB
                rsCustOrdHdr.GetResult(strCustOrdHdr, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                txtExciseDuty.Text = IIf(rsCustOrdHdr.GetValue("Excise_Duty") Is System.DBNull.Value, "0.00", rsCustOrdHdr.GetValue("Excise_Duty"))
                txtAddExciseDuty.Text = IIf(rsCustOrdHdr.GetValue("Extra_Excise_Duty") Is System.DBNull.Value, "0.00", rsCustOrdHdr.GetValue("Extra_Excise_Duty"))
                txtSalesTax.Text = IIf(rsCustOrdHdr.GetValue("Sales_tax") Is System.DBNull.Value, "0.00", rsCustOrdHdr.GetValue("Sales_tax"))
                txtSurcharge.Text = IIf(rsCustOrdHdr.GetValue("Surcharge_Sales_Tax") Is System.DBNull.Value, "0.00", rsCustOrdHdr.GetValue("Surcharge_Sales_Tax"))
                txtSaleTaxType.Text = rsCustOrdHdr.GetValue("SalesTax_Type")
                rsCustOrdHdr.ResultSetClose()
                rsCustOrdHdr = Nothing
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If (Trim(CmbInvType.Text) = "NORMAL INVOICE") Or (Trim(CmbInvType.Text) = "JOBWORK INVOICE") Then
                    If Len(Trim(txtRefNo.Text)) Then
                        strCustOrdHdr = "Select max(Order_date),Excise_Duty,Extra_Excise_Duty,Sales_Tax,"
                        strCustOrdHdr = strCustOrdHdr & "Surcharge_Sales_Tax ,SalesTax_Type from Cust_ord_hdr"
                        strCustOrdHdr = strCustOrdHdr & " Where UNIT_CODE='" & gstrUNITID & "' AND Account_code='" & txtCustCode.Text & "' and Cust_Ref ='"
                        strCustOrdHdr = strCustOrdHdr & mstrRefNo & "'and Amendment_No ='" & mstrAmmNo & "'"
                        strCustOrdHdr = strCustOrdHdr & " and active_flag = 'A' group by Excise_Duty,Extra_Excise_Duty,Sales_Tax,"
                        strCustOrdHdr = strCustOrdHdr & "Surcharge_Sales_Tax,SalesTax_Type"
                        rsCustOrdHdr = New ClsResultSetDB
                        rsCustOrdHdr.GetResult(strCustOrdHdr, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        txtExciseDuty.Text = IIf(rsCustOrdHdr.GetValue("Excise_Duty") Is System.DBNull.Value, "0.00", rsCustOrdHdr.GetValue("Excise_Duty"))
                        txtAddExciseDuty.Text = IIf(rsCustOrdHdr.GetValue("Extra_Excise_Duty") Is System.DBNull.Value, "0.00", rsCustOrdHdr.GetValue("Extra_Excise_Duty"))
                        txtSalesTax.Text = IIf(rsCustOrdHdr.GetValue("Sales_tax") Is System.DBNull.Value, "0.00", rsCustOrdHdr.GetValue("Sales_tax"))
                        txtSurcharge.Text = IIf(rsCustOrdHdr.GetValue("Surcharge_Sales_Tax") Is System.DBNull.Value, "0.00", rsCustOrdHdr.GetValue("Surcharge_Sales_Tax"))
                        txtSaleTaxType.Text = rsCustOrdHdr.GetValue("SalesTax_Type")
                        rsCustOrdHdr.ResultSetClose()
                        rsCustOrdHdr = Nothing
                    End If
                End If
        End Select
        Call DisplayDetailsInSpread()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SetMaxLengthInSpread()
        On Error GoTo ErrHandler
        Dim intRow As Short
        With Me.SpChEntry
            For intRow = 1 To .MaxRows
                .Row = intRow
                .Col = 1 : .TypeMaxEditLen = 16
                .Col = 2 : .TypeMaxEditLen = 30
                'Issue id :10227422
                'If mblnInvocieforMTL = True Or (mblnInvoicelike_MTLsharjah = True And mblncustomer_like_MTLsharjah = True) Then
                If mblnInvocieforMTL = True Or (InvAgstDispAdvise() = True And mblncustomer_agstdispatchadvice = True) Then
                    .Col = 3 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 6 : .TypeFloatMin = CDbl(Val("0.000000")) : .TypeFloatMax = CDbl("99999999999999.999999")
                Else
                    .Col = 3 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 4 : .TypeFloatMin = CDbl(Val("0.0000")) : .TypeFloatMax = CDbl("99999999999999.99999")
                End If
                'Issue id :10227422 done 
                .Col = 4 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 4 : .TypeFloatMin = CDbl(Val("0.0000")) : .TypeFloatMax = CDbl("99999999999999.9999")
                .Col = 5 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("99999999.99")
                .Col = 6 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 4 : .TypeFloatMin = CDbl("0.0000") : .TypeFloatMax = CDbl("99999999999999.9999")
                .Col = 7 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("99999999999999.99")
                .Col = 8 : .TypeMaxEditLen = 4
                .Col = 9 : .TypeMaxEditLen = 4
                'If mblnInvocieforMTL = True Or (mblnInvoicelike_MTLsharjah = True And mblncustomer_like_MTLsharjah = True) Then
                If mblnInvocieforMTL = True Or (InvAgstDispAdvise() = True And mblncustomer_agstdispatchadvice = True) Then
                    .Col = 11 : .TypeMaxEditLen = 30
                    .Col = 12 : .TypeMaxEditLen = 30
                End If

            Next intRow
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Public Function DeleteRecords() As Boolean
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Delete Records
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        DeleteRecords = False
        strupSalechallan = "Delete SalesChallan_Dtl where UNIT_CODE='" & gstrUNITID & "' AND Doc_No =" & Trim(txtChallanNo.Text)
        strupSalechallan = strupSalechallan & " and Location_Code ='" & Trim(txtLocationCode.Text) & "'"

        strupSaleDtl = "Delete Sales_Dtl where UNIT_CODE='" & gstrUNITID & "' AND Doc_No =" & Trim(txtChallanNo.Text)
        strupSaleDtl = strupSaleDtl & " and Location_Code ='" & Trim(txtLocationCode.Text) & "'"

        STREXPDET = "Delete EXPORT_SALES_EXTRA_DETAIL where UNIT_CODE='" & gstrUNITID & "' AND Doc_No =" & Trim(txtChallanNo.Text)
        STREXPDET = STREXPDET & " and Unt_CodeID ='" & Trim(txtLocationCode.Text) & "'"
        DeleteRecords = True
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function CheckMeasurmentUnit(ByRef strItem As Object, ByRef strQuantity As Object, ByRef intRow As Short) As Boolean
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   strItem - Item code
        '                       strQuantity - Quantity of ITem
        '                       introw -Current Row count in Spread
        'Return Value       :   Boolean YES OR No
        'Function           :   To check if decimal allowed flag is yes or No
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        Dim strMeasure As String
        Dim rsMeasure As ClsResultSetDB
        strMeasure = "select a.Decimal_allowed_flag from Measure_Mst a,Item_Mst b"
        strMeasure = strMeasure & " where a.unit_code=b.unit_code and b.cons_Measure_Code=a.Measure_Code and b.Item_Code = '" & strItem & "' and a.UNIT_CODE='" & gstrUNITID & "'"
        rsMeasure = New ClsResultSetDB
        rsMeasure.GetResult(strMeasure, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If rsMeasure.GetValue("Decimal_allowed_flag") = False Then
            If System.Math.Round(Double.Parse(strQuantity), 0) - Val(strQuantity) <> 0 Then
                Call ConfirmWindow(10455, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                CheckMeasurmentUnit = False
                Call SpChEntry.SetText(5, intRow, CShort(strQuantity))
                SpChEntry.Col = 5
                SpChEntry.Row = SpChEntry.ActiveRow
                SpChEntry.Focus()
                rsMeasure.ResultSetClose()
                rsMeasure = Nothing
                Exit Function
            Else
                CheckMeasurmentUnit = True
            End If
        Else
            CheckMeasurmentUnit = True
        End If
        rsMeasure.ResultSetClose()
        rsMeasure = Nothing
    End Function
    Public Function BomCheck() As Boolean
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Check Bom details & Qty required in Case of Jobwork Challan
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim intSpreadRow As Short 'Max count of spread row
        Dim intSpCurrentRow As Short 'Currunt count of spread in loop
        Dim intBomMaxItem As Short 'Max Parent row in Bom_Mst for finished Item
        Dim intCurrentItem As Short 'Current count of row in Parent loop
        Dim intBomMaxRaw As Short 'Max Count of Child Items in Bom_Mst
        Dim intCurrentRaw As Short 'Current raw count in Bom_Mst for finished row
        Dim inti As Short 'To Change Array Size
        Dim intTotalReqQty As Double 'Req_Qty + Waste_Qty in Bom_Mst
        Dim VarFinishedItem As Object 'to get finished Item code from Spread
        Dim VarFinishedQty As Object 'To get Qty of Finished Item from Spread
        Dim strCustAnnexDtl As String
        Dim strBomMst As String
        Dim strBomMstRaw As String
        Dim strBomItem As String
        Dim arrItem() As String
        Dim arrQty() As Double
        Dim strParent As String
        Dim rsCustAnnexDtl As ClsResultSetDB
        Dim rsBomMst As ClsResultSetDB
        Dim rsBomMstRaw As ClsResultSetDB
        rsBomMst = New ClsResultSetDB
        rsCustAnnexDtl = New ClsResultSetDB
        rsBomMstRaw = New ClsResultSetDB
        BomCheck = False
        intSpreadRow = SpChEntry.MaxRows
        inti = 0
        If SpChEntry.MaxRows >= 1 Then
            'Loop for Spread
            For intSpCurrentRow = 1 To intSpreadRow
                With SpChEntry
                    VarFinishedItem = Nothing
                    VarFinishedQty = Nothing
                    Call .GetText(1, intSpCurrentRow, VarFinishedItem)
                    Call .GetText(5, intSpCurrentRow, VarFinishedQty)
                End With
                'String for Parent Item in Bom_Mst
                strBomMst = "Select distinct(Item_Code),"
                strBomMst = strBomMst & " Bom_level from Bom_Mst where UNIT_CODE='" & gstrUNITID & "' AND Finished_Product_code ='"
                strBomMst = strBomMst & VarFinishedItem & "' Order By Bom_Level"
                rsBomMst.GetResult(strBomMst, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                intBomMaxItem = rsBomMst.GetNoRows
                rsBomMst.MoveFirst()
                strParent = ""
                'for Making all parent String before going to loop
                For intCurrentItem = 1 To intBomMaxItem
                    strBomItem = rsBomMst.GetValue("Item_code")
                    If Len(Trim(strParent)) > 0 Then
                        strParent = Trim(strParent) & "," & Chr(34) & strBomItem & Chr(34)
                    Else
                        strParent = Chr(34) & strBomItem & Chr(34)
                    End If
                    rsBomMst.MoveNext()
                Next
                rsBomMst.MoveFirst()
                'Loop for Parent Items
                For intCurrentItem = 1 To intBomMaxItem
                    strBomItem = ""
                    strBomItem = rsBomMst.GetValue("Item_code")
                    strParent = Replace(strParent, Chr(34) & strBomItem & Chr(34), Chr(34) & "Found" & Chr(34))
                    'String for CustAnnex_dtl
                    strCustAnnexDtl = "Select Item_Code,Balance_qty from CustAnnex_hdr where UNIT_CODE='" & gstrUNITID & "' AND Customer_code ='"
                    strCustAnnexDtl = strCustAnnexDtl & Trim(txtCustCode.Text) & "' and ref57f4_no ='"
                    strCustAnnexDtl = strCustAnnexDtl & Trim(txtAnnex.Text) & "' and " & GetServerDate() <= ""
                    strCustAnnexDtl = strCustAnnexDtl & " DateAdd(d, 180, ref57f4_date)"
                    strCustAnnexDtl = strCustAnnexDtl & " and Item_code ='" & strBomItem & "'"
                    rsCustAnnexDtl.GetResult(strCustAnnexDtl, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    If rsCustAnnexDtl.GetNoRows >= 1 Then
                        rsCustAnnexDtl.MoveFirst()
                        ReDim Preserve arrItem(inti)
                        ReDim Preserve arrQty(inti)
                        arrItem(inti) = rsCustAnnexDtl.GetValue("Item_code")
                        arrQty(inti) = rsCustAnnexDtl.GetValue("Balance_qty")
                        intTotalReqQty = ParentQty(strBomItem, VarFinishedItem)
                        If arrQty(inti) < intTotalReqQty * VarFinishedQty Then
                            MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & "is" & arrQty(inti) & ".", MsgBoxStyle.OkOnly, "empower")
                            SpChEntry.Row = intSpCurrentRow
                            SpChEntry.Col = 5
                            SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            BomCheck = False
                            Exit Function

                        End If
                    Else
                        strBomMstRaw = "Select RawMaterial_Code,Required_qty + Waste_qty "
                        strBomMstRaw = strBomMstRaw & " As TotalReqQty from Bom_Mst where UNIT_CODE='" & gstrUNITID & "' AND "
                        strBomMstRaw = strBomMstRaw & " item_Code ='" & strBomItem
                        strBomMstRaw = strBomMstRaw & "'and finished_product_code ='"
                        strBomMstRaw = strBomMstRaw & VarFinishedItem & "'"
                        rsBomMstRaw.GetResult(strBomMstRaw, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                        intBomMaxRaw = rsBomMstRaw.GetNoRows
                        rsBomMstRaw.MoveFirst()
                        For intCurrentRaw = 1 To intBomMaxRaw
                            strBomItem = ""
                            strBomItem = rsBomMstRaw.GetValue("RawMaterial_code")
                            intTotalReqQty = rsBomMstRaw.GetValue("TotalReqQty")
                            strCustAnnexDtl = "Select Item_Code,Balance_qty from CustAnnex_hdr where UNIT_CODE='" & gstrUNITID & "' AND Customer_code ='"
                            strCustAnnexDtl = strCustAnnexDtl & Trim(txtCustCode.Text) & "' and ref57f4_no ='"
                            strCustAnnexDtl = strCustAnnexDtl & Trim(txtAnnex.Text) & "' and " & GetServerDate() <= ""
                            strCustAnnexDtl = strCustAnnexDtl & " DateAdd(d, 180, ref57f4_date)"
                            strCustAnnexDtl = strCustAnnexDtl & " and Item_code ='" & strBomItem & "'"
                            rsCustAnnexDtl.ResultSetClose()
                            rsCustAnnexDtl = Nothing
                            rsCustAnnexDtl = New ClsResultSetDB
                            rsCustAnnexDtl.GetResult(strCustAnnexDtl, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                            If rsCustAnnexDtl.GetNoRows >= 1 Then
                                rsCustAnnexDtl.MoveFirst()
                                ReDim Preserve arrItem(inti)
                                ReDim Preserve arrQty(inti)
                                arrItem(inti) = rsCustAnnexDtl.GetValue("Item_code")
                                arrQty(inti) = rsCustAnnexDtl.GetValue("Balance_qty")
                                If arrQty(inti) < intTotalReqQty * VarFinishedQty Then
                                    MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & " is " & arrQty(inti) & " .", MsgBoxStyle.OkOnly, "empower")
                                    SpChEntry.Row = intSpCurrentRow
                                    SpChEntry.Col = 5
                                    SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    BomCheck = False
                                    Exit Function
                                End If
                            Else
                                If InStr(1, strParent, Chr(34) & strBomItem & Chr(34), CompareMethod.Text) = 0 Then
                                    MsgBox("Item " & strBomItem & " is not supplied.", MsgBoxStyle.OkOnly, "empower")
                                    SpChEntry.Row = intSpCurrentRow
                                    SpChEntry.Col = 5
                                    SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    BomCheck = False
                                    Exit Function
                                End If
                            End If
                            rsBomMstRaw.MoveNext()
                        Next
                    End If
                    rsBomMst.MoveNext()
                    inti = inti + 1
                Next  'Parent Item Loop
                intSpCurrentRow = intSpCurrentRow + 1
            Next  'Spread Item Loop
        End If
        rsBomMst.ResultSetClose()
        rsBomMst = Nothing
        rsCustAnnexDtl.ResultSetClose()
        rsCustAnnexDtl = Nothing
        rsBomMstRaw.ResultSetClose()
        rsBomMstRaw = Nothing
        BomCheck = True
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function ParentQty(ByRef pstrItemCode As String, ByRef pstrfinished As Object) As Double
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   pstrItemCode - Item Code to be Calculated from BOM
        '                       pstrfinished - Finished Product code For which invoice has to be done
        'Return Value       :   Quantity
        'Function           :   To Used in Jobwork invoice while Bom consideration
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strParentQty As String
        Dim rsParentQty As ClsResultSetDB

        strParentQty = "select sum(required_qty + waste_Qty) as TotalQty from Bom_Mst where UNIT_CODE='" & gstrUNITID & "' AND finished_Product_code ='"
        strParentQty = strParentQty & pstrfinished & "' and rawMaterial_Code ='" & pstrItemCode & "'"
        rsParentQty = New ClsResultSetDB
        rsParentQty.GetResult(strParentQty, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

        ParentQty = rsParentQty.GetValue("TotalQty")
        rsParentQty.ResultSetClose()
        rsParentQty = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function StockLocationSalesConf(ByRef pstrInvType As String, ByRef pstrInvSubtype As String, ByRef pstrFeild As String, ByRef pstrCondition As String) As String
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   pstrInvType - Invoice type
        '                       pstrInvSubtype -invoice Sub type
        '                       pstrFeild Feild - to Be Selected
        '                       pstrCondition - Condition in Query
        'Return Value       :   Stock Location as String
        'Function           :   To Check Stock Location Acc to Selected invoice type & sub type.
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        Dim rsSalesConf As ClsResultSetDB
        Dim StockLocation As String
        rsSalesConf = New ClsResultSetDB
        Select Case pstrFeild
            Case "DESCRIPTION"
                rsSalesConf.GetResult("Select Stock_Location from SaleConf Where UNIT_CODE='" & gstrUNITID & "' AND Description ='" & Trim(pstrInvType) & "' and Sub_type_Description ='" & Trim(pstrInvSubtype) & "' and " & pstrCondition)
            Case "TYPE"
                rsSalesConf.GetResult("Select Stock_Location from SaleConf Where UNIT_CODE='" & gstrUNITID & "' AND Invoice_type ='" & Trim(pstrInvType) & "' and Sub_type ='" & Trim(pstrInvSubtype) & "' and " & pstrCondition)
        End Select
        If rsSalesConf.GetNoRows > 0 Then
            StockLocation = rsSalesConf.GetValue("Stock_Location")
        End If
        StockLocationSalesConf = StockLocation
        rsSalesConf.ResultSetClose()
        rsSalesConf = Nothing
    End Function
    Public Sub EDitExpDetails()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Save Export details in String
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        strExpEditDetails = ""
        strExpEditDetails = ArrExpDetails(0) & "§" & ArrExpDetails(1) & "§" & ArrExpDetails(2) & "§" & ArrExpDetails(3) & "§" & ArrExpDetails(4) & "§" & ArrExpDetails(5) & "§" & ArrExpDetails(6) & "§" & ArrExpDetails(7) & "§" & ArrExpDetails(8) & "§" & ArrExpDetails(9) & "§" & ArrExpDetails(10) & "§" & ArrExpDetails(11) & "§" & ArrExpDetails(12) & "§" & ArrExpDetails(13) & "§" & ArrExpDetails(14) & "§" & ArrExpDetails(15)
    End Sub
    Private Function DateIsAppropriate() As Boolean
        'Function           :   Checks for Specified Date is within LIMITs From SalesChallan_DTL
        On Error GoTo ErrHandler
        Dim MaxInvoiceDate As Date 'Get Max Date of Last Invoice made
        Dim CurrentDate As Date
        MaxInvoiceDate = CDate(SelectDataFromTable("INVOICE_DATE", "SalesChallan_Dtl", " WHERE UNIT_CODE='" & gstrUNITID & "' AND BILL_FLAG = 1 and invoice_type = 'EXP' ORDER BY INVOICE_DATE "))
        CurrentDate = GetServerDate()
        If (CurrentDate >= dtpDateDesc.Value) And (dtpDateDesc.Value >= MaxInvoiceDate) Then
            DateIsAppropriate = True
        Else
            DateIsAppropriate = False
        End If
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function SelectDataFromTable(ByRef mstrFieldName As String, ByRef mstrTableName As String, ByRef mstrCondition As String) As String
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   Get Data from BackEnd
        'Comments           :   NA
        'Creation Date      :   27/06/2002
        '*******************************************************************************
        Dim StrSQLQuery As String
        Dim GetDataFromTable As ClsResultSetDB
        On Error GoTo ErrHandler
        StrSQLQuery = "Select " & mstrFieldName & " From " & mstrTableName & mstrCondition
        GetDataFromTable = New ClsResultSetDB
        If GetDataFromTable.GetResult(StrSQLQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic) Then
            If GetDataFromTable.GetNoRows > 0 Then
                SelectDataFromTable = GetDataFromTable.GetValue(mstrFieldName)
            Else
                SelectDataFromTable = CStr(GetServerDate())
            End If
        Else
            SelectDataFromTable = CStr(GetServerDate())
        End If
        GetDataFromTable.ResultSetClose()
        GetDataFromTable = Nothing
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function GetCurrencyINSO(ByVal pstrMode As String) As String
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   Get Data from BackEnd
        'Comments           :   NA
        'Creation Date      :   27/06/2002
        '*******************************************************************************
        Dim StrSQLQuery As String
        Dim GetDataFromTable As ClsResultSetDB
        On Error GoTo ErrHandler
        If Trim(pstrMode) = "ADD" Then
            StrSQLQuery = "SELECT Currency_code FROM Cust_Ord_Hdr WHERE UNIT_CODE='" & gstrUNITID & "' AND Account_code='" & Trim(txtCustCode.Text) & "' AND Cust_Ref='" & Trim(txtRefNo.Text) & "' AND Po_Type='E'"
            GetDataFromTable = New ClsResultSetDB
            If GetDataFromTable.GetResult(StrSQLQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic) Then
                If GetDataFromTable.GetNoRows > 0 Then
                    GetCurrencyINSO = GetDataFromTable.GetValue("Currency_Code")
                Else
                    GetCurrencyINSO = ""
                End If
            Else
                GetCurrencyINSO = ""
            End If
        Else
            StrSQLQuery = "SELECT Currency_code FROM SalesChallan_dtl WHERE UNIT_CODE='" & gstrUNITID & "' AND Location_Code ='" & Trim(txtLocationCode.Text) & "' AND Doc_No=" & Trim(txtChallanNo.Text)
            GetDataFromTable = New ClsResultSetDB
            If GetDataFromTable.GetResult(StrSQLQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic) Then
                If GetDataFromTable.GetNoRows > 0 Then
                    GetCurrencyINSO = GetDataFromTable.GetValue("Currency_Code")
                Else
                    GetCurrencyINSO = ""
                End If
            Else
                GetCurrencyINSO = ""
            End If
        End If
        GetDataFromTable.ResultSetClose()
        GetDataFromTable = Nothing
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GetCurrencyINSO = ""
    End Function
    Private Function CalculateTotalInvoiceAmount() As Double
        '---------------------------------------------------------------------------------------
        'Name       :   CalculateTotalInvoiceAmount
        'Type       :   Function
        'Author     :   Tapan Jain
        Dim lintLoopCounter As Short
        Dim ldblRate As Double
        Dim ldblQuantity As Double
        Dim ldblFinalAmount As Double
        Dim strsql As String
        On Error GoTo ErrHandler
        CalculateTotalInvoiceAmount = 0
        ldblQuantity = 0
        ldblRate = 0
        ldblFinalAmount = 0
        With SpChEntry
            For lintLoopCounter = 1 To .MaxRows
                .Row = lintLoopCounter
                .Col = 3
                ldblRate = Val(.Text)

                .Col = 5
                ldblQuantity = Val(.Text)
                ldblFinalAmount = ldblFinalAmount + Val(CStr(ldblQuantity * ldblRate))

            Next
        End With
        '10871426
        If mblnInvocieforMTL = True Then
            strsql = "select dbo.UFN_ROUNDOFF_DECIMAL_VALUE_MTL(" & ldblFinalAmount & " )"
            CalculateTotalInvoiceAmount = SqlConnectionclass.ExecuteScalar(strsql)
            If Len(txtSalesTax.Text) > 0 Then
                CalculateTotalInvoiceAmount = CalculateTotalInvoiceAmount + txtSalesTax.Text
            Else
                CalculateTotalInvoiceAmount = CalculateTotalInvoiceAmount
            End If

        Else
            CalculateTotalInvoiceAmount = ldblFinalAmount
        End If
        '10871426
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        CalculateTotalInvoiceAmount = 0
    End Function

    Private Function InvAgstDispAdvise() As Boolean
        '-------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Function      : To Read Sales_Parameter table to check if Invoice Agst Despatch advise is on
        ' Datetime      : 12-Feb-2007
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strQry As String
        Dim Rs As ClsResultSetDB
        InvAgstDispAdvise = False
        strQry = "select ExpInvAgstDispAdvice from Sales_Parameter where UNIT_CODE='" & gstrUNITID & "' "
        Rs = New ClsResultSetDB
        If Rs.GetResult(strQry) = False Then GoTo ErrHandler
        If Rs.GetValue("ExpInvAgstDispAdvice") = "True" Then
            InvAgstDispAdvise = True
        End If
        Rs.ResultSetClose()
        Rs = Nothing
        Exit Function
ErrHandler:
        Rs.ResultSetClose()
        Rs = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Sub FillDataAgstDispAdv()
        '-------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Function      : To fill data in grid in ADD and VIEW modes if Invoice is agst Despatch Advise
        ' Datetime      : 12-Feb-2007
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim Com As ADODB.Command
        Dim Rs As ADODB.Recordset
        Dim strsql As String
        Dim lngDispadvNo As Integer
        Dim intbinquantity As Integer

        Select Case CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                strsql = "SELECT CUST_REF,AMENDMENT_NO From BAR_DISPATCHADVICE_HDR"
                strsql = strsql & " WHERE UNIT_CODE='" & gstrUNITID & "' AND DOCNO ='" & Trim(txtDispAdvNo.Text) & "' AND STATUS = 0 And CustomerCode = '" & Trim(txtCustCode.Text) & "' And isnull(InvoiceNo,0) = 0"
                Rs = New ADODB.Recordset
                Rs.Open(strsql, mP_Connection)
                If Rs.BOF And Rs.EOF Then
                    MsgBox("No Data Found !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Rs.Close()
                    Rs = Nothing
                    Exit Sub
                Else
                    lngDispadvNo = CInt(Trim(txtDispAdvNo.Text))
                    txtRefNo.Text = Rs.Fields("Cust_Ref").Value

                    If Not ((mblnInvocieforMTL = True) Or (InvAgstDispAdvise() = True And mblncustomer_agstdispatchadvice = True)) Then
                        txtDispAdvNo.Text = CStr(lngDispadvNo)
                        mstrRefNo = Trim(txtRefNo.Text)
                        mstrAmmNo = Rs.Fields("Amendment_no").Value
                    End If
                    'If mblnInvoicelike_MTLsharjah = False Then
                    '   txtDispAdvNo.Text = CStr(lngDispadvNo)
                    '  mstrRefNo = Trim(txtRefNo.Text)
                    ' mstrAmmNo = Rs.Fields("Amendment_no").Value
                    'End If

                End If
                Rs.Close()
                Rs = Nothing

                Com = New ADODB.Command
                Rs = New ADODB.Recordset
                With Com
                    .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    'If mblnInvocieforMTL = True Or (mblnInvoicelike_MTLsharjah = True And mblncustomer_like_MTLsharjah = True) Then
                    If mblnInvocieforMTL = True Or (InvAgstDispAdvise() = True And mblncustomer_agstdispatchadvice = True And mblnInvoicelike_MTLsharjah = True) Then
                        .CommandText = "INVOICE_AGST_DISPATCHADVICE_MTL"
                    Else
                        .CommandText = "INVOICE_AGST_DISPATCHADVICE"
                    End If

                    .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                    .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adInteger, Trim(txtDispAdvNo.Text)))
                    .Parameters.Append(.CreateParameter("@MODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 5, "ADD"))
                    .Parameters.Append(.CreateParameter("@ERR", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 500))
                    .let_ActiveConnection(mP_Connection)
                    Rs = .Execute
                    If Len(.Parameters(.Parameters.Count - 1).Value) > 0 Then
                        MsgBox(.Parameters(.Parameters.Count - 1).Value, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        Com = Nothing
                        Exit Sub
                    End If
                    Com = Nothing
                End With

                If Not (Rs.BOF And Rs.EOF) Then
                    With SpChEntry
                        .MaxRows = 0
                        mstrItemCode = ""
                        While Not Rs.EOF
                            Call AddBlankRow()
                            Call .SetText(1, .MaxRows, Rs.Fields("Item_Code"))
                            Call .SetText(2, .MaxRows, Rs.Fields("Cust_DrgNo"))
                            mstrItemCode = mstrItemCode & ",'" & Trim(Rs.Fields("Cust_DrgNo").Value) & "'"
                            Call .SetText(3, .MaxRows, Rs.Fields("Rate"))
                            Call .SetText(4, .MaxRows, Rs.Fields("Cust_Mtrl"))
                            Call .SetText(5, .MaxRows, Rs.Fields("Despatch_Qty"))
                            Call .SetText(6, .MaxRows, Rs.Fields("Packing"))
                            Call .SetText(7, .MaxRows, Rs.Fields("Others"))
                            'If mblnInvocieforMTL = True Or (mblnInvoicelike_MTLsharjah = True And mblncustomer_like_MTLsharjah = True) Then
                            If mblnInvocieforMTL = True Or (InvAgstDispAdvise() = True And mblncustomer_agstdispatchadvice = True And mblnInvoicelike_MTLsharjah = True) Then
                                Call .SetText(10, .MaxRows, Rs.Fields("boxQuantity"))
                                Call .SetText(11, .MaxRows, Rs.Fields("Cust_Ref"))
                                Call .SetText(12, .MaxRows, Rs.Fields("Amendment_no"))
                                .Row2 = .MaxRows : .Col = 10 : .Col2 = 12 : .BlockMode = True : .Lock = True : .BlockMode = False
                                mstrCreditTermId = Find_Value("select term_payment from cust_ord_hdr where cust_ref='" & Rs.Fields("Cust_Ref").Value & "' and amendment_no='" & Rs.Fields("Amendment_no").Value & "'")
                            Else
                                '10895403
                                mstrCreditTermId = Find_Value("select term_payment from cust_ord_hdr where cust_ref='" & mstrRefNo & "' and amendment_no='" & mstrAmmNo & "'")
                                intbinquantity = Find_Value("select BinQuantity from custitem_mst where unit_code='" & gstrUNITID & "' and account_code='" & Trim(txtCustCode.Text) & "' and Cust_Drgno='" & Rs.Fields("Cust_DrgNo").Value & "' and item_code='" & Rs.Fields("item_code").Value & "' and active=1 ")
                                Call .SetText(10, .MaxRows, intbinquantity)
                                '10895403
                            End If

                            Rs.MoveNext()
                        End While
                        mstrItemCode = Mid(mstrItemCode, 2)
                        .Row = 1
                        .Row2 = .MaxRows
                        .Col = 1
                        .Col2 = 5
                        .BlockMode = True
                        .Lock = True
                        .BlockMode = False
                        ReDim mdblPrevQty(.MaxRows - 1) ' To get value of Quantity in Arrey for updation in despatch
                        ReDim mdblToolCost(.MaxRows - 1) ' To get value of Quantity i
                    End With
                Else
                    MsgBox("No Data is available to be displayed !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Exit Sub
                End If
                If txtDispAdvNo.Enabled Then txtDispAdvNo.Focus()
                Rs.Close()
                Rs = Nothing
                Com = Nothing
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT

                strsql = "SELECT DOCNO FROM BAR_DISPATCHADVICE_HDR"
                strsql = strsql & " WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMERCODE ='" & Trim(txtCustCode.Text) & "'"
                strsql = strsql & " AND INVOICENO ='" & Trim(txtChallanNo.Text) & "'"
                Rs = New ADODB.Recordset
                Rs.Open(strsql, mP_Connection)
                If Rs.BOF And Rs.EOF Then
                    Rs.Close()
                    Rs = Nothing
                    Exit Sub
                Else
                    txtDispAdvNo.Text = Trim(Rs.Fields("docno").Value)
                End If
                Rs.Close()
                Rs = Nothing
                Rs = New ADODB.Recordset
                strsql = "SELECT CUST_ITEM_CODE FROM SALES_DTL WHERE UNIT_CODE='" & gstrUNITID & "' AND DOC_NO = '" & Trim(txtChallanNo.Text) & "'"
                Rs.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If Not (Rs.BOF And Rs.EOF) Then
                    With SpChEntry
                        .MaxRows = 0
                        .MaxRows = 1
                        mstrItemCode = ""
                        While Not Rs.EOF
                            mstrItemCode = mstrItemCode & ",'" & Trim(Rs.Fields("Cust_Item_Code").Value) & "'"
                            Rs.MoveNext()
                        End While
                        mstrItemCode = Mid(mstrItemCode, 2)
                        Call DisplayDetailsInSpread()
                        .Row = 1
                        .Row2 = .MaxRows
                        .Col = 1
                        .Col2 = 5
                        .BlockMode = True
                        .Lock = True
                        .BlockMode = False
                    End With
                Else
                    MsgBox("No Data is available to be displayed !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                End If
                Rs.Close()
                Rs = Nothing

        End Select

        Exit Sub
ErrHandler:

        Rs = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub AddBlankRow()
        ' Function      : To Add a Blank row in Grid
        On Error GoTo ErrHandler
        With SpChEntry
            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            .ColsFrozen = 1
            .ColsFrozen = 2
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .set_RowHeight(.Row, 300)
            .Col = 4 ''Cust Matt
            .ColHidden = True
            .Col = 6 ''Packing
            .ColHidden = True
            .Col = 7 ''Others
            .ColHidden = True
            If .MaxRows > 4 Then .ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsBoth
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Function CheckMktSchedules() As String
        '-------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : NIL
        ' Return Value  : 'Error'  - If error occured during processing
        '                 Msg if Schedule doesn't exist for Item(s)
        ' Function      : To Update Daily and Monthly Schedules
        ' Datetime      : 12-Feb-2007
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim Com As ADODB.Command
        Dim strsql As String
        Dim intCtr As Short
        Dim strMSG As String
        Dim strYYYYmm As String

        ReDim mSchTypeArr(0)
        CheckMktSchedules = ""
        strYYYYmm = Year(ConvertToDate(lblDateDes.Text)) & VB.Right("0" & Month(ConvertToDate(lblDateDes.Text)), 2)

        With SpChEntry
            For intCtr = 1 To .MaxRows Step 1
                ReDim Preserve mSchTypeArr(intCtr)
                Com = New ADODB.Command
                Com.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Com.CommandText = "MKT_SCHUDULE_CHECK"
                .Row = intCtr
                Com.Parameters.Append(Com.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                Com.Parameters.Append(Com.CreateParameter("@CUSTOMER_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(txtCustCode.Text)))
                Com.Parameters.Append(Com.CreateParameter("@CONSIGNEE_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(txtConsCode.Text)))
                .Col = 1
                Com.Parameters.Append(Com.CreateParameter("@ITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(.Text)))
                .Col = 2
                Com.Parameters.Append(Com.CreateParameter("@CUSTDRG_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(.Text)))
                .Col = 5
                Com.Parameters.Append(Com.CreateParameter("@REQ_QTY", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adCurrency, Trim(.Text) - mdblPrevQty(intCtr - 1)))
                Com.Parameters.Append(Com.CreateParameter("@YYYYMM", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adInteger, strYYYYmm))
                Com.Parameters.Append(Com.CreateParameter("@SCH_TYPE", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamOutput, 1))
                Com.Parameters.Append(Com.CreateParameter("@MSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 500))
                Com.Parameters.Append(Com.CreateParameter("@ERR", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 100))
                Com.let_ActiveConnection(mP_Connection)
                Com.Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                If Len(Com.Parameters(9).Value) > 0 Then
                    MsgBox(Com.Parameters(9).Value, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    CheckMktSchedules = "Error"
                    Com = Nothing
                    Exit Function
                End If
                If Len(Com.Parameters(8).Value) > 0 Then
                    strMSG = strMSG & Com.Parameters(8).Value
                End If
                mSchTypeArr(intCtr) = Com.Parameters(7).Value
                Com = Nothing
            Next intCtr
        End With
        CheckMktSchedules = strMSG
        Exit Function
ErrHandler:
        Com = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Function UpdateMktSchedules(ByVal pstrUpdType As String) As Boolean
        '-------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : '+' - If Despatch is to be Updated agst Schedule
        '                 '-' - If Reversal is to be made agst Despatched Qty
        ' Return Value  : TRUE  - If Successfull
        '                 FALSE - If error Occured during processing
        ' Function      : To Update Daily and Monthly Schedules
        ' Datetime      : 12-Feb-2007
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim Com As ADODB.Command
        Dim strsql As String
        Dim intCtr As Short
        Dim strMSG As String
        Dim strYYYYmm As String
        Dim curQty As Decimal

        UpdateMktSchedules = True

        strYYYYmm = Year(ConvertToDate(lblDateDes.Text)) & VB.Right("0" & Month(ConvertToDate(lblDateDes.Text)), 2)
        With SpChEntry
            For intCtr = 1 To .MaxRows Step 1
                Com = New ADODB.Command
                Com.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Com.CommandText = "MKT_SCHUDULE_KNOCKOFF"
                .Row = intCtr
                Com.Parameters.Append(Com.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                Com.Parameters.Append(Com.CreateParameter("@CUSTOMER_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(txtCustCode.Text)))
                Com.Parameters.Append(Com.CreateParameter("@CONSIGNEE_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(txtConsCode.Text)))
                .Col = 1
                Com.Parameters.Append(Com.CreateParameter("@ITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(.Text)))
                .Col = 2
                Com.Parameters.Append(Com.CreateParameter("@CUSTDRG_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(.Text)))
                Com.Parameters.Append(Com.CreateParameter("@FLAG", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, pstrUpdType))
                Com.Parameters.Append(Com.CreateParameter("@SCH_TYPE", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(mSchTypeArr(intCtr))))
                Com.Parameters.Append(Com.CreateParameter("@YYYYMM", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adInteger, strYYYYmm))
                .Col = 5
                If pstrUpdType = "+" Then
                    Com.Parameters.Append(Com.CreateParameter("@REQ_QTY", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adCurrency, Trim(.Text)))
                Else
                    Com.Parameters.Append(Com.CreateParameter("@REQ_QTY", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adCurrency, mdblPrevQty(intCtr - 1)))
                End If
                Com.Parameters.Append(Com.CreateParameter("@ERR", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 100))

                Com.let_ActiveConnection(mP_Connection)
                Com.Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                If Len(Com.Parameters(9).Value) > 0 Then
                    MsgBox(Com.Parameters(9).Value, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    UpdateMktSchedules = False
                    Com = Nothing
                    Exit Function
                End If

                Com = Nothing
            Next intCtr
        End With
        Exit Function
ErrHandler:

        Com = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function CheckScannedPallete(ByRef pstrDocNo As String) As Boolean
        '-------------------------------------------------------------------------------------------
        ' Author        : Manoj Kr. Vaish
        ' Arguments     :
        ' Return Value  : TRUE  - If Successfull
        '                 FALSE - If error Occured during processing
        ' Function      : To check scanned pallete
        ' Datetime      : 07-Sep-2007
        '--------------------------------------------------------------------------------------------------
        Dim strsql As String
        Dim intCount As Short
        Dim rsdispatch As ClsResultSetDB

        rsdispatch = New ClsResultSetDB
        If Len(Trim(txtDispAdvNo.Text)) > 0 Then
            strsql = "select count(1) as totalcount from bar_DispatchAdvice_dtl where unit_code='" & gstrUNITID & "' and docno='" & pstrDocNo & "'"
            Call rsdispatch.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsdispatch.GetNoRows > 0 Then
                intCount = rsdispatch.GetValue("totalcount")
            End If
            strsql = "select count(1) as totalscanned from bar_DispatchAdvice_dtl where unit_code='" & gstrUNITID & "' and  Qty = pallete_scanned_qty And docno ='" & pstrDocNo & "'"
            Call rsdispatch.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsdispatch.GetNoRows > 0 Then
                If intCount = rsdispatch.GetValue("totalscanned") Then
                    CheckScannedPallete = True
                Else
                    CheckScannedPallete = False
                End If
            End If
            rsdispatch.ResultSetClose()
            rsdispatch = Nothing
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Function
    Private Sub SpChEntry_Change(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ChangeEvent) Handles SpChEntry.Change
        Dim intRowCount As Short
        Dim intmaxrows As Short
        Dim varFromBox As Object
        Dim varItem As Object
        Dim VarToBox As Object
        Dim varQty As Object
        Dim boxqty As Object
        Dim GetDataFromTable As ClsResultSetDB
        Dim strsql As String
        Dim bagqty As Object
        Dim varCumulativeBoxes As Object
        Dim varMaxQty As Object
        If GetPlantName() = "WCS" Then
            With SpChEntry
                If e.col = 5 Then
                    intmaxrows = SpChEntry.MaxRows
                    For intRowCount = 1 To intmaxrows
                        varItem = Nothing
                        Call .GetText(1, intRowCount, varItem)
                        varQty = Nothing
                        Call .GetText(5, intRowCount, varQty)
                        If Len(varItem) = 0 Then Exit For
                        varFromBox = Nothing
                        Call .GetText(8, intRowCount, varFromBox)
                        VarToBox = Nothing
                        Call .GetText(9, intRowCount, VarToBox)
                        bagqty = Nothing
                        strsql = "SELECT ISNULL(BAG_QTY,0)AS BAG_QTY,ISNULL(BOX_QTY,0)AS BOX_QTY FROM ITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE = '" & Trim(varItem) & "'"
                        bagqty = SelectDataFromTable("BAG_QTY", "Item_mst", " WHERE UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE = '" & Trim(varItem) & "'")
                        If bagqty > 0 Then
                            If varQty > 0 Then
                                If (varQty Mod bagqty) > 0 Then
                                    'MsgBox("Sales Quantity Should be Multiple Of Bag Qty (in Product Master )" & bagqty, MsgBoxStyle.Information, ResolveResString(100))
                                    'Call .SetText(5, intRowCount, 0.0)
                                    Call .SetText(8, intRowCount, 1)
                                    Call .SetText(9, intRowCount, Int(varQty / bagqty) + 1)
                                    'Exit Sub
                                Else
                                    Call .SetText(8, intRowCount, 1)
                                    Call .SetText(9, intRowCount, Int(varQty / bagqty))
                                End If
                            End If
                        Else
                            MsgBox("Define First Bag Qty in Product master for Item Code " & varItem, MsgBoxStyle.Information, ResolveResString(100))
                        End If
                    Next
                End If
                If e.col = 8 Then
                    intmaxrows = SpChEntry.MaxRows
                    For intRowCount = 1 To intmaxrows
                        varQty = Nothing
                        Call .GetText(5, intRowCount, varQty)
                        varFromBox = Nothing
                        Call .GetText(8, intRowCount, varFromBox)
                        If varFromBox <> 1 Then
                            MsgBox("From Box Should always be =1 ", MsgBoxStyle.Information, ResolveResString(100))
                            Call .SetText(8, intRowCount, 1)
                            Call .SetText(9, intRowCount, Int(varQty / bagqty))

                        End If
                    Next
                End If
                If e.col = 9 Then
                    intmaxrows = SpChEntry.MaxRows
                    For intRowCount = 1 To intmaxrows
                        varItem = Nothing
                        Call .GetText(1, intRowCount, varItem)
                        varQty = Nothing
                        Call .GetText(5, intRowCount, varQty)
                        VarToBox = Nothing
                        Call .GetText(9, intRowCount, VarToBox)
                        bagqty = Nothing
                        strsql = "SELECT ISNULL(BAG_QTY,0)AS BAG_QTY,ISNULL(BOX_QTY,0)AS BOX_QTY FROM ITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE = '" & Trim(varItem) & "'"
                        bagqty = SelectDataFromTable("BAG_QTY", "Item_mst", " WHERE UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE = '" & Trim(varItem) & "'")
                        'If varQty > 0 And (VarToBox <> Int(varQty / bagqty)) Then
                        If varQty > 0 And (VarToBox <> Int(varQty / bagqty)) And Int(varQty / bagqty) > 0 Then
                            MsgBox("Difference b/w From Box and To Box Should always be " & Int(varQty / bagqty), MsgBoxStyle.Information, ResolveResString(100))
                            Call .SetText(8, intRowCount, 1)
                            Call .SetText(9, intRowCount, Int(varQty / bagqty))
                        End If
                    Next
                End If
            End With
        End If
        If GetPlantName() = "HILEX" Then
            With SpChEntry
                If e.col = 5 Then
                    intmaxrows = SpChEntry.MaxRows
                    For intRowCount = 1 To intmaxrows
                        varItem = Nothing
                        Call .GetText(1, intRowCount, varItem)
                        varQty = Nothing
                        Call .GetText(5, intRowCount, varQty)
                        If Len(varItem) = 0 Then Exit For
                        boxqty = Nothing
                        Call .GetText(10, intRowCount, boxqty)

                        If boxqty > 0 Then
                            If varQty > 0 Then
                                If (varQty / boxqty) - Int(varQty / boxqty) > 0 Then
                                    ''''--------------------------------Code Changed due to increament in grid colums---------------
                                    If intRowCount = 1 Then
                                        Call .SetText(8, intRowCount, 1)
                                        Call .SetText(9, intRowCount, Int(varQty / boxqty) + 1)
                                    Else
                                        VarToBox = Nothing
                                        Call .GetText(9, intRowCount - 1, VarToBox)
                                        Call .SetText(8, intRowCount, VarToBox + 1)
                                        Call .SetText(9, intRowCount, VarToBox + Int(varQty / boxqty) + 1)
                                    End If
                                Else
                                    If intRowCount = 1 Then
                                        Call .SetText(8, intRowCount, 1)
                                        Call .SetText(9, intRowCount, (Int(varQty / boxqty)))
                                    Else
                                        VarToBox = Nothing
                                        Call .GetText(9, intRowCount - 1, VarToBox)
                                        Call .SetText(8, intRowCount, VarToBox + 1)
                                        Call .SetText(9, intRowCount, VarToBox + Int(varQty / boxqty))
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End With
        End If

    End Sub
    Private Sub SpChEntry_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles SpChEntry.Enter
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   to check If Grid have items in it or not
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        Dim strItemCode As String
        If Len(Trim(txtDispAdvNo.Text)) = 0 Then
            strItemCode = Replace(mstrItemCode, "'", "")
            If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If Len(Trim(strItemCode)) = 0 Then
                    MsgBox("First Select atleast one item First", MsgBoxStyle.OkOnly, "empower")
                    If Cmditems.Enabled = True Then
                        Cmditems.Focus()
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub SpChEntry_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles SpChEntry.KeyPressEvent
        'Function           :   At Enter Key Press Set Focus To Next Control
        On Error GoTo ErrHandler
        Select Case e.keyAscii
            Case 39, 34, 96, 45
                e.keyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtExciseDuty_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles txtExciseDuty.KeyPress
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        txtAddExciseDuty.Focus()
                End Select
            Case 39, 34, 96
                e.KeyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlInsurance_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles ctlInsurance.KeyPress
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                If txtvesselnumber.Enabled = True Then txtvesselnumber.Focus()
            Case 39, 34, 96
                e.KeyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtAddExciseDuty_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles txtAddExciseDuty.KeyPress
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        ctlSVD.Focus()
                End Select
            Case 39, 34, 96
                e.KeyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub txtFreight_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles txtFreight.KeyPress
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If (CmbInvType.Text = "SAMPLE INVOICE") Or (CmbInvType.Text = "TRANSFER INVOICE") Or (CmbInvType.Text = "JOBWORK INVOICE") Then
                            txtSurcharge.Focus()
                        Else
                            txtSaleTaxType.Focus()
                        End If
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If txtSaleTaxType.Enabled Then
                            txtSaleTaxType.Focus()
                        Else
                            txtSurcharge.Focus()
                        End If
                End Select
            Case 39, 34, 96
                e.KeyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSalesTax_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles txtSalesTax.KeyPress
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        txtSurcharge.Focus()
                End Select
            Case 39, 34, 96
                e.KeyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlSVD_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles ctlSVD.KeyPress
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        ctlInsurance.Focus()
                End Select
            Case 39, 34, 96
                e.KeyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSurcharge_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles txtSurcharge.KeyPress
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        With Me.SpChEntry
                            .Row = 1 : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                        End With
                End Select
            Case 39, 34, 96
                e.KeyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtvesselnumber_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtvesselnumber.KeyPress
        '*******************************************************************************
        'Author             :   Manoj Vaish
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   eMpro(-20090227- 27987) 05 Mar 2009
        Dim KeyAscii As Short = Asc(e.KeyChar)

        Try
            Select Case KeyAscii
                Case 39, 34, 96
                    e.Handled = True
                Case Keys.Enter, Keys.Tab
                    If txtRemarks.Enabled = True Then txtRemarks.Focus()
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub SetControlsforASNDetails(ByVal pstraccountcode As String)
        'Revised By     : Manoj Kr. Vaish
        'Revised On     : 13 May 2009
        'Arguments      : Account Code
        'Issue ID       : eMpro-20090513-31282
        'Reason         : Set controls to capture additional ASN details
        '--------------------------------------------------------------------------------------
        On Error GoTo ErrHandler

        CmbTransType.Items.Clear()
        If AllowASNTextFileGeneration(pstraccountcode) = True Then
            txtPlantCode.Enabled = True
            txtPlantCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            txtActualReceivingLoc.Enabled = True
            txtActualReceivingLoc.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)

            CmbTransType.Items.Add("A - Air")
            CmbTransType.Items.Add("S - Ocean")
            CmbTransType.Items.Add("M - Motor")
            CmbTransType.Items.Add("W - Inland Waterway")
            CmbTransType.Items.Add("H - Customer Pickup")
            CmbTransType.Items.Add("R - Rail")
            CmbTransType.Items.Add("O - Containerized Ocean")
            CmbTransType.Items.Add("C - Consolidation")
            CmbTransType.Items.Add("U - UPS")
            CmbTransType.Items.Add("E - Expedited Truck")
            CmbTransType.SelectedIndex = 0

            Call SelectDescriptionForField("Plant_Code", "Customer_Code", "Customer_Mst", txtPlantCode, (txtCustCode.Text))
            If txtPlantCode.Text.Length > 0 Then txtPlantCode.Enabled = False
        Else
            txtPlantCode.Text = String.Empty
            txtActualReceivingLoc.Text = String.Empty
            txtPlantCode.Enabled = False
            txtPlantCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtActualReceivingLoc.Enabled = False
            txtActualReceivingLoc.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Call AddTransPortTypeToCombo()      'Add default transport types
        End If
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function FordASNFileGeneration(ByVal pintdocno As Integer, ByVal pstraccountcode As String) As Boolean
        'Revised By     : Manoj Kr. Vaish
        'Revised On     : 14 May 2009
        'Arguments      : INvoice No
        'Issue ID       : eMpro-20090513-31282
        'Reason         : Generate ASN File for FORD
        '--------------------------------------------------------------------------------------
        On Error GoTo ErrHandler

        Dim rsgetData As New ClsResultSetDB
        Dim strquery As String
        Dim fs As FileStream
        Dim sw As StreamWriter
        Dim strASNdata As String
        Dim Dcount As Integer
        Dim TotalQty As Double
        Dim dblSalesQty As Double
        Dim strcontainerdespQty As String
        Dim dblcummulativeQty As Double
        Dim dblContainerQty As Double
        Dim strTotalQty As String
        Dim strASNFilepath As String
        Dim strASNFilepathforEDI As String

        strASNdata = ""
        strquery = "select * from dbo.FN_GETASNDETAIL(" & pintdocno & ",'" & gstrUNITID & "')"
        rsgetData.GetResult(strquery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsgetData.GetNoRows > 0 Then
            If rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Length = 0 Then
                MessageBox.Show("Customer vendor code is not defined for Customer : " & pstraccountcode, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                FordASNFileGeneration = False
                Exit Function
            Else
                If rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim().Length = 0 Then
                    MessageBox.Show("Unable To Get Plant Code For The Customer: " & pstraccountcode & " While Generating ASN File." & vbCrLf & _
                                    "Invoice Can't Be Locked", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    FordASNFileGeneration = False
                    Exit Function
                End If
                strASNdata = "856HD20200000000000000000" & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) + rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & ",     ,     *" & vbCrLf
                strASNdata = strASNdata & "856A M" & mInvNo.ToString.Trim() & Space(10 - mInvNo.ToString.Trim().Length) & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & "09" & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & Space(10) & "+00000000100KG+00000000080KG" & "AE0N" & Space(8) & rsgetData.GetValue("TRANSPORT_TYPE").ToString() & Space("12") & "M" & VB.Right(mInvNo, 5) & Space(4) & Space(35) & "M" & VB.Right(mInvNo, 5) & Space(5) & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim() & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & Space(6) & rsgetData.GetValue("ARL_CODE").ToString.Trim() & Space(5 - rsgetData.GetValue("ARL_CODE").ToString.Trim().Length) & VB6.Format(GetServerDateTime, "mmddhhmm") & Space(3) & "0000000.00" & vbCrLf
                Dcount = 2
                strcontainerdespQty = Find_Value("select sum(isnull(to_box,0)-isnull(from_box,0)+1) as Desp_Qty from sales_dtl where unit_code ='" & gstrUNITID & "' and doc_no=" & pintdocno)
                strASNdata = strASNdata & "856TD"
                Select Case rsgetData.GetValue("CONTAINER").ToString.Trim.Length()
                    Case 3
                        strASNdata = strASNdata & rsgetData.GetValue("CONTAINER").ToString.Trim() & "90+" & Mid("000000", strcontainerdespQty.Length(), 6) & strcontainerdespQty.ToString() & vbCrLf
                    Case 4
                        strASNdata = strASNdata & rsgetData.GetValue("CONTAINER").ToString.Trim() & " +" & Mid("000000", strcontainerdespQty.Length(), 6) & strcontainerdespQty.ToString() & vbCrLf
                    Case 5
                        strASNdata = strASNdata & rsgetData.GetValue("CONTAINER").ToString.Trim() & "+" & Mid("000000", strcontainerdespQty.Length(), 6) & strcontainerdespQty.ToString() & vbCrLf
                    Case 1, 2
                        strASNdata = strASNdata & rsgetData.GetValue("CONTAINER").ToString.Trim() & Space(3 - rsgetData.GetValue("CONTAINER").ToString.Trim.Length()) & "  +" & Mid("000000", strcontainerdespQty.Length(), 6) & strcontainerdespQty.ToString() & vbCrLf
                    Case Else
                        strASNdata = strASNdata & VB.Left(rsgetData.GetValue("CONTAINER").ToString.Trim(), 5) & "+" & Mid("000000", strcontainerdespQty.Length(), 6) & strcontainerdespQty.ToString() & vbCrLf
                End Select
                Dcount = Dcount + 1

                rsgetData.MoveFirst()
                Do While Not rsgetData.EOFRecord
                    dblcummulativeQty = 0
                    dblSalesQty = 0
                    dblContainerQty = 0
                    dblcummulativeQty = Find_Value("SELECT DBO.UDF_GET_CUMMULATIVEQTY('" & gstrUNITID & "','" & rsgetData.GetValue("CUST_PLANTCODE").ToString() & "','" & rsgetData.GetValue("CUST_PART_CODE").ToString() & "'," & pintdocno & ")")
                    dblSalesQty = rsgetData.GetValue("SALES_QUANTITY")
                    dblcummulativeQty = dblcummulativeQty + dblSalesQty
                    dblContainerQty = rsgetData.GetValue("CONTAINER_QTY")
                    strASNdata = strASNdata & "856P "
                    strASNdata = strASNdata & rsgetData.GetValue("CUST_PART_CODE").ToString().Trim & Space(30 - rsgetData.GetValue("CUST_PART_CODE").ToString().Length())
                    dblSalesQty = rsgetData.GetValue("Sales_Quantity")
                    strASNdata = strASNdata & "BP+" & Mid("0000000", dblSalesQty.ToString.Length(), 8) & dblSalesQty & "EA+"
                    strTotalQty = Val(strTotalQty) + Val(rsgetData.GetValue("SALES_QUANTITY"))
                    Dcount = Dcount + 1
                    strASNdata = strASNdata & Mid("000000000", dblcummulativeQty.ToString().Length(), 10) & dblcummulativeQty
                    strASNdata = strASNdata & "+0000000000" & Space(10) & mInvNo.ToString & Space(11 - mInvNo.ToString.Length()) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString & VB6.Format(rsgetData.GetValue("INVOICE_DATE").ToString(), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE").ToString(), "hhmm") & vbCrLf
                    strASNdata = strASNdata & "856PA" & Space(30) & "+00000000000  +00000000000  " & vbCrLf
                    strASNdata = strASNdata & "856V " & "+000000000000000" & vbCrLf
                    Dcount = Dcount + 2
                    strASNdata = strASNdata & "856C +" & Mid("0000000", dblContainerQty.ToString.Length(), 8) & dblContainerQty & "+" & Mid("0000", rsgetData.GetValue("CONTAINER_DESP_QTY").ToString.Length, 5) & rsgetData.GetValue("CONTAINER_DESP_QTY").ToString & rsgetData.GetValue("CONTAINER").ToString & "90" & vbCrLf
                    Dcount = Dcount + 1
                    mstrupdateASNdtl = Trim(mstrupdateASNdtl) & "UPDATE MKT_ASN_INVDTL SET ASN_STATUS=1,CUMMULATIVE_QTY=" & dblcummulativeQty & " WHERE UNIT_CODE='" & gstrUNITID & "' AND DOC_NO=" & pintdocno & " AND CUST_PART_CODE='" & rsgetData.GetValue("CUST_PART_CODE").ToString().Trim() & "' AND CUST_PLANTCODE='" & rsgetData.GetValue("CUST_PLANTCODE").ToString().Trim & "'" & vbCrLf
                    mstrupdateASNCumFig = Trim(mstrupdateASNCumFig) & "UPDATE MKT_ASN_CUMFIG SET CUMMULATIVE_QTY=" & dblcummulativeQty & " WHERE UNIT_CODE='" & gstrUNITID & "' AND CUST_PART_CODE='" & rsgetData.GetValue("CUST_PART_CODE").ToString().Trim() & "' AND CUST_PLANTCODE='" & rsgetData.GetValue("CUST_PLANTCODE").ToString().Trim & "'" & vbCrLf
                    rsgetData.MoveNext()
                Loop

                Dcount = Dcount + 1
                strASNdata = strASNdata & "856T " & Mid("0000", Dcount.ToString.Length, 5) & Dcount & Mid("00000000", strTotalQty.ToString.Length(), 9) & strTotalQty
                If Directory.Exists(gstrASNPath) = False Then
                    Directory.CreateDirectory(gstrASNPath)
                End If
                If Directory.Exists(gstrASNPathForEDI) = False Then
                    Directory.CreateDirectory(gstrASNPathForEDI)
                End If
                strASNFilepath = gstrASNPath & "\" & mInvNo.ToString() & ".dat"
                strASNFilepathforEDI = gstrASNPathForEDI & "\" & mInvNo.ToString() & ".dat"

                fs = File.Create(strASNFilepath)
                sw = New StreamWriter(fs)
                sw.WriteLine(strASNdata)
                sw.Close()
                fs.Close()
                If File.Exists(strASNFilepathforEDI) = False Then
                    File.Copy(strASNFilepath, strASNFilepathforEDI)
                End If
                rsgetData.ResultSetClose()
                rsgetData = Nothing

                FordASNFileGeneration = True
            End If
        Else
            MessageBox.Show("Unable To Generate ASN File. Invoice Can't Be Locked", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
            FordASNFileGeneration = False
        End If

        Exit Function
ErrHandler:
        FordASNFileGeneration = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

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


    Private Function AllowASNTextFileGeneration(ByVal pstraccountcode As String) As Boolean
        'Revised By     : Manoj Kr. Vaish
        'Revised On     : 14 May 2009
        'Arguments      : Account Code
        'Return Value   : True/False
        'Issue ID       : eMpro-20090513-31282
        'Reason         : Check ASNTextFileGeneration from Customer Master
        '--------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strQry As String
        Dim Rs As ClsResultSetDB
        AllowASNTextFileGeneration = False
        strQry = "Select isnull(AllowASNTextGeneration,0) as AllowASNTextGeneration from customer_mst where UNIT_CODE='" & gstrUNITID & "' AND Customer_Code='" & Trim(pstraccountcode) & "'"
        Rs = New ClsResultSetDB
        If Rs.GetResult(strQry) = False Then GoTo ErrHandler
        If Rs.GetValue("AllowASNTextGeneration") = "True" Then
            AllowASNTextFileGeneration = True
        Else
            AllowASNTextFileGeneration = False
        End If
        Rs.ResultSetClose()
        Rs = Nothing
        Exit Function
ErrHandler:
        Rs = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function


    Private Sub txtPlantCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPlantCode.KeyPress
        On Error GoTo ErrHandler

        Dim Keyascii As Short = Asc(e.KeyChar)
        Select Case Keyascii
            Case System.Windows.Forms.Keys.Return
                If txtActualReceivingLoc.Enabled = True Then txtActualReceivingLoc.Focus()
            Case 39, 34, 96
                Keyascii = 0
        End Select
        If Keyascii = 0 Then
            e.Handled = True
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Public Function CheckcustorddtlQty(ByRef pstrMode As String, ByRef pstrItemCode As String, ByRef pstrDrgno As String, ByRef pdblQty As Double) As Boolean
        'Issue id 10266201
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim dblSaleQuantity As Double
        Dim strCustOrdDtl As String
        On Error GoTo ErrHandler
        rsCustOrdDtl = New ClsResultSetDB
        strCustOrdDtl = "Select openso,balance_Qty = order_qty - Despatch_qty from Cust_ord_dtl where "
        strCustOrdDtl = strCustOrdDtl & "unit_code='" & gstrUNITID & "' and Account_code ='" & txtCustCode.Text & "'" & " and Item_code ='"
        strCustOrdDtl = strCustOrdDtl & pstrItemCode & "' and cust_drgNo ='" & pstrDrgno
        strCustOrdDtl = strCustOrdDtl & "' and Authorized_flag = 1 and cust_ref = '" & txtRefNo.Text & "'"
        strCustOrdDtl = strCustOrdDtl & " and amendment_no='" & mstrAmmNo & "'"

        rsCustOrdDtl.GetResult(strCustOrdDtl)
        If rsCustOrdDtl.GetValue("OpenSO") = True Then
            CheckcustorddtlQty = True
        Else
            Select Case pstrMode
                Case "ADD"

                    If Val(rsCustOrdDtl.GetValue("Balance_Qty")) < pdblQty Then
                        MsgBox("Balance Quantity available in SO for Customer Part code [ " & pstrDrgno & "] is " & Val(rsCustOrdDtl.GetValue("Balance_Qty")) & ".", MsgBoxStyle.Information, ResolveResString(100))
                        CheckcustorddtlQty = False
                    Else
                        CheckcustorddtlQty = True
                    End If
                Case "EDIT"
                    
                    If Val(rsCustOrdDtl.GetValue("Balance_Qty")) < pdblQty Then
                        MsgBox("Balance Quantity available in SO for Customer Part code [ " & pstrDrgno & "] is " & Val(rsCustOrdDtl.GetValue("Balance_Qty")) & ".", MsgBoxStyle.Information, ResolveResString(100))
                        CheckcustorddtlQty = False
                    Else
                        CheckcustorddtlQty = True
                    End If
            End Select
        End If
        rsCustOrdDtl.ResultSetClose()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Function


    Private Sub txtSalesTax_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtSalesTax.Validating

    End Sub


    'Samiksha Ship address code changes

    Private Sub CmdhelpShipAddCode_Click(sender As Object, e As EventArgs) Handles CmdhelpShipAddCode.Click
        On Error GoTo ErrHandler
        Dim StrSql As String = String.Empty
        Dim StrSrvCHelp() As String

        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                StrSql = "select Distinct ISNULL(Shipping_Code,'')Shipping_Code,ISNULL(Shipping_Desc,'')Shipping_Desc,ISNULL(Ship_Address1,'')Ship_Address1,ISNULL(Ship_Address2,'')Ship_Address2,ISNULL(Ship_State,'')Ship_State,ISNULL(GSTIN_ID,'')GSTIN_ID from Customer_Shipping_Dtl where unit_code='" & gstrUNITID & "'and Customer_Code = '" & txtCustCode.Text.Trim & "'and InActive_Flag=0 "
                StrSrvCHelp = Me.ctlExportChallanEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, StrSql, "Ship Address Code Help")
                If UBound(StrSrvCHelp) <= 0 Then
                    Exit Sub
                End If
                If StrSrvCHelp(0) = "0" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtboxShipadddesc.Text = "" : txtboxShipadddesc.Focus() : Exit Sub
                Else
                    txtshippingaddcode.Text = StrSrvCHelp(0)
                    '                            lblservicedesc.Text = StrSrvCHelp(1)
                    txtboxShipadddesc.Text = StrSrvCHelp(1)
                End If
        End Select


ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred


    End Sub
End Class