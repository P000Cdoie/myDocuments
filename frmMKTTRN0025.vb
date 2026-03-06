Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.Data.SqlClient
Friend Class frmMKTTRN0025
	Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
	' File Name         :   FRMMKTTRN0025.frm
	' Function          :   Invoice Cancellation
	' Created By        :   Tapan Jain
	' Created On        :
	' Revision          :   Changes made by Jasmeet Singh Bawa on 20/01/2004 for rejection Invoice Cancellation
	' Revision          :   Changes made by Jasmeet Singh Bawa on 09/09/2004 for Rolling back of Schedules
	' Revision          :   Changes made by Nisha Rai on 14/09/2004 for Rolling back of Schedules -1(DSTracking-10623)
	' Revision          :   Changes made by Nisha Rai on 21/03/2005 For Rejection invoice Reversal Posting
	' Revision          :   Changes made by Sandeep on 30/03/2005 For Updating Cancel_flag in Rejection invoice Tracking Table
	' And Change the Help Ceteria for Challan No Bill_flag=1
	'===================================================================================
	'Revised By      : Manoj Kr. Vaish
	'Issue ID        : 19992
	'Revision Date   : 30 JUNE 2007
	'History         : To add the functionality of Multiple SO for Export Invoice.
	'***********************************************************************************
	'Revised By      : Manoj Kr. Vaish
	'Issue ID        : 21105
	'Revision Date   : 14 Sep 2007
	'History         : To add the Bar Code functionality of for MATE MANESAR
	'***********************************************************************************
	'Revised By      : Manoj Kr. Vaish
	'Issue ID        : 21551
	'Revision Date   : 21-Nov-2007
	'History         : Add New Tax VAT with Sale Tax help
	'***********************************************************************************
    'Revised By      : Manoj Kr.Vaish
    'Issue ID        : eMpro-20080930-22159
    'Revision Date   : 01 Oct 2008
    'History         : BatchWise Tracking of Invoices Made from 01M1 Location including BarCode Tracking
    '                  Knocking Off Daily Marketing Schedule on DayWise
    '*******************************************************************************************************
    '***********************************************************************************
    'Revised By      : Parveen Kumar
    'Issue ID        : eMpro-20100202-41764
    'Revision Date   : 02 FEB 2010
    'History         : At the time of invoice cancellation party code and voucher type must be checked in finance module instead of invoice no only
    'Modified By Nitin Mehta on 16 May 2011
    'Modified to support MultiUnit functionality
    '*******************************************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue id        : HILEX POST IMPLEMENTATION CHANGES
    'Revised Date    : 08 May 21014
    '*******************************************************************************************************
    'Revised By         :   Vinod Singh
    'Revision Date      :   13 Jan 2015
    'Issue id           :   10736222 - eMPro - CT2 - ARE3 functionality
    '*******************************************************************************************************
    'REVISED BY     :  VINOD SINGH
    'REVISED ON     :  06 MAY 2015
    'ISSUE ID       :  10804443 - MULTI LOCATION IN BARCODE - HILEX 
    '********************************************************************************************************
    'REVISED BY     :  PRASHANT RAJPAL
    'REVISED ON     :  27 MAR 2017
    'ISSUE ID       :  HILEX RELATED CHANGES FOR FTS ITEM 
    '********************************************************************************************************
    'REVISED BY     :  ASHISH SHARMA
    'REVISED ON     :  05 OCT 2018
    'ISSUE ID       :  101631219 - SMIIEL Auto Invoicing Functionality
    '********************************************************************************************************


    Dim mintIndex As Short 'Declared To Hold The Form Count
    Dim mdblPrevQty() As Object 'to store prev quantity in edit mode
    Dim mdblToolCost() As Object 'to insert tool cost item wise
    Dim mstrItemCode As String 'To Get The Value Of Item Code
    Dim mstrInvType As String 'To Get Value Of Inv Type From SalesChallan_Dtl
    Dim mstrInvSubType As String 'To Get Value Of Inv SubType From SalesChallan_Dtl
    Dim blnEOU_FLAG As Boolean
    Dim mstrInvNo As String 'To get the Document no against which invoice is raised
    Dim MBlnDsTracking As Boolean 'Schedule updations DS wise
    Dim strUpdateDailyMktSchedule As String
    Dim strUpdateMonthlyMktSchedule As String
    Dim strUpdateInvDsHistory As String
    Dim mblnRejTracking As Boolean
    Dim mblnMultipleSOAllowed As Boolean
    Dim mblnSORequired As Boolean
    Dim mSchTypeArr() As String
    Dim mstrupdateBarBondedStockFlag As String
    Dim mstrupdateBarBondedStockQty As String
    Dim mblnBatckTrackingAllowed As Boolean
    Dim mblnINVOICECANCELLATION_CURRENTDATE As Boolean 'Schedule updations DS wise
    Private Sub txtLocationCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtLocationCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '****************************************************
        'Created By     -  Nisha
        'Description    -  Check Validity Of Location Code In The Location_Mst
        '****************************************************
        On Error GoTo ErrHandler
        If Len(txtLocationCode.Text) > 0 Then
            If CheckExistanceOfFieldData((txtLocationCode.Text), "Location_Code", "SalesChallan_Dtl", "UNIT_CODE='" & gstrUNITID & "'") Then
                lblLocCodeDes.Text = SelectDataFromTable("Unt_UnitName", "Gen_UnitMaster", " WHERE Unt_CodeID='" & txtLocationCode.Text & "'")
                If txtChallanNo.Enabled Then
                    txtChallanNo.Focus()
                End If
            Else
                Call ConfirmWindow(10411, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtLocationCode.Text = ""
                txtLocationCode.Focus()
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
            Call RefreshForm("LOCATION")
        End If
        txtCustCode.Text = ""
        lblCustCodeDes.Text = ""
        lblLocCodeDes.Text = ""
        txtRefNo.Text = ""
        SpChEntry.MaxRows = 0
        mstrItemCode = ""
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
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                If Len(txtLocationCode.Text) > 0 Then
                    Call txtLocationCode_Validating(txtLocationCode, New System.ComponentModel.CancelEventArgs(False))
                Else
                    Me.CmdGrpChEnt.Focus()
                End If
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
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At F1 Key Press Display Help From Location Master
        '****************************************************
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdLocCodeHelp.Enabled Then Call CmdLocCodeHelp_Click(CmdLocCodeHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub CmdChallanNo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdChallanNo.Click
        '****************************************************
        'Created By     -  Nisha
        'Description    -  To Display Help From SalesChallan_Dtl
        '****************************************************
        On Error GoTo ErrHandler
        Dim strHelpString As String
        Dim strInvoiceConditionDate As String
        'Check Location Code Field
        If Trim(txtLocationCode.Text) = "" Then
            Call ConfirmWindow(10239, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO, 100)
            If txtLocationCode.Enabled Then txtLocationCode.Focus()
            Exit Sub
        End If
        If mblnINVOICECANCELLATION_CURRENTDATE = True Then
            strInvoiceConditionDate = " AND INVOICE_DATE =CONVERT(char(12), GETDATE(), 106)"
        Else
            strInvoiceConditionDate = ""
        End If
        If Len(Trim(txtChallanNo.Text)) = 0 Then
            ' Change by Sandeep Bill_Flag=1
            strHelpString = ShowList(1, (txtChallanNo.MaxLength), "", "Doc_No", DateColumnNameInShowList("Invoice_Date") & " as Invoice_Date ", "SalesChallan_Dtl ", "AND Location_Code='" & Trim(txtLocationCode.Text) & "' AND Cancel_Flag=0 and Bill_flag=1 " & strInvoiceConditionDate)
            If strHelpString = "-1" Then 'If No Record Found
                MsgBox("No Challan available for Cancellation", MsgBoxStyle.Information, "empower")
                If txtChallanNo.Enabled Then txtChallanNo.Focus()
                Exit Sub
            ElseIf strHelpString = "" Then
                txtChallanNo.Focus() 'User Cancels HELP - Nitin Sood
            Else
                txtChallanNo.Text = strHelpString
            End If
        Else
            'Changed for Issue ID 19992 Starts
            strHelpString = ShowList(1, (txtChallanNo.MaxLength), txtChallanNo.Text, "Doc_No", DateColumnNameInShowList("Invoice_Date") & " as Invoice_Date ", "SalesChallan_Dtl ", "AND Location_Code='" & Trim(txtLocationCode.Text) & "' AND Cancel_Flag=0 and Bill_Flag=1")
            'Changed for Issue ID 19992 Ends
            If strHelpString = "-1" Then 'If No Record Found
                MsgBox("No Challan available for Cancellation", MsgBoxStyle.Information, "empower")
                If txtChallanNo.Enabled Then txtChallanNo.Focus()
                Exit Sub
            ElseIf strHelpString = "" Then
                txtChallanNo.Focus() 'User Cancels HELP - Nitin Sood
            Else
                txtChallanNo.Text = strHelpString
            End If
        End If
        Call txtChallanNo_Validating(txtChallanNo, New System.ComponentModel.CancelEventArgs(False))
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Function UpDateSalesOrder() As Boolean
        '****************************************************
        'Created By     -  Nitin Sood
        'Description    -  Update DispatchQty in Cust_Ord_Dtl for Invoice Cancellation
        '                   Active Flag = 'A'
        '****************************************************
        On Error GoTo ErrHandler
        Dim intCount As Short 'For Next Loop . . .
        Dim strItemCode As String 'Item Code
        Dim dbl_Quantity As Double 'Qty
        Dim mCust_Item_Code As String 'Drwg No
        Dim strupdatecustodtdtl As String 'SQL
        UpDateSalesOrder = True
        With SpChEntry
            For intCount = 1 To .MaxRows
                .Col = 1 : .Col2 = 1 : .Row = intCount : .Row2 = intCount : strItemCode = Trim(.Text)
                .Col = 2 : .Col2 = 2 : .Row = intCount : .Row2 = intCount : mCust_Item_Code = Trim(.Text)
                .Col = 5 : .Col2 = 5 : .Row = intCount : .Row2 = intCount : dbl_Quantity = Val(.Text)
                'Make Updating Query
                'Added for Issue ID 19992 Starts
                If mblnMultipleSOAllowed = True Then
                    .Col = 16 : .Col2 = 16 : .Row = intCount : .Row2 = intCount : txtRefNo.Text = Trim(.Text)
                    .Col = 17 : .Col2 = 17 : .Row = intCount : .Row2 = intCount : txtAmendNo.Text = Trim(.Text)
                End If
                'Added for Issue ID 19992 Ends
                strupdatecustodtdtl = strupdatecustodtdtl & "Update Cust_ord_dtl set Despatch_Qty = Despatch_Qty - "
                strupdatecustodtdtl = strupdatecustodtdtl & dbl_Quantity & " ,Active_Flag = 'A' where UNIT_CODE='" & gstrUNITID & "' AND Account_code ='"
                strupdatecustodtdtl = strupdatecustodtdtl & Trim(txtCustCode.Text) & "' AND Item_Code = '" & strItemCode & "'  And Cust_DrgNo = '"
                strupdatecustodtdtl = strupdatecustodtdtl & mCust_Item_Code & "' and Cust_ref = '" & Trim(txtRefNo.Text)
                strupdatecustodtdtl = strupdatecustodtdtl & "' And Amendment_No = '" & Trim(txtAmendNo.Text) & "'"
            Next
            If Len(strupdatecustodtdtl) > 0 Then mP_Connection.Execute(strupdatecustodtdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        End With
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        UpDateSalesOrder = False
    End Function
    Private Sub CmdGrpChEnt_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As UCActXCtl.cmdGrpAuthorise.ButtonClickEventArgs) Handles CmdGrpChEnt.ButtonClick
        On Error GoTo ErrHandler
        Dim strItemCode As String
        Dim strDrgNo As String
        Dim strDSNo As String
        Dim rsDSTracking As New ClsResultSetDB
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim intRowCount As Short
        Dim strMktScheduleCheck As String
        Dim strRemoveInvFromLoadingSlip As String

        Select Case eventArgs.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_AUTHORIZE 'CANCELLATION OF INVOICE

                SqlConnectionclass.OpenGlobalConnection()
                Dim latestDispatchAdviceNo As String = String.Empty

                If Trim(txtremark.Text) = "" Then
                    MsgBox("Cancel Remark is not present", MsgBoxStyle.Information, "empower")
                    If txtremark.Enabled Then txtremark.Focus()
                    Exit Sub
                End If

                'ADDED FOR ISSUE ID 10804443 
                If ValidateInvoiceStockLocation() = False Then
                    Exit Sub
                End If
                'END OF ADDITION

                'Added by Abhijit on 27 July 2017
                If INVOICE_CHECK_IN_FRIEGT_OUTWARD() = True Then
                    MsgBox("This Invoice No. Can't be cancelled.." & vbNewLine & "Ourward Entry has been already done.", MsgBoxStyle.Information, "empower")
                    Exit Sub
                End If

                If INVOICE_CHECK_IN_OUTWARD() = True Then
                    MsgBox("This Invoice No. Can't be cancelled.." & vbNewLine & "Ourward Entry has been already done.", MsgBoxStyle.Information, "empower")
                    Exit Sub
                End If
                ' Added by Abhijit on 27 July 2017

                If CheckIRNCancel() Then
                    MsgBox("This Invoice No. Can't be cancelled.." & vbNewLine & "Please first cancel IRN No.", MsgBoxStyle.Information, "eMPower")
                    Exit Sub
                End If

                If ConfirmWindow(10054, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    mP_Connection.BeginTrans()
                    If mblnMultipleSOAllowed = False Then
                        If MBlnDsTracking Then
                            With SpChEntry
                                For intRowCount = 1 To .MaxRows 'For All Items In Spread Revise Schedule
                                    .Col = 1 : .Col2 = 1 : .Row = intRowCount : .Row2 = intRowCount : strItemCode = Trim(.Text)
                                    .Col = 2 : .Col2 = 2 : .Row = intRowCount : .Row2 = intRowCount : strDrgNo = Trim(.Text)
                                    rsDSTracking.GetResult("Select Item_code,Cust_Part_Code,DSno from Mkt_InvDSHistory where UNIT_CODE='" & gstrUNITID & "' AND Doc_No = '" & txtChallanNo.Text & "' and Item_code = '" & strItemCode & "' and Cust_Part_Code = '" & strDrgNo & "'")
                                    intMaxLoop = rsDSTracking.RowCount
                                    rsDSTracking.MoveFirst()
                                    strUpdateDailyMktSchedule = ""
                                    strUpdateInvDsHistory = ""
                                    strUpdateMonthlyMktSchedule = ""
                                    For intLoopCounter = 1 To intMaxLoop
                                        strDSNo = rsDSTracking.GetValue("DSNo")
                                        strItemCode = rsDSTracking.GetValue("Item_code")
                                        strDrgNo = rsDSTracking.GetValue("Cust_Part_code")
                                        If DispatchQuantityFromDailyScheduleForDSTracking(txtCustCode.Text, strDrgNo, strItemCode, lblDateDes.Text, strDSNo) = True Then
                                            If Not UpdateDailyMktSchedule(strItemCode, strDrgNo, strDSNo) Then mP_Connection.RollbackTrans() : Exit Sub
                                        Else
                                            If DispatchQuantityFromMonthlyScheduleForDSTracking(txtCustCode.Text, strDrgNo, strItemCode, lblDateDes.Text, strDSNo) = True Then
                                                If Not UpdateMonthlyMktSchedule(strItemCode, strDrgNo, strDSNo) Then mP_Connection.RollbackTrans() : Exit Sub
                                            Else
                                                MsgBox("No Schedule available for Item Code " & strItemCode & " and DSNo " & strDSNo & ".", MsgBoxStyle.Information, "Empower")
                                                mP_Connection.RollbackTrans()
                                                Exit Sub
                                            End If
                                        End If
                                        rsDSTracking.MoveNext()
                                    Next
                                Next
                                If Len(Trim(strUpdateInvDsHistory)) > 0 Then
                                    mP_Connection.Execute(strUpdateInvDsHistory, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                End If
                            End With
                        Else
                            If gstrUNITID <> "MS1" And gstrUNITID <> "MP2" Then
                                strMktScheduleCheck = CheckMktSchedules()
                                If Len(Trim(strMktScheduleCheck)) > 0 Then
                                    If strMktScheduleCheck = "Error" Then GoTo ErrHandler
                                End If
                                If UpdateMktSchedules("-") = False Then
                                    mP_Connection.RollbackTrans()
                                    GoTo ErrHandler
                                End If
                            End If

                        End If
                    Else
                        If gstrUNITID <> "MS1" And gstrUNITID <> "MP2" Then
                            strMktScheduleCheck = CheckMktSchedules()
                            If Len(Trim(strMktScheduleCheck)) > 0 Then
                                If strMktScheduleCheck = "Error" Then GoTo ErrHandler
                            End If
                            If UpdateMktSchedules("-") = False Then
                                mP_Connection.RollbackTrans()
                                GoTo ErrHandler
                            End If
                        End If
                        'Samiksha
                        'Serial Number Wise Despatch Qty KnockOff 
                        If gstrUNITID = "MS1" Or gstrUNITID = "MP2" Then
                            Dim oCmd = New ADODB.Command
                            With oCmd
                                .ActiveConnection = mP_Connection
                                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                .CommandText = "USP_DAILYMKT_KNOCKOFF_MTL"
                                .Parameters.Append(.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrUNITID))
                                .Parameters.Append(.CreateParameter("@InvoiceNo", ADODB.DataTypeEnum.adBigInt, ADODB.ParameterDirectionEnum.adParamInput, , Val(txtChallanNo.Text.Trim)))
                                .Parameters.Append(.CreateParameter("@ERRMSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 1000))
                                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End With
                            If Convert.ToString(oCmd.Parameters(oCmd.Parameters.Count - 1).Value) <> "" Then
                                mP_Connection.RollbackTrans()
                                MsgBox(oCmd.Parameters(oCmd.Parameters.Count - 1).Value, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                                oCmd = Nothing
                                Exit Sub
                            End If
                            oCmd = Nothing
                        End If
                    End If
                    If Not UpDateSalesOrder() Then mP_Connection.RollbackTrans() : Exit Sub
                    'AMIT RANA
                    If Me.lblInvoiceType.Text.Trim.ToUpper = "NORMAL INVOICE" And Me.lblInvoiceSubType.Text.Trim.ToUpper = "TRADING GOODS" Then
                        If Not UpdatGRN_TradingInvoice() Then
                            mP_Connection.RollbackTrans()
                            Exit Sub
                        End If
                    End If
                    'AMIT RANA
                    If UCase(mstrInvType) = "REJ" Then
                        If Len(Trim(txtRefNo.Text)) > 0 Then
                            If Not UpdateGrnHdr(CDbl(txtRefNo.Text), CDbl(Trim(txtChallanNo.Text))) Then mP_Connection.RollbackTrans() : Exit Sub
                        End If
                    End If
                    If UCase(mstrInvType) = "JOB" Then
                        If Not UpdateCustAnnex() Then mP_Connection.RollbackTrans() : Exit Sub
                    End If
                    If InvAgstBarCode() = True Then
                        If BarCodeTracking(Trim(txtChallanNo.Text)) = True Then
                            mP_Connection.Execute(mstrupdateBarBondedStockQty, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            mP_Connection.Execute(mstrupdateBarBondedStockFlag, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If
                    End If
                    If Not updatesalesconfandsaleschallan() Then mP_Connection.RollbackTrans() : Exit Sub
                    If Not UpdateinSale_Dtl() Then mP_Connection.RollbackTrans() : Exit Sub
                    If mblnBatckTrackingAllowed = True Then ''And InvoiceAgainstBatch() = True  InvoiceAgainstBatch commented by priti on 23 Jan to solve batch issue on rejection invoice
                        If Not UpdateItemBatchDetail() Then mP_Connection.RollbackTrans() : Exit Sub
                    End If

                    'ADDED AGAINST ISSUE ID 10736222 
                    mP_Connection.Execute("EXEC USP_CANCEL_CT2_INVOICE '" & gstrUNITID & "'," & Val(txtChallanNo.Text.Trim) & "", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    'END OF ADDITION
                    If UCase(Trim(GetPlantName)) = "MTL" Then
                        mP_Connection.Execute("update DispatchAdvice_ScheduleQty_Knockoff SET IsCancel = 1  where DISPATCH_ADVICE_NO in (select B.DOCNO from bar_DispatchAdvice_hdr B where B.InvoiceNo= " & Trim(txtChallanNo.Text) & "  and B.Unit_Code='" & gstrUNITID & "') and Unit_Code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute("update bar_Palette_mst set Invoice_No=null  Where Invoice_No =" & Trim(txtChallanNo.Text) & " and   unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    'Added by praveen on 07.12.2017
                    If UCase(Trim(GetPlantName)) = "HILEX" And CHECK_FTSINVOICE("CHECK_FTS") = True Then
                        'MsgBox("Invoice is made for FTS Items , cannot be Cancelled.", MsgBoxStyle.Exclamation, ResolveResString(100))
                        If CHECK_FTSINVOICE("CHECK_FTS_BARCODE_TARCKING") = True Then
                            If FTS_INVOICECANCELLATION() = False Then
                                mP_Connection.RollbackTrans()
                                Exit Sub
                            End If
                        End If


                        'Added by praveen on 19.12.2017 for adding entry in Stock ledger for FTS items in Hilex
                        Dim oCmd = New ADODB.Command
                        With oCmd
                            .ActiveConnection = mP_Connection
                            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                            .CommandText = "USP_FTS_PRODUCTIONSLIP_CANCELLATION"
                            .Parameters.Append(.CreateParameter("@Unit_Code", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrUNITID))
                            .Parameters.Append(.CreateParameter("@Doc_No", ADODB.DataTypeEnum.adBigInt, ADODB.ParameterDirectionEnum.adParamInput, , Val(txtChallanNo.Text.Trim)))
                            .Parameters.Append(.CreateParameter("@ERRMSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 1000))
                            .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End With
                        If oCmd.Parameters(oCmd.Parameters.Count - 1).Value <> "" Then
                            mP_Connection.RollbackTrans()
                            MsgBox(oCmd.Parameters(oCmd.Parameters.Count - 1).Value, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                            oCmd = Nothing
                            Exit Sub
                        End If
                        oCmd = Nothing
                        'Code End
                    End If
                    'Code End


                    If UpdateFinance() Then
                        '101631219
                        Dim objCmd As New ADODB.Command

                        With objCmd
                            .ActiveConnection = mP_Connection
                            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                            .CommandText = "USP_BSR_INVOICE_UPDATION"
                            .CommandTimeout = 0
                            .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                            .Parameters.Append(.CreateParameter("@TEMP_INV_NO", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, , 0))
                            .Parameters.Append(.CreateParameter("@INV_NO", ADODB.DataTypeEnum.adBigInt, ADODB.ParameterDirectionEnum.adParamInput, , Val(txtChallanNo.Text.Trim)))
                            .Parameters.Append(.CreateParameter("@USER_ID", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, mP_User))
                            .Parameters.Append(.CreateParameter("@OPERATION_TYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, "CANCEL"))
                            .Parameters.Append(.CreateParameter("@MESSAGE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInputOutput, 8000, ""))
                            .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End With

                        If objCmd.Parameters(objCmd.Parameters.Count - 1).Value.ToString().Trim.Length <> 0 Then
                            MessageBox.Show(objCmd.Parameters(objCmd.Parameters.Count - 1).Value.ToString(), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
                            objCmd = Nothing
                            mP_Connection.RollbackTrans()
                            'Exit Sub
                        End If
                        objCmd = Nothing
                        '101631219

                        If mblnRejTracking = True And UCase(mstrInvType) = "REJ" Then
                            mP_Connection.Execute("Update MKT_INVREJ_DTL Set Cancel_flag=1,Upd_DT=Getdate(), Upd_UserID='" & mP_User & "'" & " where UNIT_CODE='" & gstrUNITID & "' AND Invoice_No=" & Trim(txtChallanNo.Text), , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If

                        'Code Added By Shubhra To Remove Cancelled Invoice from ILVS
                        'Begin
                        strRemoveInvFromLoadingSlip = "Update Loadingslip set InvoiceNo = NULL, ACT_INV_NO = NULL" &
                            " where Unit_Code = '" & gstrUNITID & "' and ACT_INV_NO = " & Val(txtChallanNo.Text.Trim) & ""
                        mP_Connection.Execute(strRemoveInvFromLoadingSlip, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        'End

                        mP_Connection.CommitTrans()
                        MsgBox("Challan Cancelled Successfully", MsgBoxStyle.Information, "empower")
                    Else
                        mP_Connection.RollbackTrans()
                        Exit Sub
                    End If
                    Call EnableControls(False, Me, True)
                    txtLocationCode.Enabled = True
                    txtLocationCode.BackColor = System.Drawing.Color.White
                    CmdGrpChEnt.Enabled(0) = False
                    CmdGrpChEnt.Enabled(1) = False
                    CmdLocCodeHelp.Enabled = True
                    txtChallanNo.Enabled = True
                    txtChallanNo.BackColor = System.Drawing.Color.White
                    CmdChallanNo.Enabled = True
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_REFRESH 'REFRESH FORM DATA
                Call frmMKTTRN0025_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        mP_Connection.RollbackTrans()
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub


    Public Function CHECK_FTSINVOICE(ByVal strCondition As String) As Boolean
        '----------------------------------------------------
        'Created By     -  Prashant Rajpal
        'Description    -  TO CHECK WHETHER THE INVOICE IS OF FTS ITEM 
        '------------------------------------------------------------------
        On Error GoTo ErrHandler

        Dim strsql As String

        CHECK_FTSINVOICE = False
        strsql = "SELECT DBO.UFN_FTS_CHECK_INVOICECANCELLATION('" & gstrUNITID & "','" & txtChallanNo.Text.Trim & "','" & strCondition & "') "
        CHECK_FTSINVOICE = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strsql))

        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Function DispatchQuantityFromMonthlySchedule(ByVal pstrAccountCode As String, ByVal pstrCustomerDrawNo As String, ByVal pstrItemCode As String, ByVal pstrDate As String, ByVal pstrMode As String, ByVal pdblPrevQty As Double) As Boolean
        'Commented for Issue ID eMpro-20080930-22159 Starts
        '        '---------------------------------------------------------------------------------------
        '		'Name       :   GetTotalDispatchQuantityFromDailySchedule
        '		'Type       :   Function
        '		'Author     :   Tapan Jain (Modified By     -   Nitin Sood)
        '		'Arguments  :
        '		'Return     :   TRUE    -   If Schedule is Monthly
        '		'Purpose    :
        '		'---------------------------------------------------------------------------------------
        '		Dim strScheduleSql As String
        '		Dim objRsForSchedule As New ADODB.Recordset
        '		Dim ldblTotalDispatchQuantity As Double
        '		Dim ldblTotalScheduleQuantity As Double
        '		Dim lintLoopCounter As Short
        '		Dim strMakeDate As String
        '		On Error GoTo ErrHandler
        '		ldblTotalDispatchQuantity = 0
        '		ldblTotalScheduleQuantity = 0
        '        mP_Connection.Execute("SET DATEFORMAT 'mdy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        '		If Val(CStr(Month(CDate(pstrDate)))) < 10 Then
        '			strMakeDate = Year(CDate(pstrDate)) & "0" & Month(CDate(pstrDate))
        '		Else
        '			strMakeDate = Year(CDate(pstrDate)) & Month(CDate(pstrDate))
        '		End If
        '		If objRsForSchedule.State = 1 Then objRsForSchedule.Close()
        '		objRsForSchedule.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        '		strScheduleSql = "Select Schedule_Qty,isnull(Despatch_Qty,0) AS Despatch_qty  from MonthlyMktSchedule where Account_Code='" & pstrAccountCode & "' and "
        '		strScheduleSql = strScheduleSql & " Year_Month =" & Val(Trim(strMakeDate)) & ""
        '		strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND Item_code = '" & pstrItemCode & "' and status =1 "
        '		objRsForSchedule.Open(strScheduleSql, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        '		If objRsForSchedule.EOF Or objRsForSchedule.BOF Then
        '			DispatchQuantityFromMonthlySchedule = False
        '            mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        '			objRsForSchedule.Close()
        '			Exit Function
        '		Else
        '			objRsForSchedule.MoveFirst()
        '			For lintLoopCounter = 1 To objRsForSchedule.RecordCount
        '				'Get Total Schedule Qty , Dispatch Qty
        '				ldblTotalScheduleQuantity = ldblTotalScheduleQuantity + Val(objRsForSchedule.Fields("Schedule_Qty").Value)
        '				ldblTotalDispatchQuantity = ldblTotalDispatchQuantity + Val(objRsForSchedule.Fields("Despatch_Qty").Value)
        '				objRsForSchedule.MoveNext()
        '			Next 
        '			DispatchQuantityFromMonthlySchedule = True 'Dispatch Against Monthly Schedule
        '            mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        '			objRsForSchedule.Close()
        '			Exit Function
        '		End If
        '		Exit Function 'This is to avoid the execution of the error handler
        'ErrHandler: 
        '        mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        '		DispatchQuantityFromMonthlySchedule = False
        '		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection)
        'Commented for Issue ID eMpro-20080930-22159 Ends
    End Function
    Private Function ReviseSchedule() As Boolean
        'Commented for Issue ID eMpro-20080930-22159 Starts
        '        '****************************************************
        '		'Created By     -  Nitin Sood
        '		'Description    -  Decrease Dispatch Quantity for a Particular Item
        '		'                   from Marketing Schedule (Daily / Monthly)
        '		'
        '		'****************************************************
        '		Dim dblNetDispatchQty As Double 'Net Dispatch Qty
        '		Dim intRowCount As Short 'For Next Loop . . .
        '		Dim strDrgNo As String 'Cust DRWG no.
        '		Dim strItemCode As String 'Item Code
        '		Dim dblQty As Double 'Qty
        '		Dim mstrUpdDispatchSql As String 'SQL
        '		Dim strMakeDate As String 'To Make Date in Case of MonthlySchedule
        '		On Error GoTo ErrHandler
        '		ReviseSchedule = True
        '		With SpChEntry
        '			For intRowCount = 1 To .maxRows 'For All Items In Spread Revise Schedule
        '				.Col = 1 : .Col2 = 1 : .Row = intRowCount : .Row2 = intRowCount : strItemCode = Trim(.Text)
        '				.Col = 2 : .Col2 = 2 : .Row = intRowCount : .Row2 = intRowCount : strDrgNo = Trim(.Text)
        '				.Col = 5 : .Col2 = 5 : .Row = intRowCount : .Row2 = intRowCount : dblQty = Val(.Text)
        '				'Schedule May Be Daily or Monthly
        '				If DispatchQuantityFromDailySchedule(Trim(txtCustCode.Text), strDrgNo, strItemCode, Trim(lblDateDes.Text), "ADD", 0) Then
        '					'Dispatch was made from Daily Schedule,Reduce Dispatch Qty from this table
        '					mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
        '					mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) - (" & Val(CStr(dblQty)) & "),Status = 1 "
        '					mstrUpdDispatchSql = mstrUpdDispatchSql & " Where Account_Code='" & Trim(txtCustCode.Text) & "' and "
        '					mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(CDate(Trim(lblDateDes.Text))) & "'"
        '					mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(CDate(Trim(lblDateDes.Text))) & "'"
        '					mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(CDate(Trim(lblDateDes.Text))) & "'"
        '					mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & strDrgNo & "'and Item_code = '" & strItemCode & "'" & vbCrLf
        '				ElseIf DispatchQuantityFromMonthlySchedule(Trim(txtCustCode.Text), strDrgNo, strItemCode, Trim(lblDateDes.Text), "ADD", 0) Then 
        '					'Dispatch was made against Monthly Schedule,Subtract From Dispatch and Update Status Back to 1
        '					'Make date 1st
        '					If Val(CStr(Month(CDate(lblDateDes.Text)))) < 10 Then
        '						strMakeDate = Year(CDate(lblDateDes.Text)) & "0" & Month(CDate(lblDateDes.Text))
        '					Else
        '						strMakeDate = Year(CDate(lblDateDes.Text)) & Month(CDate(lblDateDes.Text))
        '					End If
        '					mstrUpdDispatchSql = mstrUpdDispatchSql & "Update MonthlyMktSchedule set Despatch_qty ="
        '					mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) - (" & Val(CStr(dblQty)) & ") , Status = 1 "
        '					mstrUpdDispatchSql = mstrUpdDispatchSql & " Where Account_Code='" & Trim(txtCustCode.Text) & "' and "
        '					mstrUpdDispatchSql = mstrUpdDispatchSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
        '					mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(strDrgNo) & "' and Item_code = '" & strItemCode & "'" & vbCrLf
        '				End If
        '			Next 
        '            If Len(mstrUpdDispatchSql) > 0 Then mP_Connection.Execute(mstrUpdDispatchSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords) 'Execute Revising Schedule Script
        '		End With
        '		Exit Function
        'ErrHandler: 'The Error Handling Code Starts here
        '		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        '		ReviseSchedule = False
        'Commented for Issue ID eMpro-20080930-22159 Ends
    End Function
    Private Function DispatchQuantityFromDailySchedule(ByVal pstrAccountCode As String, ByVal pstrCustomerDrawNo As String, ByVal pstrItemCode As String, ByVal pstrDate As String, ByVal pstrMode As String, ByVal pdblPrevQty As Double) As Boolean
        'Commented for Issue ID eMpro-20080930-22159 Starts
        '---------------------------------------------------------------------------------------
        'Name       :   DispatchQuantityFromDailySchedule
        'Type       :   Function
        'Author     :   Tapan Jain (Modified By     -   Nitin Sood)
        'Arguments  :
        'Return     :   TRUE - Despatch was done against DailyMKTSchedule
        'Purpose    :   To Find out ,If Despatch was againt Monthly Schedule for Daily Schedule
        '---------------------------------------------------------------------------------------
        '		On Error GoTo ErrHandler
        '		Dim strScheduleSql As String
        '		Dim objRsForSchedule As New ADODB.Recordset
        '		Dim ldblTotalDispatchQuantity As Double
        '		Dim ldblTotalScheduleQuantity As Double
        '		Dim lintLoopCounter As Short
        '		ldblTotalDispatchQuantity = 0 'INITIALIZE
        '		ldblTotalScheduleQuantity = 0
        '        mP_Connection.Execute("SET DATEFORMAT 'mdy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        '		strScheduleSql = "Select Schedule_Quantity,Despatch_Qty from DailyMktSchedule where Account_Code='" & pstrAccountCode & "' and "
        '		strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(CDate(pstrDate)) & "'"
        '		strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date)='" & Month(CDate(pstrDate)) & "'"
        '		strScheduleSql = strScheduleSql & " and Trans_Date <='" & VB6.Format(pstrDate, "mm/dd/yyyy") & "'"
        '		strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND ITEM_CODE  = '" & pstrItemCode & "' and Status =1  ORDER BY Trans_Date DESC" '''and Schedule_Flag =1   ( Now Not Consider)
        '		If objRsForSchedule.State = 1 Then objRsForSchedule.Close()
        '		objRsForSchedule.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        '		objRsForSchedule.Open(strScheduleSql, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        '		If objRsForSchedule.EOF Or objRsForSchedule.BOF Then
        '			DispatchQuantityFromDailySchedule = False
        '            mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        '			objRsForSchedule.Close()
        '			Exit Function
        '		Else
        '			objRsForSchedule.MoveFirst()
        '			For lintLoopCounter = 1 To objRsForSchedule.RecordCount
        '				ldblTotalScheduleQuantity = ldblTotalScheduleQuantity + Val(objRsForSchedule.Fields("Schedule_Quantity").Value)
        '				ldblTotalDispatchQuantity = ldblTotalDispatchQuantity + Val(objRsForSchedule.Fields("Despatch_Qty").Value)
        '				objRsForSchedule.MoveNext()
        '			Next 
        '			If ldblTotalScheduleQuantity > 0 And ldblTotalDispatchQuantity > 0 Then
        '				DispatchQuantityFromDailySchedule = True
        '			End If
        '            mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        '			objRsForSchedule.Close()
        '			Exit Function
        '		End If
        '		Exit Function 'This is to avoid the execution of the error handler
        'ErrHandler: 
        '		DispatchQuantityFromDailySchedule = False
        '        mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        '		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection)
        'Commented for Issue ID eMpro-20080930-22159 Ends
    End Function
    Private Sub Cmditems_Click()
        '****************************************************
        'Created By     -  Nisha
        'Description    -  Display Another Form for User To Select Item Code >From CustOrd_Dtl
        '                  And After Selecting Item Code Select Data From Sales_Dtl and Display
        '                  That Details In The Spread
        '****************************************************
        On Error GoTo ErrHandler
        Dim salechallan As String
        Dim strItemNotIn As String
        Dim varItemCode As Object
        Dim strStockLocation As String
        Dim rsCurrencyType As ClsResultSetDB
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        With Me.SpChEntry
            .MaxRows = 1
            .Row = 1 : .Row2 = .MaxRows : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
        End With
        strStockLocation = StockLocationSalesConf(mstrInvType, mstrInvSubType, "TYPE")
        If Len(Trim(strStockLocation)) > 0 Then
            'CHANGED ON 15/07/2002 FOR EXPORT INVOICE
            If (UCase(mstrInvType)) = "INV" Or (UCase(mstrInvType)) = "EXP" Then
                mstrItemCode = SelectItemfromsaleDtl(Trim(txtChallanNo.Text))
                If Len(Trim(mstrItemCode)) = 0 Then SpChEntry.MaxRows = 0 : frmMKTTRN0021.Close()
            Else
                mstrItemCode = SelectItemfromsaleDtl(Trim(txtChallanNo.Text))
                If Len(Trim(mstrItemCode)) = 0 Then SpChEntry.MaxRows = 0 : frmMKTTRN0021.Close()
            End If
        Else
            MsgBox("Please Define Stock Location in Sales Conf", MsgBoxStyle.Information, "empower")
            Exit Sub
        End If
        Dim intDecimalPlace As Short
        Dim strCurrency As String
        If Len(mstrItemCode) > 0 Then
            mstrItemCode = Mid(mstrItemCode, 1, Len(mstrItemCode) - 1)
            '*************** to get refrence detail for curenct
            rsCurrencyType = New ClsResultSetDB
            rsCurrencyType.GetResult("Select Currency_code from saleschallan_dtl where UNIT_CODE='" & gstrUNITID & "' AND doc_No = " & Val(txtChallanNo.Text))
            If rsCurrencyType.GetNoRows > 0 Then
                rsCurrencyType.MoveFirst()
                strCurrency = rsCurrencyType.GetValue("Currency_code")
            End If
            rsCurrencyType.ResultSetClose()
            rsCurrencyType = Nothing
            intDecimalPlace = ToGetDecimalPlaces(strCurrency)
            If intDecimalPlace < 2 Then
                intDecimalPlace = 2
            End If
            DisplayDetailsInSpread(strCurrency) 'Procedure Call To Select Data >From Sales_Dtl
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdLocCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdLocCodeHelp.Click
        '****************************************************
        'Created By     -  Nisha
        'Description    -  To Display Help From Location Master
        '****************************************************
        On Error GoTo ErrHandler
        Dim strHelp As String
        If Len(Me.txtLocationCode.Text) = 0 Then 'To check if There is No Text Then Show All Help
            strHelp = ShowList(1, (txtLocationCode.MaxLength), "", "s.Location_Code", "l.Description", "Location_mst l,SaleConf s", "and s.Location_Code=l.Location_Code AND s.UNIT_CODE=l.UNIT_CODE", , , , , , "s.UNIT_CODE")
            If strHelp = "-1" Then 'If No Record Exists In The Table
                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Exit Sub
            Else
                txtLocationCode.Text = strHelp
            End If
        Else
            'To Display All Possible Help Starting With Text in TextField
            strHelp = ShowList(1, (txtLocationCode.MaxLength), txtLocationCode.Text, "s.Location_Code", "l.Description", "Location_mst l,SaleConf s", "and s.Location_Code=l.Location_Code AND s.UNIT_CODE=l.UNIT_CODE", , , , , , "s.UNIT_CODE")
            If strHelp = "-1" Then 'If No Record Exists In The Table
                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Exit Sub
            Else
                txtLocationCode.Text = strHelp
            End If
        End If
        'Procedure Call To Select The Location Code Description
        Call SelectDescriptionForField("Description", "Location_Code", "Location_Mst", lblLocCodeDes, (txtLocationCode.Text))
        Call txtLocationCode_Validating(txtLocationCode, New System.ComponentModel.CancelEventArgs(False))
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub ctlFormHeader1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        On Error GoTo ErrHandler
        Call ShowHelp("HLPMKTTRN0005.HTM")
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0025_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Err_Handler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlFormHeader1_ClickEvent(ctlFormHeader1, New System.EventArgs())
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0025_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                'If user press the ESC Key ,the Form will be in View Mode
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
    Private Sub frmMKTTRN0025_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '-----Checks the Schedule Flag whether DS wise or not.
        Dim RsobjSchedules As ADODB.Recordset
        Dim RsobjINVOICECANCELLATION_CURRENTDATE As ADODB.Recordset
        On Error GoTo ErrHandler
        RsobjSchedules = New ADODB.Recordset
        RsobjINVOICECANCELLATION_CURRENTDATE = New ADODB.Recordset
        'Add Form Name To Window List
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        'Fill Lebels From Resource File
        Call FitToClient(Me, FraChEnt, ctlFormHeader1, CmdGrpChEnt, 500)
        'Change Picture on Button - Nitin
        CmdGrpChEnt.Caption(0) = "Cancel"
        CmdGrpChEnt.Caption(1) = "Refresh"
        CmdGrpChEnt.Enabled(1) = False
        '***** rr00rr 
        CmdGrpChEnt.Picture(UCActXCtl.clsDeclares.ButtonEnabledEnum.NEW_BUTTON) = My.Resources.resEmpower.ico123.ToBitmap
        '*********** rr00rr
        'Set Help Pictures At Command Button
        CmdLocCodeHelp.Image = My.Resources.ico111.ToBitmap
        CmdChallanNo.Image = My.Resources.ico111.ToBitmap
        'Check If Company is 100% EOU then CVD SVD fields are SHOWN - NITIN SOOD
        gobjDB = New ClsResultSetDB
        If gobjDB.GetResult("Select EOU_FLAG From Company_Mst WHERE UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic) Then
            If gobjDB.GetNoRows > 0 Then
                blnEOU_FLAG = gobjDB.GetValue("EOU_FLAG")
            End If
        End If
        gobjDB.ResultSetClose()
        gobjDB = Nothing
        'Initially Disable All Controls
        Call EnableControls(False, Me, True)
        lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        'Get Server Date
        lblDateDes.Text = VB6.Format(GetServerDate(), gstrDateFormat)
        'Date is Also Added in DatePicker,and Its Visible Property is set to False - Nitin Sood
        'Add Transport Type To Combo
        Call AddTransPortTypeToCombo()
        txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        CmdLocCodeHelp.Enabled = True : CmdChallanNo.Enabled = True
        Me.SpChEntry.Enabled = False
        'Set Column Headers
        With Me.SpChEntry
            .DisplayRowHeaders = True
            .set_ColWidth(0, 300)
            .MaxCols = 15
            .Row = 0 : .Col = 1 : .Text = "Internal Part No."
            .Row = 0 : .Col = 2 : .Text = "Cust.Part No."
            '***
            .Row = 0 : .Col = 3 : .Text = "Rate"
            .Row = 0 : .Col = 4 : .Text = "Cust Material"
            .Row = 0 : .Col = 5 : .Text = "Quantity"
            .Row = 0 : .Col = 6 : .Text = "Packing(%)"
            .Row = 0 : .Col = 7 : .Text = "EXC(%)"
            '.EditMode = False
            .Row = 0 : .Col = 8 : .Text = "CVD(%)"
            .Row = 0 : .Col = 9 : .Text = "SAD(%)"
            If Not blnEOU_FLAG Then
                .Col = 8 : .Col2 = 8
                .BlockMode = True
                .ColHidden = True
                .BlockMode = False
                .Col = 9 : .Col2 = 9
                .BlockMode = True
                .ColHidden = True
                .BlockMode = False
            End If
            .Row = 0 : .Col = 10 : .Text = "Others"
            .Row = 0 : .Col = 11 : .Text = "From Box"
            .Row = 0 : .Col = 12 : .Text = "To Box"
            .Row = 0 : .Col = 13 : .Text = "Cumulative Boxes" : .set_ColWidth(10, 1500)
            .Row = 0 : .Col = 14 : .Text = "Delete"
            .Col = 14 : .Col2 = 14 : .BlockMode = True : .ColHidden = True : .BlockMode = False
            .Row = 0 : .Col = 15 : .Text = "Tool Cost"
            .Col = 15 : .Col2 = 15 : .BlockMode = True : .ColHidden = True : .BlockMode = False
        End With
        'Add Row
        Call addRowAtEnterKeyPress(1)
        lblRGPDes.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        '----------Checks the flag whether Schedule is DSwise or old procedure
        If RsobjSchedules.State = ADODB.ObjectStateEnum.adStateOpen Then RsobjSchedules.Close()
        RsobjSchedules.Open("SELECT isnull(DSWiseTracking,0) FROM sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not RsobjSchedules.EOF Then
            If IIf(RsobjSchedules.Fields(0).Value, 1, 0) = 1 Then
                MBlnDsTracking = True
            Else
                MBlnDsTracking = False
            End If
        End If
        If RsobjSchedules.State = ADODB.ObjectStateEnum.adStateOpen Then RsobjSchedules.Close()
        RsobjSchedules.Open("Select REJINV_Tracking from sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not RsobjSchedules.EOF Then
            If RsobjSchedules.Fields("RejINV_Tracking").Value = True Then
                mblnRejTracking = True
            Else
                mblnRejTracking = False
            End If
        End If
        If RsobjSchedules.State = ADODB.ObjectStateEnum.adStateOpen Then RsobjSchedules.Close()
        RsobjSchedules = Nothing
        If RsobjINVOICECANCELLATION_CURRENTDATE.State = ADODB.ObjectStateEnum.adStateOpen Then RsobjINVOICECANCELLATION_CURRENTDATE.Close()
        RsobjINVOICECANCELLATION_CURRENTDATE.Open("SELECT INVOICECANCELLATION_CURRENTDATE_REQ  FROM sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not RsobjINVOICECANCELLATION_CURRENTDATE.EOF Then
            If IIf(RsobjINVOICECANCELLATION_CURRENTDATE.Fields(0).Value, 1, 0) = 1 Then
                mblnINVOICECANCELLATION_CURRENTDATE = True
            Else
                mblnINVOICECANCELLATION_CURRENTDATE = False
            End If
        End If
        If RsobjINVOICECANCELLATION_CURRENTDATE.State = ADODB.ObjectStateEnum.adStateOpen Then RsobjINVOICECANCELLATION_CURRENTDATE.Close()
        RsobjINVOICECANCELLATION_CURRENTDATE = Nothing
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0025_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        CmdGrpChEnt.Enabled(1) = False
        mdifrmMain.CheckFormName = mintIndex
        txtLocationCode.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0025_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0025_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        On Error GoTo ErrHandler
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
            Me.Dispose()
        End If
        'Checking The Status
        If gblnCancelUnload = True Then eventArgs.Cancel = 1
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTTRN0025_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub addRowAtEnterKeyPress(ByRef pintRows As Short)
        '****************************************************
        'Created By     -  Nisha
        'Description    -  Add Row At Enter Key Press Of Last Column Of Spread
        '****************************************************
        On Error GoTo ErrHandler
        Dim intRowHeight As Short
        With Me.SpChEntry
            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            For intRowHeight = 1 To pintRows
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .set_RowHeight(.Row, 300)
            Next intRowHeight
            If .MaxRows > 4 Then .ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsBoth
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub SelectDescriptionForField(ByRef pstrFieldName1 As String, ByRef pstrFieldName2 As String, ByRef pstrTableName As String, ByRef pContrName As System.Windows.Forms.Control, ByRef pstrControlText As String)
        '****************************************************
        'Created By     -  Nisha
        'Description    -  To Select The Field Description In The Description Labels
        'Arguments      -  pstrFieldName1 - Field Name1,pstrFieldName2 - Field Name2,pstrTableName - Table Name
        '               -  pContName - Name Of The Control where Caption Is To Be Set
        '               -  pstrControlText - Field Text
        '****************************************************
        On Error GoTo ErrHandler
        Dim strDesSql As String 'Declared to make Select Query
        Dim rsDescription As ClsResultSetDB
        strDesSql = "Select " & Trim(pstrFieldName1) & " from " & Trim(pstrTableName) & " where " & Trim(pstrFieldName2) & "='" & Trim(pstrControlText) & "' and UNIT_CODE='" & gstrUNITID & "'"
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
    Private Sub lblCurrencyDes_Change()
        If Trim(lblCurrencyDes.Text) <> "" Then
            If Trim(lblCurrencyDes.Text) = Trim(gstrCURRENCYCODE) Then
                lblExchangeRateValue.Text = CStr(1.0#)
            Else
                If UCase(Trim(mstrInvType)) = "INV" Or UCase(Trim(mstrInvType)) = "SMP" Or UCase(Trim(mstrInvType)) = "TRF" Or UCase(Trim(mstrInvType)) = "JOB" Or UCase(Trim(mstrInvType)) = "EXP" Then
                    lblExchangeRateValue.Text = CStr(GetExchangeRate(lblCurrencyDes.Text, lblDateDes.Text, True))
                Else
                    lblExchangeRateValue.Text = CStr(GetExchangeRate(lblCurrencyDes.Text, lblDateDes.Text, False))
                End If
                If Val(Trim(lblExchangeRateValue.Text)) = 1 Then
                    MsgBox("Exchange Rate for " & Trim(lblCurrencyDes.Text) & " is not defined on " & lblDateDes.Text, MsgBoxStyle.Information, "empower")
                    lblExchangeRateValue.Text = ""
                End If
            End If
        Else
            lblExchangeRateValue.Text = ""
        End If
    End Sub
    Private Sub txtChallanNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtChallanNo.TextChanged
        On Error GoTo ErrHandler
        If Len(Trim(txtChallanNo.Text)) = 0 Then
            Call RefreshForm("CHALLAN")
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtChallanNo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtChallanNo.Enter
        On Error GoTo ErrHandler
        Me.txtChallanNo.SelectionStart = 0
        Me.txtChallanNo.SelectionLength = Len(Me.txtChallanNo.Text)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtChallanNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtChallanNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                If Len(txtChallanNo.Text) > 0 Then
                    Call txtChallanNo_Validating(txtChallanNo, New System.ComponentModel.CancelEventArgs(False))
                Else
                    Me.CmdGrpChEnt.Focus()
                End If
        End Select
        'Allow only Numbers
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
        '****************************************************
        'Created By     -  Nisha
        'Description    -  If F1 Key Press Then Display Help From SalesChallan_Dtl
        '****************************************************
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
        '****************************************************
        'Created By     -  Nisha
        'Description    -  Check Validity Of Challan No. In SalesChallan_Dtl
        '****************************************************
        Dim strCondition As String
        On Error GoTo ErrHandler
        If Len(txtChallanNo.Text) > 0 Then
            'Check Existance Of Doc No In The SalesChallan_Dtl
            If CheckExistanceOfFieldData((txtChallanNo.Text), "Doc_No", "SalesChallan_Dtl", "UNIT_CODE='" & gstrUNITID & "'") Then
                'If Challan No. Exists
                'Get Data From Challan_Dtl,Cust_Ord_Dtl,Sales_Dtl
                If Len(txtLocationCode.Text) > 0 Then
                    If GetDataInViewMode() Then 'if record found
                        'Added for Issue ID 19992 Starts
                        Call CheckMultipleSOAllowed(lblInvoiceType.Text, lblInvoiceSubType.Text)
                        If mblnMultipleSOAllowed = True Then
                            With SpChEntry
                                .MaxCols = 17
                                .Row = 0 : .Col = 16 : .Text = "Reference No." : .set_ColWidth(16, 1600) : .ColHidden = False
                                .Row = 0 : .Col = 17 : .Text = "Amendement No." : .set_ColWidth(17, 1600) : .ColHidden = False
                                txtRefNo.Text = ""
                            End With
                        Else
                            With SpChEntry
                                .MaxCols = 17
                                .Row = 0 : .Col = 16 : .Text = "Reference No." : .set_ColWidth(16, 1600) : .ColHidden = True
                                .Row = 0 : .Col = 17 : .Text = "Amendement No." : .set_ColWidth(17, 1600) : .ColHidden = True
                            End With
                        End If
                        Call Cmditems_Click()
                        With CmdGrpChEnt
                            .Enabled(0) = True 'CANCEL
                            '.Enabled(1) = True  'REFRESH
                        End With
                        With txtremark 'Enable Cancellation Remarks
                            .Enabled = True : .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            .Focus()
                        End With
                    Else 'if no record found then display message
                        Call ConfirmWindow(10414, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
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
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Function CheckExistanceOfFieldData(ByRef pstrFieldText As String, ByRef pstrColumnName As String, ByRef pstrTableName As String, Optional ByRef pstrCondition As String = "") As Boolean
        '****************************************************
        'Created By     -  Nisha
        'Description    -  To Check Validity Of Field Data Whethet it Exists In The
        '                  Database Or Not
        'Arguments      -  pstrFieldText - Field Text,pstrColumnName - Column Name
        '               -  pstrTableName - Table Name,pstrCondition - Optional Parameter For Condition
        '****************************************************
        On Error GoTo ErrHandler
        CheckExistanceOfFieldData = False
        Dim strTableSql As String 'Declared To Make Select Query
        Dim rsExistData As ClsResultSetDB
        Dim strInvoiceConditionDate As String
        If Len(Trim(pstrCondition)) > 0 Then
            strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and " & pstrCondition
        Else
            strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' AND cancel_flag=0  "
        End If
        If mblnINVOICECANCELLATION_CURRENTDATE = True Then
            strTableSql = strTableSql + " AND INVOICE_DATE =CONVERT(char(12), GETDATE(), 106)"
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
        '****************************************************
        'Created By     -  Nisha
        'Modified By    -  Nitin Sood
        'Modifcation    -  Amendment Number Displayed
        'Description    -  To display data in view mode from SalasChallan_Dtl,Sales_Dtl acc.to
        'LocationCode & Challan_No.
        '****************************************************
        On Error GoTo ErrHandler
        GetDataInViewMode = False
        Dim strGetData As String
        Dim rsGetData As ClsResultSetDB
        Dim rsCustMst As ClsResultSetDB
        Dim strSalesChallanDtl As String
        Dim strRGPNOs As String
        Dim strCustMst As String
        strSalesChallanDtl = "SELECT Transport_type,Vehicle_No,Account_Code,Cust_ref,Amendment_No,SalesTax_Type,"
        strSalesChallanDtl = strSalesChallanDtl & "Insurance,Invoice_Date,Payment_terms,"
        strSalesChallanDtl = strSalesChallanDtl & "Invoice_Type,Sub_Category,Cust_Name,Carriage_Name,Frieght_Amount, "
        strSalesChallanDtl = strSalesChallanDtl & "Surcharge_salesTaxType,Amendment_No,ref_doc_no,Currency_Code,Exchange_Rate,Remarks,PerValue From Saleschallan_dtl WHERE UNIT_CODE='" & gstrUNITID & "' and Location_Code ='"
        strSalesChallanDtl = strSalesChallanDtl & Trim(txtLocationCode.Text) & "' and Doc_No = " & Val(txtChallanNo.Text)
        rsGetData = New ClsResultSetDB
        rsGetData.GetResult(strSalesChallanDtl, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsGetData.GetNoRows > 0 Then
            GetDataInViewMode = True
            txtCustCode.Text = rsGetData.GetValue("Account_Code")
            txtRefNo.Text = rsGetData.GetValue("Cust_ref")
            txtAmendNo.Text = rsGetData.GetValue("Amendment_No")
            txtCarrServices.Text = rsGetData.GetValue("Carriage_Name")
            ctlInsurance.Text = rsGetData.GetValue("Insurance")
            txtFreight.Text = rsGetData.GetValue("Frieght_Amount")
            txtSaleTaxType.Text = rsGetData.GetValue("SalesTax_Type")
            lblSaltax_Per.Text = SelectDataFromTable("TxRt_Percentage", "Gen_Taxrate", " WHERE UNIT_CODE='" & gstrUNITID & "' and TxRt_Rate_No='" & Trim(txtSaleTaxType.Text) & "'")
            txtSurchargeTaxType.Text = rsGetData.GetValue("Surcharge_salesTaxType")
            lblSurcharge_Per.Text = SelectDataFromTable("TxRt_Percentage", "Gen_Taxrate", " WHERE UNIT_CODE='" & gstrUNITID & "' and TxRt_Rate_No='" & Trim(txtSurchargeTaxType.Text) & "'")
            strRGPNOs = rsGetData.GetValue("ref_doc_no")
            strRGPNOs = Replace(strRGPNOs, "§", ", ", 1)
            lblRGPDes.Text = strRGPNOs
            lblCustCodeDes.Text = rsGetData.GetValue("Cust_Name")
            lblDateDes.Text = VB6.Format(rsGetData.GetValue("Invoice_Date"), gstrDateFormat)
            mstrInvType = rsGetData.GetValue("Invoice_Type")
            lblInvoiceType.Text = SelectDataFromTable("Description", "Saleconf", " WHERE UNIT_CODE='" & gstrUNITID & "' and Invoice_type='" & Trim(mstrInvType) & "'")
            mstrInvSubType = rsGetData.GetValue("Sub_Category")
            ''Commented by priti on 23 Jan 2026 to impact barcode issue in Raw material invoice
            'lblInvoiceSubType.Text = SelectDataFromTable("Sub_Type_Description", "Saleconf", " WHERE UNIT_CODE='" & gstrUNITID & "' and Sub_type='" & Trim(mstrInvSubType) & "'")
            lblInvoiceSubType.Text = SelectDataFromTable("Sub_Type_Description", "Saleconf", " WHERE UNIT_CODE='" & gstrUNITID & "' and Invoice_type='" & Trim(mstrInvType) & "' and Sub_type='" & Trim(mstrInvSubType) & "'")
            lblCurrency.Visible = True : lblCurrencyDes.Visible = True
            lblCurrencyDes.Text = rsGetData.GetValue("Currency_code")
            lblExchangeRateLable.Visible = True : lblExchangeRateValue.Visible = True
            '*******************
            lblCreditTerm.Text = IIf(IsDBNull(rsGetData.GetValue("payment_terms")), "", rsGetData.GetValue("payment_terms"))
            If Len(Trim(lblCreditTerm.Text)) > 0 Then
                Call SelectDescriptionForField("crTrm_desc", "crtrm_termID", "Gen_CreditTrmMaster", lblCreditTermDesc, Trim(lblCreditTerm.Text))
            Else
                lblCreditTermDesc.Text = ""
            End If
            ''Have to Add Remarks & Per Value
        Else
            GetDataInViewMode = False
        End If
        '***To Display invoice Address of Customer
        If Len(Trim(txtCustCode.Text)) > 0 Then
            rsCustMst = New ClsResultSetDB
            strCustMst = "SELECT Bill_Address1 + ', '  +  Bill_Address2 + ', ' + Bill_City + ' - ' + Bill_Pin as  invoiceAddress from Customer_Mst where UNIT_CODE='" & gstrUNITID & "' and Customer_code ='" & txtCustCode.Text & "'"
            rsCustMst.GetResult(strCustMst)
            If rsCustMst.GetNoRows > 0 Then
                lblAddressDes.Text = rsCustMst.GetValue("InvoiceAddress")
            End If
            rsCustMst.ResultSetClose()
            rsCustMst = Nothing
        End If
        '***
        rsGetData.ResultSetClose()
        rsGetData = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function DisplayDetailsInSpread(ByRef pstrCurrency As String) As Boolean
        '****************************************************
        'Created By     -  Nisha
        'Description    -  To display Details From Sales_Dtl Acc To Location Code,Challan No and Drawing No
        '****************************************************
        On Error GoTo ErrHandler
        Dim intLoopCounter As Short
        Dim intRecordCount As Short
        Dim inti As Short
        Dim strsaledtl As String
        Dim dblPacking As Double
        Dim varItem_Code As Object
        Dim varItemAlready As Object
        Dim rsSalesDtl As ClsResultSetDB
        Dim intDecimal As Short
        strsaledtl = ""
        strsaledtl = "SELECT * from Sales_Dtl WHERE UNIT_CODE='" & gstrUNITID & "' and Location_Code='" & Trim(txtLocationCode.Text) & "'"
        strsaledtl = strsaledtl & " and Doc_No=" & Val(txtChallanNo.Text) & " and Cust_Item_Code in(" & Trim(mstrItemCode) & ")"
        rsSalesDtl = New ClsResultSetDB
        rsSalesDtl.GetResult(strsaledtl, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim intLoopcount As Short
        Dim varCumulative As Object
        If rsSalesDtl.GetNoRows > 0 Then
            intRecordCount = rsSalesDtl.GetNoRows
            ReDim mdblPrevQty(intRecordCount - 1) ' To get value of Quantity in Arrey for updation in despatch
            ReDim mdblToolCost(intRecordCount - 1) ' To get value of Quantity i
            Call addRowAtEnterKeyPress(intRecordCount - 1)
            rsSalesDtl.MoveFirst()
            If Trim(UCase(lblInvoiceType.Text)) = "NORMAL INVOICE" Or Trim(UCase(lblInvoiceType.Text)) = "JOBWORK INVOICE" Or Trim(UCase(lblInvoiceType.Text)) = "EXPORT INVOICE" Then
                If UCase(Trim(lblInvoiceSubType.Text)) <> "SCRAP" Then
                    For intLoopcount = 1 To intRecordCount
                        mdblToolCost(intLoopcount - 1) = Val(rsSalesDtl.GetValue("Tool_Cost"))
                        rsSalesDtl.MoveNext()
                        ' to incorporated in new form
                    Next
                End If
            End If
            rsSalesDtl.MoveFirst()
            intDecimal = ToGetDecimalPlaces(pstrCurrency)
            Call SetMaxLengthInSpread(intDecimal)
            inti = 1
            For intLoopCounter = inti To intRecordCount
                With Me.SpChEntry
                    .Row = 1 : .Row2 = .MaxRows : .Col = 0 : .Col2 = .MaxCols
                    .Enabled = True : .BlockMode = True : .Lock = True : .BlockMode = False
                    Call .SetText(1, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
                    Call .SetText(2, intLoopCounter, rsSalesDtl.GetValue("Cust_Item_Code"))
                    Call .SetText(3, intLoopCounter, rsSalesDtl.GetValue("Rate"))
                    Call .SetText(4, intLoopCounter, rsSalesDtl.GetValue("Cust_Mtrl"))
                    Call .SetText(5, intLoopCounter, rsSalesDtl.GetValue("Sales_Quantity"))
                    mdblPrevQty(intLoopCounter - 1) = Nothing
                    Call .GetText(5, intLoopCounter, mdblPrevQty(intLoopCounter - 1))
                    Call .SetText(6, intLoopCounter, rsSalesDtl.GetValue("Packing"))
                    Call .SetText(7, intLoopCounter, rsSalesDtl.GetValue("Excise_Type"))
                    Call .SetText(8, intLoopCounter, rsSalesDtl.GetValue("CVD_type"))
                    Call .SetText(9, intLoopCounter, rsSalesDtl.GetValue("SAD_type"))
                    Call .SetText(10, intLoopCounter, rsSalesDtl.GetValue("Others"))
                    Call .SetText(11, intLoopCounter, rsSalesDtl.GetValue("From_Box"))
                    Call .SetText(12, intLoopCounter, rsSalesDtl.GetValue("To_Box"))
                    Call .SetText(15, intLoopCounter, rsSalesDtl.GetValue("tool_Cost"))
                    If intLoopCounter = 1 Then
                        Call .SetText(13, intLoopCounter, (rsSalesDtl.GetValue("To_Box") - rsSalesDtl.GetValue("From_Box")) + 1)
                    Else
                        varCumulative = Nothing
                        Call .GetText(13, intLoopCounter - 1, varCumulative)
                        Call .SetText(13, intLoopCounter, varCumulative + ((rsSalesDtl.GetValue("To_Box") - rsSalesDtl.GetValue("From_Box")) + 1))
                    End If
                    If mblnMultipleSOAllowed = True Then
                        Call .SetText(16, intLoopCounter, rsSalesDtl.GetValue("cust_ref"))
                        Call .SetText(17, intLoopCounter, rsSalesDtl.GetValue("Amendment_no"))
                    End If
                End With
                rsSalesDtl.MoveNext()
            Next intLoopCounter
        End If
        If SpChEntry.MaxRows > 5 Then
            SpChEntry.ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsBoth
        End If
        rsSalesDtl.ResultSetClose()
        rsSalesDtl = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub RefreshForm(ByRef pstrType As String)
        '*****************************************************
        'Created By     -  Kapil
        'Description    -  To Refresh All The Fields
        '*****************************************************
        On Error GoTo ErrHandler
        CmdGrpChEnt.Enabled(1) = False
        Select Case UCase(pstrType)
            Case "LOCATION"
                txtLocationCode.Text = "" : lblLocCodeDes.Text = "" : lblRGPDes.Text = ""
                txtChallanNo.Text = "" : txtCustCode.Text = "" : lblCustCodeDes.Text = "" : lblAddressDes.Text = ""
                txtCarrServices.Text = "" : txtVehNo.Text = ""
                txtFreight.Text = "" : txtSaleTaxType.Text = "" : lblSaltax_Per.Text = "0.00"
                txtSurchargeTaxType.Text = "" : lblSurcharge_Per.Text = "0.00"
                ctlInsurance.Text = "" : lblCurrencyDes.Text = "" : txtRefNo.Text = ""
                '03/06/2002
                lblInvoiceType.Text = "" : lblInvoiceSubType.Text = ""
                '***
                Me.CmdGrpChEnt.Enabled(1) = False
                Me.CmdGrpChEnt.Enabled(2) = False
            Case "CHALLAN"
                txtChallanNo.Text = "" : txtCustCode.Text = "" : lblCustCodeDes.Text = "" : lblAddressDes.Text = ""
                txtCarrServices.Text = "" : txtVehNo.Text = ""
                txtFreight.Text = "" : txtSaleTaxType.Text = "" : lblSaltax_Per.Text = "0.00"
                txtSurchargeTaxType.Text = "" : lblSurcharge_Per.Text = "0.00"
                ctlInsurance.Text = "" : lblRGPDes.Text = ""
                lblInvoiceType.Text = "" : lblInvoiceSubType.Text = "" : lblCurrencyDes.Text = "" : txtRefNo.Text = ""
                '***
                Me.CmdGrpChEnt.Enabled(1) = False
                Me.CmdGrpChEnt.Enabled(2) = False
                Me.CmdGrpChEnt.Enabled(0) = False
        End Select
        With Me.SpChEntry
            .MaxRows = 1
            .Row = 1 : .Row2 = 1 : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
        End With
        txtremark.Enabled = False
        txtremark.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        lblCreditTerm.Text = ""
        lblCreditTermDesc.Text = ""
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub AddTransPortTypeToCombo()
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
    Private Function SetMaxLengthInSpread(ByRef pintDecimalSize As Short) As Object
        '*****************************************************
        'Created By     -  Kapil
        'Description    -  To Set Max Length Of Columns Of Spread
        '*****************************************************
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
                .Col = 1 : .TypeMaxEditLen = 16
                .Col = 2 : .TypeMaxEditLen = 30
                .Col = 3 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax)
                .Col = 4 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax)
                .Col = 5 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("99999999999999.99")
                .Col = 6 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax)
                .Col = 8 : .TypeMaxEditLen = 6
                .Col = 9 : .TypeMaxEditLen = 6
                If Trim(UCase(lblInvoiceType.Text)) = "NORMAL INVOICE" Or Trim(UCase(lblInvoiceType.Text)) = "JOBWORK INVOICE" Or Trim(UCase(lblInvoiceType.Text)) = "EXPORT INVOICE" Then
                    If Trim(UCase(lblInvoiceSubType.Text)) <> "SCRAP" Then
                        .Col = 7
                        .CtlEditMode = False
                    Else
                        .Col = 7
                        .CtlEditMode = True
                    End If
                Else
                    .Col = 7 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                    .CtlEditMode = True
                End If
                .Col = 10 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("99999999999999.99")
                .Col = 13 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                .Col = 14 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                .Col = 15 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("99999999999999.99")
            Next intRow
        End With
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function StockLocationSalesConf(ByRef pstrInvType As String, ByRef pstrInvSubtype As String, ByRef pstrFeild As String) As String
        Dim rsSalesConf As ClsResultSetDB
        Dim StockLocation As String
        On Error GoTo ErrHandler
        rsSalesConf = New ClsResultSetDB
        Select Case pstrFeild
            Case "DESCRIPTION"
                rsSalesConf.GetResult("Select Stock_Location from SaleConf Where UNIT_CODE='" & gstrUNITID & "' and Description ='" & Trim(pstrInvType) & "' and Sub_type_Description ='" & Trim(pstrInvSubtype) & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
            Case "TYPE"
                rsSalesConf.GetResult("Select Stock_Location from SaleConf Where UNIT_CODE='" & gstrUNITID & "' and Invoice_type ='" & Trim(pstrInvType) & "' and Sub_type ='" & Trim(pstrInvSubtype) & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
        End Select
        If UCase(Trim(GetPlantName)) = "HILEX" And CHECK_FTSINVOICE("CHECK_FTS") = True Then
            rsSalesConf.GetResult("Select FTS_STOCK_LOCATION Stock_Location from SaleConf Where UNIT_CODE='" & gstrUNITID & "' and Invoice_type ='" & Trim(pstrInvType) & "' and Sub_type ='" & Trim(pstrInvSubtype) & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
        End If
        If rsSalesConf.GetNoRows > 0 Then
            StockLocation = rsSalesConf.GetValue("Stock_Location")
        End If
        rsSalesConf.ResultSetClose()
        rsSalesConf = Nothing
        StockLocationSalesConf = StockLocation
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    '    Private Function GetServerDate() As Date
    '        Dim objServerDate As ClsResultSetDB 'Class Object
    '        Dim strsql As String 'Stores the SQL statement
    '        On Error GoTo ErrHandler
    '        'Build the SQL statement
    '        strsql = "SELECT CONVERT(datetime,getdate(),103)"
    '        'Creating the instance
    '        objServerDate = New ClsResultSetDB
    '        With objServerDate
    '            'Open the recordset
    '            Call .GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '            'If we have a record, then getting the financial year else exiting
    '            If .GetNoRows <= 0 Then Exit Function
    '            'Getting the date
    '            GetServerDate = CDate(VB6.Format(DateValue(.GetValueByNo(0)), "dd/MM/yyyy"))
    '            'Closing the recordset
    '            .ResultSetClose()
    '        End With
    '        'Releasing the object
    '        objServerDate = Nothing
    '        Exit Function
    'ErrHandler:  'The Error Handling Code Starts here
    '        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    '    End Function
    Public Function ToGetDecimalPlaces(ByRef pstrCurrency As String) As Short
        Dim rscurrency As ClsResultSetDB
        rscurrency = New ClsResultSetDB
        rscurrency.GetResult("Select Decimal_Place from Currency_Mst where UNIT_CODE='" & gstrUNITID & "' and Currency_code ='" & pstrCurrency & "'")
        ToGetDecimalPlaces = Val(rscurrency.GetValue("Decimal_Place"))
        rscurrency.ResultSetClose()
        rscurrency = Nothing
    End Function
    Private Function SelectDataFromTable(ByRef mstrFieldName As String, ByRef mstrTableName As String, ByRef mstrCondition As String) As String
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        'Created By     -   Nitin Sood
        'Description    -   Get Data from BackEnd
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        Dim StrSQLQuery As String
        Dim GetDataFromTable As ClsResultSetDB
        On Error GoTo ErrHandler
        StrSQLQuery = "Select " & mstrFieldName & " From " & mstrTableName & mstrCondition
        GetDataFromTable = New ClsResultSetDB
        If GetDataFromTable.GetResult(StrSQLQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic) Then
            If GetDataFromTable.GetNoRows > 0 Then
                SelectDataFromTable = GetDataFromTable.GetValue(mstrFieldName)
            Else
                SelectDataFromTable = ""
            End If
        Else
            SelectDataFromTable = ""
        End If
        GetDataFromTable.ResultSetClose()
        GetDataFromTable = Nothing
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function GetTaxRate(ByRef pstrFieldText As String, ByRef pstrColumnName As String, ByRef pstrTableName As String, ByRef pstrFieldName_WhichValueRequire As String, Optional ByRef pstrCondition As String = "") As Double
        '****************************************************
        'Created By     -  Tapan
        'Description    -  To Check Validity Of Field Data Whether it Exists In The
        '                  Database Or Not and Return it's Tax Rate
        'Arguments      -  pstrFieldText - Field Text,pstrColumnName - Column Name
        '               -  pstrTableName - Table Name,pstrCondition - Optional Parameter For Condition
        '****************************************************
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
    Private Function GetExchangeRate(ByVal pstrCurrencyCode As String, ByVal pstrDate As String, ByVal IsCustomer As Boolean) As Double
        On Error GoTo ErrHandler
        GetExchangeRate = 1.0#
        Dim strTableSql As String 'Declared To Make Select Query
        Dim rsExistData As ClsResultSetDB
        pstrDate = getDateForDB(pstrDate)
        If IsCustomer = True Then
            strTableSql = "SELECT CExch_MultiFactor FROM Gen_CurExchMaster WHERE UNIT_CODE='" & gstrUNITID & "' and CExch_CurrencyTo='" & Trim(pstrCurrencyCode) & "' AND CExch_InOut=1 AND '" & pstrDate & "' BETWEEN CExch_DateFrom AND CExch_DateTo"
        Else
            strTableSql = "SELECT CExch_MultiFactor FROM Gen_CurExchMaster WHERE UNIT_CODE='" & gstrUNITID & "' and CExch_CurrencyTo='" & Trim(pstrCurrencyCode) & "' AND CExch_InOut=0 AND '" & pstrDate & "' BETWEEN CExch_DateFrom AND CExch_DateTo"
        End If
        rsExistData = New ClsResultSetDB
        rsExistData.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsExistData.GetNoRows > 0 Then
            GetExchangeRate = rsExistData.GetValue("CExch_MultiFactor")
        Else
            GetExchangeRate = 1.0#
        End If
        rsExistData.ResultSetClose()
        rsExistData = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub txtremark_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtremark.Enter
        On Error GoTo ErrHandler
        txtremark.SelectionStart = 0
        txtremark.SelectionLength = Len(txtremark.Text)
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub TxtRemark_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtremark.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Call txtremark_Validating(txtremark, New System.ComponentModel.CancelEventArgs(False))
            Case 34, 39, 96
                KeyAscii = 0
        End Select
ErrHandler:
        GoTo EventExitSub
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtremark_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtremark.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        With CmdGrpChEnt
            .Focus() 'Move Focus To Cancellation Button
        End With
        GoTo EventExitSub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtSaleTaxType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtSaleTaxType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If Len(txtSaleTaxType.Text) > 0 Then
            If CheckExistanceOfFieldData((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST' OR Tx_TaxeID='VAT') and UNIT_CODE='" & gstrUNITID & "'") Then
                lblSaltax_Per.Text = CStr(GetTaxRate((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST' OR Tx_TaxeID='VAT') AND  UNIT_CODE='" & gstrUNITID & "'"))
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
    Private Function SelectItemfromsaleDtl(ByRef pstrchallanNo As Object) As String
        On Error GoTo ErrHandler
        Dim strsaledtl As String
        Dim rssaledtl As ClsResultSetDB
        Dim intRecordCount As Short 'To Hold Record Count
        Dim intCount As Short
        Dim strItemText As String
        'changed due to more then 4 item selection in case of Export in voice.
        '****************************
        strsaledtl = ""
        strsaledtl = "Select a.Item_Code,a.Cust_ITem_Code,a.Cust_Item_Desc,b.Tariff_Code from Sales_Dtl a,Item_Mst b where a.ITem_code = b.ITem_code and a.UNIT_CODE = b.UNIT_CODE  AND a.UNIT_CODE='" & gstrUNITID & "' and Doc_No ="
        strsaledtl = strsaledtl & pstrchallanNo
        rssaledtl = New ClsResultSetDB
        rssaledtl.GetResult(strsaledtl, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        intRecordCount = rssaledtl.GetNoRows 'assign record count to integer variable
        If intRecordCount > 0 Then '          'if record found
            rssaledtl.MoveFirst() 'move to first record
            For intCount = 1 To intRecordCount
                strItemText = strItemText & "'" & Trim(rssaledtl.GetValue("Cust_Item_code")) & "',"
                rssaledtl.MoveNext() 'move to next record
            Next intCount
            rssaledtl.ResultSetClose()
            rssaledtl = Nothing
        End If
        SelectItemfromsaleDtl = strItemText
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function updatesalesconfandsaleschallan() As Boolean
        Dim strsql As String
        Dim rsSalesChallan As ClsResultSetDB
        Dim dblInvoiceAmt As Double
        Dim saleschallan As String
        Dim salesconf As String
        On Error GoTo Err_Handler
        updatesalesconfandsaleschallan = True
        strsql = "select *  from Saleschallan_dtl where UNIT_CODE='" & gstrUNITID & "' AND Doc_No = " & Trim(txtChallanNo.Text)
        strsql = strsql & " and Invoice_type = '" & mstrInvType & "'  and  sub_category =  '" & mstrInvSubType & "' and Location_Code='" & Trim(txtLocationCode.Text) & "'"
        rsSalesChallan = New ClsResultSetDB
        rsSalesChallan.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSalesChallan.GetNoRows > 0 Then
            dblInvoiceAmt = rsSalesChallan.GetValue("total_amount")
        End If
        rsSalesChallan.ResultSetClose()
        rsSalesChallan = Nothing
        If blnEOU_FLAG = True Then
            salesconf = "update saleconf set OpenningBal = openningBal - " & dblInvoiceAmt & " where UNIT_CODE='" & gstrUNITID & "' AND Invoice_type <> 'EXP' and Location_Code='" & Trim(txtLocationCode.Text) & "'"
        End If
        saleschallan = "UPDATE SalesChallan_Dtl SET Cancel_flag=1 , remarks='" & Trim(txtremark.Text) & "',Upd_Userid='" & mP_User & "',Upd_dt=getdate() WHERE UNIT_CODE='" & gstrUNITID & "' AND Doc_No=" & txtChallanNo.Text & " and Location_Code='" & Trim(txtLocationCode.Text) & "'"
        If Len(salesconf) > 0 Then mP_Connection.Execute(salesconf, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        If Len(saleschallan) > 0 Then mP_Connection.Execute(saleschallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        updatesalesconfandsaleschallan = False
    End Function
    Private Function UpdateGrnHdr(ByRef pdblGrinNo As Double, ByRef pdblinvoiceNo As Double) As Boolean
        '---------------------------------------------------------------------------------------
        'Name       :   UpdateGrnHdr
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        Dim rsSalesDtl As ClsResultSetDB
        Dim intMaxLoop As Short
        Dim strItemCode As String
        Dim dblQty As Double
        Dim intLoopcount As Short
        Dim strupdateGrinhdr As String
        On Error GoTo ErrHandler
        UpdateGrnHdr = True
        rsSalesDtl = New ClsResultSetDB
        rsSalesDtl.GetResult("select * from sales_dtl where UNIT_CODE='" & gstrUNITID & "' AND Doc_No = " & pdblinvoiceNo & " and Location_Code='" & Trim(txtLocationCode.Text) & "'")
        If rsSalesDtl.GetNoRows > 0 Then
            intMaxLoop = rsSalesDtl.GetNoRows
            rsSalesDtl.MoveFirst()
            strupdateGrinhdr = ""
            For intLoopcount = 1 To intMaxLoop
                strItemCode = rsSalesDtl.GetValue("ITem_code")
                dblQty = rsSalesDtl.GetValue("Sales_Quantity")
                If Len(Trim(strupdateGrinhdr)) = 0 Then
                    strupdateGrinhdr = "Update Grn_Dtl Set Despatch_Quantity = isnull(Despatch_Quantity,0) -" & dblQty
                    strupdateGrinhdr = strupdateGrinhdr & " Where UNIT_CODE='" & gstrUNITID & "' AND ITem_Code = '" & strItemCode & "' and Doc_No = " & pdblGrinNo
                Else
                    strupdateGrinhdr = strupdateGrinhdr & vbCrLf & "Update Grn_Dtl Set Despatch_Quantity = isnull(Despatch_Quantity,0) - " & dblQty
                    strupdateGrinhdr = strupdateGrinhdr & " Where UNIT_CODE='" & gstrUNITID & "' AND ITem_Code = '" & strItemCode & "' and Doc_No = " & pdblGrinNo
                End If
                rsSalesDtl.MoveNext()
            Next
            If Len(strupdateGrinhdr) > 0 Then mP_Connection.Execute(strupdateGrinhdr, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Else
            MsgBox("No Items Available in Invoice " & txtChallanNo.Text)
            UpdateGrnHdr = False
        End If
        rsSalesDtl.ResultSetClose()
        rsSalesDtl = Nothing
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        UpdateGrnHdr = False
    End Function
    Private Function UpdateinSale_Dtl() As Boolean
        Dim rssaledtl As ClsResultSetDB
        Dim strupdateitbalmst As String
        Dim strsql As String
        Dim strStockLocCode As String
        Dim intRow, intLoopcount As Short
        Dim mItem_Code As String
        Dim mSales_Quantity As Double
        strupdateitbalmst = ""
        On Error GoTo Err_Handler
        UpdateinSale_Dtl = True

        'If UCase(Trim(GetPlantName)) = "HILEX" And CHECK_FTSINVOICE() = True Then
        '    ''MsgBox("Invoice is made for FTS Items , cannot be Cancelled.", MsgBoxStyle.Exclamation, ResolveResString(100))
        '    'If FTS_INVOICECANCELLATION() = False Then
        '    '    mP_Connection.RollbackTrans()
        '    '    Exit Function
        '    'End If
        '    strStockLocCode = "01P2"
        'Else
        '    strStockLocCode = StockLocationSalesConf(mstrInvType, mstrInvSubType, "TYPE")
        'End If

        'Added by praveen on 19.12.2017 
        If UCase(Trim(GetPlantName)) = "HILEX" And CHECK_FTSINVOICE("CHECK_FTS") = True Then
            UpdateinSale_Dtl = True
            Exit Function
        End If

        strStockLocCode = StockLocationSalesConf(mstrInvType, mstrInvSubType, "TYPE")
        strsql = "Select * from sales_Dtl where UNIT_CODE='" & gstrUNITID & "' AND Doc_No = " & Trim(txtChallanNo.Text) & " and Location_Code='" & Trim(txtLocationCode.Text) & "'"
        rssaledtl = New ClsResultSetDB
        rssaledtl.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rssaledtl.GetNoRows > 0 Then
            intRow = rssaledtl.GetNoRows
            rssaledtl.MoveFirst()
            For intLoopcount = 1 To intRow
                If Not rssaledtl.EOFRecord Then
                    mItem_Code = rssaledtl.GetValue("Item_Code")
                    mSales_Quantity = IIf(rssaledtl.GetValue("Sales_Quantity") = "", 0, rssaledtl.GetValue("Sales_Quantity"))
                    strupdateitbalmst = Trim(strupdateitbalmst) & "Update Itembal_mst set cur_bal= cur_bal+"
                    strupdateitbalmst = strupdateitbalmst & mSales_Quantity & " where UNIT_CODE='" & gstrUNITID & "' AND Location_code = '" & strStockLocCode
                    strupdateitbalmst = strupdateitbalmst & "' and item_code = '" & mItem_Code & "'" & vbCrLf
                    rssaledtl.MoveNext()
                End If
            Next
            If Len(strupdateitbalmst) > 0 Then mP_Connection.Execute(strupdateitbalmst, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        End If
        rssaledtl.ResultSetClose()
        rssaledtl = Nothing
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        UpdateinSale_Dtl = False
    End Function
    Private Function UpdateFinance() As Boolean
        '---------------------------------------------------------------------------------------
        'Name       :   UpdateFinance
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        Dim objDrCrNote As New prj_DrCrNote.cls_DrCrNote(GetServerDate)
        Dim strResult As String
        Dim strRemark As String
        Dim strInvoiceType As String
        Dim rsinvoice_type As New ClsResultSetDB
        On Error GoTo ErrHandler
        rsinvoice_type.GetResult("select invoice_type from saleschallan_dtl where UNIT_CODE='" & gstrUNITID & "' AND doc_no = '" & txtChallanNo.Text & "' and Location_code = '" & txtLocationCode.Text & "'")
        strInvoiceType = rsinvoice_type.GetValue("Invoice_type")
        rsinvoice_type.ResultSetClose()
        rsinvoice_type = Nothing
        strRemark = QuoteRem(Trim(txtremark.Text))
        If ChkForRejInv() Then
            strResult = objDrCrNote.ReverseAPDocument(gstrUNITID, mstrInvNo, mP_User, getDateForDB(lblDateDes.Text), strRemark, gstrCURRENCYCODE, , gstrCONNECTIONSTRING)
        Else
            If UCase(strInvoiceType) <> "REJ" Then
                strResult = objDrCrNote.ReverseARInvoiceDocument(gstrUNITID, Trim(txtChallanNo.Text), mP_User, getDateForDB(lblDateDes.Text), strRemark, gstrCURRENCYCODE, , gstrCONNECTIONSTRING)
            End If
        End If
        'Changes ends here
        If UCase(VB.Left(strResult, 1)) = "N" Then
            MsgBox(VB.Left(VB.Right(strResult, Len(strResult) - 3), Len(VB.Right(strResult, Len(strResult) - 3)) - 1), MsgBoxStyle.Critical, "empower")
            UpdateFinance = False
        Else
            UpdateFinance = True
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        UpdateFinance = False
    End Function
    Private Function UpdateCustAnnex() As Boolean
        '---------------------------------------------------------------------------------------
        'Name       :   UpdateCustAnnex
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        Dim rsCustAnnexDtl As ClsResultSetDB
        Dim strsql As String
        Dim lintLoop As Short
        Dim strItemCode As String
        Dim dblItemQty As Double
        Dim strRef57F4 As String
        Dim strCustCode As String
        Dim strUpdtCustAnnexHdr As String
        On Error GoTo ErrHandler
        UpdateCustAnnex = True
        rsCustAnnexDtl = New ClsResultSetDB
        For lintLoop = 1 To SpChEntry.MaxRows
            SpChEntry.Col = 1
            SpChEntry.Row = lintLoop
            strItemCode = Trim(SpChEntry.Text)
            strsql = "SELECT Quantity,Ref57F4_no,Customer_code FROM CustAnnex_Dtl WHERE UNIT_CODE='" & gstrUNITID & "' AND Invoice_No = " & Trim(txtChallanNo.Text) & " and Item_Code='" & Trim(strItemCode) & "'"
            rsCustAnnexDtl.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsCustAnnexDtl.GetNoRows > 0 Then
                dblItemQty = rsCustAnnexDtl.GetValue("Quantity")
                strCustCode = rsCustAnnexDtl.GetValue("Customer_code")
                strRef57F4 = rsCustAnnexDtl.GetValue("Ref57F4_no")
                strUpdtCustAnnexHdr = "UPDATE custannex_hdr SET balance_Qty=balance_Qty +" & dblItemQty & " WHERE UNIT_CODE='" & gstrUNITID & "' AND Customer_Code='" & strCustCode & "' AND Item_Code='" & strItemCode & "' AND ref57f4_no='" & strRef57F4 & "'" & vbCrLf
            End If
            rsCustAnnexDtl.ResultSetClose()
        Next
        If Len(strUpdtCustAnnexHdr) > 0 Then mP_Connection.Execute(strUpdtCustAnnexHdr, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        UpdateCustAnnex = False
    End Function
    Private Function ChkForRejInv() As Boolean
        '--------To retrieve Dr Note Raised against the Rejection Invoice
        '--------The Retrieved Document no will further be called for Cancellation by ClsDrCrNote using Function ReverseAPDocument
        Dim rsObjInv As New ADODB.Recordset
        ChkForRejInv = False
        On Error GoTo ErrHandler
        If rsObjInv.State = ADODB.ObjectStateEnum.adStateOpen Then rsObjInv.Close()
        rsObjInv.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rsObjInv.Open("SELECT apdocm_vono FROM ap_docmaster WHERE apdocM_unit ='" & gstrUNITID & "' AND apdocm_venprdocno='" & Trim(txtChallanNo.Text) & "' AND APDOCM_VENDORCODE = '" & txtCustCode.Text.Trim & "' AND APDOCM_VOTYPE = 'M' AND apdocm_open=1 AND apdocm_cancel=0 ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not rsObjInv.EOF Then
            'Column seperator 
            mstrInvNo = Trim(rsObjInv.Fields(0).Value) & "»"
            ChkForRejInv = True
        End If
        If rsObjInv.State = ADODB.ObjectStateEnum.adStateOpen Then rsObjInv.Close()
        rsObjInv = Nothing
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function UpdateMonthlyMktSchedule(ByRef pstrItemCode As String, ByRef pstrCustPartCode As String, ByRef pstrDSNo As String) As Boolean
        Dim RsObjUpdateSchedules As ADODB.Recordset
        Dim RsobjSchedules As ADODB.Recordset
        Dim strsql As String
        Dim dblQty As String
        Dim strYearMonth As String
        On Error GoTo ErrHandler
        UpdateMonthlyMktSchedule = True
        RsObjUpdateSchedules = New ADODB.Recordset
        RsobjSchedules = New ADODB.Recordset
        If RsObjUpdateSchedules.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjUpdateSchedules.Close()
        RsObjUpdateSchedules.Open("SELECT doc_no ,item_code,cust_part_code,dsno,customer_code,quantityknockedoff FROM mkt_invdshistory WHERE UNIT_CODE='" & gstrUNITID & "' AND cancellation_flag=0 AND doc_no='" & txtChallanNo.Text & "' and Item_code = '" & pstrItemCode & "' and cust_Part_code = '" & pstrCustPartCode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        strYearMonth = Year(ConvertToDate(lblDateDes.Text)) & Month(ConvertToDate(lblDateDes.Text)).ToString.PadLeft(2, "0")

        'If Val(Mid(lblDateDes.Text, 4, 2)) < 10 Then
        '    strYearMonth = Mid(lblDateDes.Text, 7, 4) & "0" & Val(Mid(lblDateDes.Text, 4, 2))
        'Else
        '    strYearMonth = Mid(lblDateDes.Text, 7, 4) & Val(Mid(lblDateDes.Text, 4, 2))
        'End If
        dblQty = RsObjUpdateSchedules.Fields(5).Value
        If CDbl(dblQty) > 0 Then
            strsql = "SELECT dsno,isNULL(despatch_Qty,0) as despatch_Qty ,item_code,CUST_DRGNO FROM Monthlymktschedule WHERE UNIT_CODE='" & gstrUNITID & "' AND account_code='" & RsObjUpdateSchedules.Fields(4).Value & "' AND item_code='" & RsObjUpdateSchedules.Fields(1).Value & "' AND status=1 AND despatch_qty > 0 AND dsno ='" & RsObjUpdateSchedules.Fields(3).Value & "' and year_Month <= '" & strYearMonth & "' and Cust_drgNo = '" & RsObjUpdateSchedules.Fields(2).Value & "' ORDER BY Year_Month DESC,dsdatetime DESC"
            mP_Connection.Execute("Set Dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            RsobjSchedules.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If Not RsobjSchedules.EOF Then
                Do While Not RsobjSchedules.EOF
                    If CDbl(dblQty) > 0 Then
                        If CDbl(dblQty) >= Val(RsobjSchedules.Fields(1).Value) Then
                            mP_Connection.Execute("UPDATE MonthlyMKTSCHEDULE SET DESPATCH_QTY= 0 WHERE UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE='" & RsObjUpdateSchedules.Fields(1).Value & "' AND CUST_DRGNO='" & RsObjUpdateSchedules.Fields(2).Value & "' AND DSNO='" & RsObjUpdateSchedules.Fields(3).Value & "' AND account_code='" & RsObjUpdateSchedules.Fields(4).Value & "' and status =1 and Year_Month = '" & strYearMonth & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            dblQty = CStr(CDbl(dblQty) - RsobjSchedules.Fields(1).Value)
                        Else
                            mP_Connection.Execute("UPDATE MONTHLYMKTSCHEDULE SET DESPATCH_QTY= (DESPATCH_QTY - " & dblQty & ") WHERE UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE='" & RsObjUpdateSchedules.Fields(1).Value & "' AND CUST_DRGNO='" & RsObjUpdateSchedules.Fields(2).Value & "' AND DSNO='" & RsObjUpdateSchedules.Fields(3).Value & "' AND account_code='" & RsObjUpdateSchedules.Fields(4).Value & "' and status =1 and Year_Month = '" & strYearMonth & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            dblQty = CStr(0)
                        End If
                        RsobjSchedules.MoveNext()
                    Else
                        Exit Do
                    End If
                Loop
            End If
            RsObjUpdateSchedules.MoveNext()
        End If
        strUpdateInvDsHistory = "UPDATE mkt_invdshistory SET cancellation_flag=1 WHERE UNIT_CODE='" & gstrUNITID & "' AND Location_code = '" & txtLocationCode.Text & "' and doc_no='" & txtChallanNo.Text & "' "
        Exit Function
ErrHandler:
        UpdateMonthlyMktSchedule = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function DispatchQuantityFromDailyScheduleForDSTracking(ByVal pstrAccountCode As String, ByVal pstrCustomerDrawNo As String, ByVal pstrItemCode As String, ByVal pstrDate As String, ByVal pstrDSNo As String) As Boolean
        '---------------------------------------------------------------------------------------
        'Name       :   DispatchQuantityFromDailySchedule
        'Type       :   Function
        'Author     :   Tapan Jain (Modified By     -   Nitin Sood)
        'Arguments  :
        'Return     :   TRUE - Despatch was done against DailyMKTSchedule
        'Purpose    :   To Find out ,If Despatch was againt Monthly Schedule for Daily Schedule
        '---------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strScheduleSql As String
        Dim objRsForSchedule As New ADODB.Recordset
        Dim ldblTotalDispatchQuantity As Double
        Dim ldblTotalScheduleQuantity As Double
        Dim lintLoopCounter As Short
        ldblTotalDispatchQuantity = 0 'INITIALIZE
        ldblTotalScheduleQuantity = 0
        mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        strScheduleSql = "Select Schedule_Quantity,Despatch_Qty from DailyMktSchedule where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & pstrAccountCode & "'"
        strScheduleSql = strScheduleSql & " and Trans_Date <='" & getDateForDB(pstrDate) & "'"
        strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND ITEM_CODE  = '"
        strScheduleSql = strScheduleSql & pstrItemCode & "' and Status =1 AND DSNo = '" & pstrDSNo & "' ORDER BY Trans_Date DESC" '''and Schedule_Flag =1   ( Now Not Consider)
        If objRsForSchedule.State = 1 Then objRsForSchedule.Close()
        objRsForSchedule.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        objRsForSchedule.Open(strScheduleSql, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        If objRsForSchedule.EOF Or objRsForSchedule.BOF Then
            DispatchQuantityFromDailyScheduleForDSTracking = False
            mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            objRsForSchedule.Close()
            Exit Function
        Else
            DispatchQuantityFromDailyScheduleForDSTracking = True
            mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            objRsForSchedule.Close()
            Exit Function
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        DispatchQuantityFromDailyScheduleForDSTracking = False
        mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function DispatchQuantityFromMonthlyScheduleForDSTracking(ByVal pstrAccountCode As String, ByVal pstrCustomerDrawNo As String, ByVal pstrItemCode As String, ByVal pstrDate As String, ByVal pstrDSNo As String) As Boolean
        '---------------------------------------------------------------------------------------
        'Name       :   GetTotalDispatchQuantityFromDailySchedule
        'Type       :   Function
        'Author     :   Tapan Jain (Modified By     -   Nitin Sood)
        'Arguments  :
        'Return     :   TRUE    -   If Schedule is Monthly
        'Purpose    :
        '---------------------------------------------------------------------------------------
        Dim strScheduleSql As String
        Dim objRsForSchedule As New ADODB.Recordset
        Dim lintLoopCounter As Short
        Dim strMakeDate As String
        On Error GoTo ErrHandler
        mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        strMakeDate = Year(ConvertToDate(pstrDate)) & Month(ConvertToDate(pstrDate)).ToString.PadLeft(2, "0")

        'If Val(Mid(pstrDate, 4, 2)) < 10 Then
        '    strMakeDate = Mid(pstrDate, 7, 4) & "0" & Val(Mid(pstrDate, 4, 2))
        'Else
        '    strMakeDate = Mid(pstrDate, 7, 4) & Val(Mid(pstrDate, 4, 2))
        'End If
        If objRsForSchedule.State = 1 Then objRsForSchedule.Close()
        objRsForSchedule.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        strScheduleSql = "Select Schedule_Qty,isnull(Despatch_Qty,0) AS Despatch_qty  from MonthlyMktSchedule where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & pstrAccountCode & "' and "
        strScheduleSql = strScheduleSql & " Year_Month <='" & Val(Trim(strMakeDate)) & "'"
        strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND Item_code = '" & pstrItemCode & "' and status =1  and DSNo = '" & pstrDSNo & "'"
        objRsForSchedule.Open(strScheduleSql, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        If objRsForSchedule.EOF Or objRsForSchedule.BOF Then
            DispatchQuantityFromMonthlyScheduleForDSTracking = False
            mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            objRsForSchedule.Close()
            Exit Function
        Else
            objRsForSchedule.MoveFirst()
            DispatchQuantityFromMonthlyScheduleForDSTracking = True 'Dispatch Against Monthly Schedule
            mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            objRsForSchedule.Close()
            Exit Function
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        DispatchQuantityFromMonthlyScheduleForDSTracking = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function UpdateDailyMktSchedule(ByRef pstrItemCode As String, ByRef pstrCustPartCode As String, ByRef pstrDSNo As String) As Boolean
        Dim RsObjUpdateSchedules As ADODB.Recordset
        Dim RsobjSchedules As ADODB.Recordset
        Dim strsql As String
        Dim dblQty As String
        On Error GoTo ErrHandler
        UpdateDailyMktSchedule = True
        RsObjUpdateSchedules = New ADODB.Recordset
        RsobjSchedules = New ADODB.Recordset
        If RsObjUpdateSchedules.State = 1 Then
            RsObjUpdateSchedules.Close()
        End If
        strsql = "SELECT doc_no ,item_code,cust_part_code,dsno,customer_code,quantityknockedoff FROM mkt_invdshistory WHERE UNIT_CODE='" & gstrUNITID & "' AND cancellation_flag=0 AND doc_no='" & txtChallanNo.Text & "' and ITem_code = '" & pstrItemCode & "' and Cust_part_code = '" & pstrCustPartCode & "' and DSNo = '" & pstrDSNo & "'"
        RsObjUpdateSchedules.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        dblQty = RsObjUpdateSchedules.Fields(5).Value
        If CDbl(dblQty) > 0 Then
            strsql = "SELECT dsno,isNULL(despatch_qty,0) as despatch_qty ,item_code,CUST_DRGNO,trans_date FROM dailymktschedule WHERE UNIT_CODE='" & gstrUNITID & "' AND account_code='" & RsObjUpdateSchedules.Fields(4).Value & "' AND item_code='" & RsObjUpdateSchedules.Fields(1).Value & "' AND status=1 AND despatch_qty > 0 AND dsno ='" & RsObjUpdateSchedules.Fields(3).Value & "' and trans_date <= '" & getDateForDB(lblDateDes.Text) & "' and Cust_drgNo = '" & RsObjUpdateSchedules.Fields(2).Value & "' ORDER BY trans_date DESC,dsdatetime DESC"
            mP_Connection.Execute("Set Dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            If RsobjSchedules.State = 1 Then RsobjSchedules.Close()
            RsobjSchedules.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If Not RsobjSchedules.EOF Then
                Do While Not RsobjSchedules.EOF
                    If CDbl(dblQty) > 0 Then
                        If CDbl(dblQty) >= Val(RsobjSchedules.Fields(1).Value) Then
                            mP_Connection.Execute("UPDATE DAILYMKTSCHEDULE SET DESPATCH_QTY= 0  WHERE UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE='" & RsObjUpdateSchedules.Fields(1).Value & "' AND CUST_DRGNO='" & RsObjUpdateSchedules.Fields(2).Value & "' AND DSNO='" & RsObjUpdateSchedules.Fields(3).Value & "' AND account_code='" & RsObjUpdateSchedules.Fields(4).Value & "' and status =1 and trans_date = '" & RsobjSchedules.Fields(4).Value & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            dblQty = CStr(CDbl(dblQty) - RsobjSchedules.Fields(1).Value)
                        Else
                            mP_Connection.Execute("UPDATE DAILYMKTSCHEDULE SET DESPATCH_QTY= (DESPATCH_QTY - " & dblQty & ") WHERE UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE='" & RsObjUpdateSchedules.Fields(1).Value & "' AND CUST_DRGNO='" & RsObjUpdateSchedules.Fields(2).Value & "' AND DSNO='" & RsObjUpdateSchedules.Fields(3).Value & "' AND account_code='" & RsObjUpdateSchedules.Fields(4).Value & "' and status =1 and trans_date = '" & RsobjSchedules.Fields(4).Value & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            dblQty = CStr(0)
                        End If
                        RsobjSchedules.MoveNext()
                    Else
                        Exit Do
                    End If
                Loop
            End If
            RsObjUpdateSchedules.MoveNext()
        End If
        RsObjUpdateSchedules.Close()
        RsObjUpdateSchedules = Nothing
        strUpdateInvDsHistory = "UPDATE mkt_invdshistory SET cancellation_flag=1 WHERE UNIT_CODE='" & gstrUNITID & "' AND Location_code = '" & txtLocationCode.Text & "' and doc_no='" & txtChallanNo.Text & "'"
        Exit Function
ErrHandler:
        UpdateDailyMktSchedule = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub CheckMultipleSOAllowed(ByVal pInvType As String, ByVal pInvSubType As String)
        '-----------------------------------------------------------------------------------
        'Created By      : Manoj Kr.Vaish
        'Issue ID        : 19992
        'Creation Date   : 30 JUNE 2007
        'Procedure       : To Check MultipleSOAllowed for Any Invoice Type
        '-----------------------------------------------------------------------------------
        Dim rsCheckSo As ClsResultSetDB
        Dim strsql As String
        On Error GoTo ErrHandler
        rsCheckSo = New ClsResultSetDB
        strsql = "select isnull(sorequired,0) as SORequired,isnull(MultipleSOAllowed,0) as MultipleSOAllowed,isnull(BatchTrackingAllowed,0) as BatchTrackingAllowed from saleconf where UNIT_CODE='" & gstrUNITID & "' AND description='" & Trim(pInvType) & "' and sub_type_description='" & Trim(pInvSubType) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())"
        rsCheckSo.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsCheckSo.GetNoRows > 0 Then
            mblnMultipleSOAllowed = rsCheckSo.GetValue("MultipleSOAllowed")
            mblnSORequired = rsCheckSo.GetValue("SORequired")
            mblnBatckTrackingAllowed = rsCheckSo.GetValue("BatchTrackingAllowed")
        End If
        rsCheckSo.ResultSetClose()
        rsCheckSo = Nothing
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
        '------------------------------------------------------------------------------------------
    End Sub
    Private Function CheckMktSchedules() As String
        '-------------------------------------------------------------------------------------------
        ' Author        : Manoj Kr. Vaish
        ' Arguments     : NIL
        ' Return Value  : 'Error'  - If error occured during processing
        '                 Msg if Schedule doesn't exist for Item(s)
        ' Function      : To Check Daily and Monthly Schedules
        ' Datetime      : 2 July 2007
        ' Issue ID      : 19992
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim Com As ADODB.Command
        Dim strsql As String
        Dim IntCtr As Short
        Dim strMSG As String
        Dim strYYYYmm As String
        ReDim mSchTypeArr(0)
        CheckMktSchedules = ""
        strYYYYmm = Year(ConvertToDate(lblDateDes.Text)) & VB.Right("0" & Month(ConvertToDate(lblDateDes.Text)), 2)
        If mdblPrevQty Is Nothing Then
            mdblPrevQty(0) = 0 ' changed by Roshan Singh on 06 july 2011
        End If
        With SpChEntry
            For IntCtr = 1 To .MaxRows Step 1
                ReDim Preserve mSchTypeArr(IntCtr)
                Com = New ADODB.Command
                Com.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Com.CommandText = "MKT_SCHEDULE_CHECK_NORTH"
                .Row = IntCtr
                Com.Parameters.Append(Com.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                Com.Parameters.Append(Com.CreateParameter("@CUSTOMER_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(txtCustCode.Text)))
                .Col = 1
                Com.Parameters.Append(Com.CreateParameter("@ITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(.Text)))
                .Col = 2
                Com.Parameters.Append(Com.CreateParameter("@CUSTDRG_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(.Text)))
                .Col = 5
                Com.Parameters.Append(Com.CreateParameter("@REQ_QTY", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adCurrency, Val(Trim(.Text)) - mdblPrevQty(IntCtr - 1)))
                Com.Parameters.Append(Com.CreateParameter("@YYYYMM", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adInteger, strYYYYmm))
                Com.Parameters.Append(Com.CreateParameter("@DATE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adVarChar, getDateForDB(lblDateDes.Text)))
                Com.Parameters.Append(Com.CreateParameter("@SCH_TYPE", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamOutput, 1))
                Com.Parameters.Append(Com.CreateParameter("@MSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 500))
                Com.Parameters.Append(Com.CreateParameter("@ERR", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 100))
                Com.let_ActiveConnection(mP_Connection)
                Com.Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                If Len(Com.Parameters(8).Value) > 0 Then
                    strMSG = strMSG & Com.Parameters(8).Value
                End If
                If Len(Com.Parameters(9).Value) > 0 Then
                    MsgBox(Com.Parameters(9).Value, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))  'commented on 14feb '00rr00' not found in spread 6.0
                    CheckMktSchedules = "Error"
                    Com = Nothing
                    Exit Function
                End If
                mSchTypeArr(IntCtr) = Com.Parameters(7).Value
                Com = Nothing
            Next IntCtr
        End With
        CheckMktSchedules = strMSG
        Exit Function
ErrHandler:
        Com = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function UpdateMktSchedules(ByVal pstrUpdType As String) As Boolean
        '-------------------------------------------------------------------------------------------
        ' Author        : Manoj Kr. Vaish
        ' Arguments     : '+' - If Despatch is to be Updated agst Schedule
        '                 '-' - If Reversal is to be made agst Despatched Qty
        ' Return Value  : TRUE  - If Successfull
        '                 FALSE - If error Occured during processing
        ' Function      : To Update Daily and Monthly Schedules
        ' Datetime      : 02 July 2007
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim Com As ADODB.Command
        Dim strsql As String
        Dim IntCtr As Short
        Dim strMSG As String
        Dim strYYYYmm As String
        Dim curQty As Decimal
        UpdateMktSchedules = True
        strYYYYmm = Year(ConvertToDate(lblDateDes.Text)) & VB.Right("0" & Month(ConvertToDate(lblDateDes.Text)), 2)
        With SpChEntry
            For IntCtr = 1 To .MaxRows Step 1
                Com = New ADODB.Command
                Com.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Com.CommandText = "MKT_SCHEDULE_KNOCKOFF_NORTH"
                .Row = IntCtr
                Com.Parameters.Append(Com.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                Com.Parameters.Append(Com.CreateParameter("@CUSTOMER_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(txtCustCode.Text)))
                .Col = 1
                Com.Parameters.Append(Com.CreateParameter("@ITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(.Text)))
                .Col = 2
                Com.Parameters.Append(Com.CreateParameter("@CUSTDRG_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(.Text)))
                Com.Parameters.Append(Com.CreateParameter("@FLAG", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, pstrUpdType))
                Com.Parameters.Append(Com.CreateParameter("@SCH_TYPE", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(mSchTypeArr(IntCtr))))
                Com.Parameters.Append(Com.CreateParameter("@YYYYMM", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adInteger, strYYYYmm))
                .Col = 5
                If pstrUpdType = "+" Then
                    Com.Parameters.Append(Com.CreateParameter("@REQ_QTY", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adCurrency, Trim(.Text)))
                Else
                    Com.Parameters.Append(Com.CreateParameter("@REQ_QTY", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adCurrency, mdblPrevQty(IntCtr - 1)))
                End If
                Com.Parameters.Append(Com.CreateParameter("@DATE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 11, getDateForDB(lblDateDes.Text)))
                Com.Parameters.Append(Com.CreateParameter("@MSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 500))
                Com.Parameters.Append(Com.CreateParameter("@ERR", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 100))
                Com.let_ActiveConnection(mP_Connection)
                Com.Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                If Len(Com.Parameters(9).Value) > 0 Then
                    MsgBox(Com.Parameters(9).Value, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    UpdateMktSchedules = False
                    Com = Nothing
                    Exit Function
                End If
                If Len(Com.Parameters(10).Value) > 0 Then
                    MsgBox(Com.Parameters(10).Value, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    UpdateMktSchedules = False
                    Com = Nothing
                    Exit Function
                End If
                Com = Nothing
            Next IntCtr
        End With
        Exit Function
ErrHandler:
        Com = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function InvAgstBarCode() As Boolean
        '----------------------------------------------------------------------------
        'Author         :   Manoj Kr. Vaish
        'Function       :   Get the BarCodefor Invoice from sales_parameter
        'Comments       :   Date: 19 Sep 2007 ,Issue Id: 21105
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strQry As String
        Dim Rs As ClsResultSetDB
        InvAgstBarCode = False
        strQry = "Select isnull(BarCodeTrackingInInvoice,0) as BarCodeTrackingInInvoice from sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'"
        Rs = New ClsResultSetDB
        If Rs.GetResult(strQry) = False Then GoTo ErrHandler
        If Rs.GetValue("BarCodeTrackingInInvoice") = "True" Then
            strQry = "Select isnull(a.BarcodeTrackingAllowed,0) as BarcodeTrackingAllowed"
            strQry = strQry & " from SaleConf a,SalesChallan_Dtl b where Doc_No ='" & Trim(txtChallanNo.Text) & "'"
            strQry = strQry & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category AND a.UNIT_CODE = b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "' and "
            strQry = strQry & " a.Location_Code = b.Location_Code And (Fin_Start_Date <= getDate() And Fin_End_Date >= getDate())"
            Rs.ResultSetClose()
            Rs = New ClsResultSetDB
            Call Rs.GetResult(strQry, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If Rs.GetNoRows > 0 Then
                If Rs.GetValue("BarcodeTrackingAllowed") = "True" Then
                    InvAgstBarCode = True
                Else
                    InvAgstBarCode = False
                End If
            End If
        End If
        Rs.ResultSetClose()
        Rs = Nothing
        Exit Function
ErrHandler:
        Rs = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function BarCodeTracking(ByVal pstrInvNo As String) As Boolean
        '----------------------------------------------------------------------------
        'Author         :   Manoj Kr. Vaish
        'Argument       :   Invoice Numbers.
        'Return Value   :   True or False
        'Function       :   Update Bar_BondedStock while invoice Cancellation
        'Comments       :   Date: 14 Sep 2007 ,Issue Id: 21105
        'Revision Date  :   28 Nov 2008 Issue ID : eMpro-20080930-22159
        'History        :   Functionality of Raw Material,Input Invoice through Bar Code
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rsGetQty As ClsResultSetDB
        Dim strsql As String
        Dim CurInvoiceQty As Decimal
        Dim CurBarBondedQty As Decimal
        rsGetQty = New ClsResultSetDB
        mstrupdateBarBondedStockQty = ""
        If UCase(Trim(lblInvoiceSubType.Text)) = "RAW MATERIAL" Or UCase(Trim(lblInvoiceSubType.Text)) = "INPUTS" Or UCase(Trim(lblInvoiceSubType.Text)) = "COMPONENTS" Then
            strsql = "select A.CRef_PacketNo,isnull(sum(A.CRef_BalQty),0)as BarQuantity,Isnull(sum(Convert(numeric(16,4),Issue_Qty)),0)as SalesQuantity "
            strsql = strsql & "from Bar_CrossReference A,Bar_Invoice_Issue B where A.CRef_PacketNo=substring(B.Issue_PartbarCode,9,len(CRef_PacketNo)) and A.UNIT_CODE=B.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "' AND "
            strsql = strsql & "A.CRef_PartCode=substring(B.Issue_partBarCode,1,8) and B.Issue_MISNO='" & Trim(pstrInvNo) & "' and Invoice_status=1 group by A.CRef_PacketNo"
            rsGetQty.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsGetQty.GetNoRows > 0 Then
                rsGetQty.MoveFirst()
                Do While Not rsGetQty.EOFRecord
                    If rsGetQty.GetValue("BarQuantity") = 0 Then
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "update A set A.CRef_BalQty=A.CRef_BalQty+" & rsGetQty.GetValue("SalesQuantity") & ","
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " A.CRef_Stage='B' from Bar_CrossReference A,Bar_Issue B"
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " Where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" & gstrUNITID & "' AND A.CRef_PartCode=substring(B.Issue_partBarCode,1,8) "
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " and B.Issue_MISNO='" & Trim(pstrInvNo) & "' and A.CRef_PacketNo='" & rsGetQty.GetValue("CRef_PacketNo") & "'" & vbCrLf
                    Else
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "update A set A.CRef_BalQty=A.CRef_BalQty+" & rsGetQty.GetValue("SalesQuantity") & ""
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " from Bar_CrossReference A,Bar_Issue B"
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " Where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" & gstrUNITID & "' AND A.CRef_PartCode=substring(B.Issue_partBarCode,1,8) "
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " and B.Issue_MISNO='" & Trim(pstrInvNo) & "' and A.CRef_PacketNo='" & rsGetQty.GetValue("CRef_PacketNo") & "'" & vbCrLf
                    End If
                    rsGetQty.MoveNext()
                Loop
                mstrupdateBarBondedStockFlag = mstrupdateBarBondedStockFlag & "Update Bar_Invoice_Issue Set Invoice_Status=0 where UNIT_CODE='" & gstrUNITID & "' AND Issue_MisNo='" & Trim(pstrInvNo) & "'" & vbCrLf
                mstrupdateBarBondedStockFlag = mstrupdateBarBondedStockFlag & "Delete from Bar_Issue where UNIT_CODE='" & gstrUNITID & "' AND Issue_MisNo='" & Trim(pstrInvNo) & "'" & vbCrLf
                BarCodeTracking = True
            End If
            rsGetQty = Nothing
        Else
            strsql = "select B.Box_label,isnull(sum(A.Quantity),0)as BarQuantity,isnull(sum(B.Quantity),0)as SalesQuantity from Bar_BondedStock A,Bar_BondedStock_Dtl B where "
            strsql = strsql & "A.Box_Label=B.Box_label AND A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" & gstrUNITID & "' and B.Invoice_No='" & Trim(pstrInvNo) & "' and B.Status_Flag='L' Group By B.Box_label"
            rsGetQty.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsGetQty.GetNoRows > 0 Then
                rsGetQty.MoveFirst()
                Do While Not rsGetQty.EOFRecord
                    If rsGetQty.GetValue("BarQuantity") = 0 Then
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "update A set A.Quantity=isnull(A.Quantity,0)+" & rsGetQty.GetValue("SalesQuantity") & ",A.Status='B'"
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " from bar_BondedStock A,bar_BondedStock_Dtl B"
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " Where A.Box_label = B.Box_label AND A.UNIT_CODE = B.UNIT_CODE AND A.UNIT_CODE='" & gstrUNITID & "' and B.Status_Flag='L'"
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " and B.Invoice_no='" & Trim(pstrInvNo) & "' and A.Box_label='" & rsGetQty.GetValue("Box_label") & "'" & vbCrLf
                    Else
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "update A set A.Quantity=isnull(A.Quantity,0)+" & rsGetQty.GetValue("SalesQuantity") & ""
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " from bar_BondedStock A,bar_BondedStock_Dtl B"
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " Where A.Box_label = B.Box_label AND A.UNIT_CODE = B.UNIT_CODE AND A.UNIT_CODE='" & gstrUNITID & "' and B.Status_Flag='L'"
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " and B.Invoice_no='" & Trim(pstrInvNo) & "' and A.Box_label='" & rsGetQty.GetValue("Box_label") & "'" & vbCrLf
                    End If
                    rsGetQty.MoveNext()
                Loop
                mstrupdateBarBondedStockFlag = "Update bar_BondedStock_Dtl set Status_Flag='C' where UNIT_CODE='" & gstrUNITID & "' AND Invoice_no='" & Trim(pstrInvNo) & "' and Status_Flag='L'"
                BarCodeTracking = True
            End If
            rsGetQty.ResultSetClose()
            rsGetQty = Nothing
        End If
        Exit Function
ErrHandler:
        BarCodeTracking = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function InvoiceAgainstBatch() As Boolean
        '----------------------------------------------------------------------------
        'Author         :   Manoj Kr. Vaish
        'Argument       :   Nil
        'Return Value   :   True or False
        'Function       :   Check Batch Tracking Functionality for Invoice
        'Comments       :   Date: 01 Oct 2008 ,Issue Id:eMpro-20080930-22159
        '----------------------------------------------------------------------------
        Dim rsGetRecord As ClsResultSetDB
        Dim strsql As String
        On Error GoTo ErrHandler
        rsGetRecord = New ClsResultSetDB
        strsql = "select isnull(Batch_Tracking,0) as Batch_Tracking from sales_parameter where UNIT_CODE='" & gstrUNITID & "'"
        rsGetRecord.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsGetRecord.GetNoRows > 0 Then
            InvoiceAgainstBatch = rsGetRecord.GetValue("Batch_Tracking")
        End If
        rsGetRecord.ResultSetClose()
        rsGetRecord = Nothing
        Exit Function
ErrHandler:
        InvoiceAgainstBatch = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function UpdateItemBatchDetail() As Boolean
        '----------------------------------------------------------------------------
        'Author         :   Manoj Kr. Vaish
        'Argument       :   Nil
        'Return Value   :   True or False
        'Function       :   Update the Itembatch_Mst and detail while invoice cancellation
        'Comments       :   Date: 01 Oct 2008 ,Issue Id:eMpro-20080930-22159
        '----------------------------------------------------------------------------
        Dim rsgetItemCode As ClsResultSetDB
        Dim strupdateItemBatchDtl As String
        Dim strupdateItemBatchMst As String
        Dim strsql As String
        Dim strStockLocation As String
        On Error GoTo ErrHandler
        UpdateItemBatchDetail = True
        strupdateItemBatchDtl = ""
        strupdateItemBatchMst = ""
        'If UCase(Trim(GetPlantName)) = "HILEX" And CHECK_FTSINVOICE() = True Then
        '    ''MsgBox("Invoice is made for FTS Items , cannot be Cancelled.", MsgBoxStyle.Exclamation, ResolveResString(100))
        '    'If FTS_INVOICECANCELLATION() = False Then
        '    '    mP_Connection.RollbackTrans()
        '    '    Exit Function
        '    'End If
        '    strStockLocation = "01P2"
        'Else
        '    strStockLocation = StockLocationSalesConf(mstrInvType, mstrInvSubType, "TYPE")
        'End If
        strStockLocation = StockLocationSalesConf(mstrInvType, mstrInvSubType, "TYPE")
        strsql = "select Item_Code,Batch_no,Batch_Qty from Itembatch_dtl where UNIT_CODE='" & gstrUNITID & "' and doc_no='" & Trim(txtChallanNo.Text) & "' and Doc_Type='9999'"
        rsgetItemCode = New ClsResultSetDB
        rsgetItemCode.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsgetItemCode.GetNoRows > 0 Then
            rsgetItemCode.MoveFirst()
            Do While Not rsgetItemCode.EOFRecord
                strupdateItemBatchDtl = "Update ItemBatch_Dtl set cancel_flag=1 ,upd_userid='" & mP_User & "',upd_dt=getdate() where UNIT_CODE='" & gstrUNITID & "' and doc_no='" & Trim(txtChallanNo.Text) & "' and doc_type='9999'"
                strupdateItemBatchMst = strupdateItemBatchMst & " Update ItemBatch_Mst set Current_Batch_Qty=Current_Batch_Qty+" & rsgetItemCode.GetValue("Batch_Qty")
                strupdateItemBatchMst = strupdateItemBatchMst & " where UNIT_CODE='" & gstrUNITID & "' and Item_code='" & rsgetItemCode.GetValue("Item_Code") & "' and Batch_no='" & rsgetItemCode.GetValue("Batch_no") & "'"
                strupdateItemBatchMst = strupdateItemBatchMst & " and Location_Code='" & strStockLocation & "'"
                rsgetItemCode.MoveNext()
            Loop
        End If
        If Len(strupdateItemBatchDtl) > 0 Then mP_Connection.Execute(strupdateItemBatchDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        If Len(strupdateItemBatchMst) > 0 Then mP_Connection.Execute(strupdateItemBatchMst, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        rsgetItemCode.ResultSetClose()
        rsgetItemCode = Nothing
        Exit Function
ErrHandler:
        UpdateItemBatchDetail = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function UpdatGRN_TradingInvoice() As Boolean
        Dim rssaledtl As ClsResultSetDB
        Dim strsql As String
        On Error GoTo Err_Handler
        UpdatGRN_TradingInvoice = False
        strsql = "UPDATE A SET DESPATCH_QTY_TRADING=ISNULL(DESPATCH_QTY_TRADING,0)-B.KNOCKOFFQTY"
        strsql = strsql + " FROM GRN_DTL A,"
        strsql = strsql + " (	"
        strsql = strsql + " SELECT GRIN_NO,KNOCKOFFQTY,GRIN_DOC_TYPE,ITEM_CODE,UNIT_CODE "
        strsql = strsql + " FROM SALES_TRADING_GRIN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND DOC_NO=" + Val(txtChallanNo.Text).ToString
        strsql = strsql + ") B "
        strsql = strsql + " WHERE A.DOC_NO=B.GRIN_NO"
        strsql = strsql + " AND A.DOC_TYPE=B.GRIN_DOC_TYPE"
        strsql = strsql + " AND A.ITEM_CODE=B.ITEM_CODE  "
        strsql = strsql + " AND A.UNIT_CODE=B.UNIT_CODE  "
        strsql = strsql + " AND A.UNIT_CODE='" + gstrUNITID + "'"
        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        UpdatGRN_TradingInvoice = True
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        UpdatGRN_TradingInvoice = False = False
    End Function
    Private Function ValidateInvoiceStockLocation() As Boolean
        '10804443 - MULTI LOCATION IN BARCODE - HILEX 
        Dim strMsg As String = String.Empty
        Try
            Using sqlcmd As SqlCommand = New SqlCommand
                With sqlcmd
                    If UCase(Trim(GetPlantName)) = "HILEX" And CHECK_FTSINVOICE("CHECK_FTS") = True Then
                        .CommandText = "USP_VALIDATE_INVOICE_LOCATION_CANCELLATION"
                    Else
                        .CommandText = "USP_VALIDATE_INVOICE_LOCATION"
                    End If

                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@INVOICE_NO", SqlDbType.BigInt).Value = Val(txtChallanNo.Text.Trim)
                    .Parameters.Add("@MSG", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
                    SqlConnectionclass.ExecuteNonQuery(sqlcmd)
                    strMsg = Convert.ToString(.Parameters("@MSG").Value)
                    If Len(strMsg) > 0 Then
                        MsgBox(strMsg, MsgBoxStyle.Exclamation, ResolveResString(100))
                        Return False
                    End If
                End With
            End Using
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function INVOICE_CHECK_IN_OUTWARD() As Boolean
        'Created By     : ABHIJIT KUMAR SINGH
        'Revised On     : 26 JULY 2017
        'Reason         : Invoice must not get cancelled if GATE Outward is done.
        '--------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strQry As String
        Dim Rs As ClsResultSetDB
        INVOICE_CHECK_IN_OUTWARD = False

        strQry = "select againstdocumentnumber from Gate_Outward_Reg_Hdr where againstdocumentnumber='" & txtChallanNo.Text & "' and UNIT_CODE='" & gstrUNITID & "'  and isnull(Cancel_YN,'N')='N' and againstdocument='INVOICE'"

        Rs = New ClsResultSetDB
        Rs.GetResult(strQry)

        If Rs.GetNoRows > 0 Then
            INVOICE_CHECK_IN_OUTWARD = True
        Else
            INVOICE_CHECK_IN_OUTWARD = False
        End If
        Rs.ResultSetClose()
        Rs = Nothing
        Exit Function
ErrHandler:
        Rs = Nothing
    End Function

    Private Function INVOICE_CHECK_IN_FRIEGT_OUTWARD() As Boolean
        'Created By     : ABHIJIT KUMAR SINGH
        'Revised On     : 26 JULY 2017
        'Reason         : Invoice must not get cancelled if GATE Outward is done.
        '--------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strQry As String
        Dim Rs As ClsResultSetDB

        INVOICE_CHECK_IN_FRIEGT_OUTWARD = False

        strQry = "select b.Trip_Doc_No from Freight_Gate_Outward_Reg_Hdr a inner join  Freight_Gate_Outward_Trip_Doc_Dtl b on a.Doc_No=b.Doc_No " &
               "and a.UNIT_CODE=b.UNIT_CODE where b.Trip_Doc_No='" & txtChallanNo.Text & "' and a.UNIT_CODE='" & gstrUNITID & "' and " &
               "b.trip_doc_type='invoice' and isnull(a.Cancel_YN,'N')='N'"

        Rs = New ClsResultSetDB
        Rs.GetResult(strQry)
        If Rs.GetNoRows > 0 Then
            INVOICE_CHECK_IN_FRIEGT_OUTWARD = True
        Else
            INVOICE_CHECK_IN_FRIEGT_OUTWARD = False
        End If

        Rs.ResultSetClose()
        Rs = Nothing
        Exit Function
ErrHandler:
        Rs = Nothing
    End Function
    Private Function FTS_INVOICECANCELLATION() As Boolean
        'HILEX VALIDATION :FTS INVOICE CANCELLED .. LABEL WILL BE ROLLBACK

        Dim strMsg As String = String.Empty
        Dim Com As New ADODB.Command
        Try

            With Com
                .ActiveConnection = mP_Connection
                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                .CommandText = "USP_FTS_INVOICECANCEL_STOCK_ADJUST"
                .CommandTimeout = 0
                .Parameters.Append(.CreateParameter("@INVOICE_NO", ADODB.DataTypeEnum.adBigInt, ADODB.ParameterDirectionEnum.adParamInput, , Val(txtChallanNo.Text.Trim)))
                .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrUNITID))
                .Parameters.Append(.CreateParameter("@MESSAGE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 3000))

                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                strMsg = Convert.ToString(.Parameters("@MESSAGE").Value)
                If Len(strMsg) > 0 Then
                    MsgBox(strMsg, MsgBoxStyle.Exclamation, ResolveResString(100))
                    Return False
                End If
            End With
            Return True
        Catch ex As Exception
            Throw ex
        End Try

    End Function
    Public Function CheckIRNCancel() As Boolean
        Try
            Dim blnEway As Boolean = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("Select ISNULL(EWAY_BILL_FUNCTIONALITY,0) EWAY_BILL_FUNCTIONALITY From SaleConf (Nolock) Where Unit_code='" & gstrUNITID & "' and Invoice_Type='" & mstrInvType & "' and Sub_Type='" & mstrInvSubType & "' and datediff(dd,'" & Convert.ToDateTime(lblDateDes.Text).ToString("dd MMM yyyy") & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & Convert.ToDateTime(lblDateDes.Text).ToString("dd MMM yyyy") & "')<=0"))
            If blnEway Then
                Dim strEway As String = Convert.ToString(SqlConnectionclass.ExecuteScalar("Select ISNULL(S.EWAY_IRN_REQUIRED,'') EWAY_IRN_REQUIRED From SalesChallan_Dtl S (Nolock) Where S.Unit_code='" & gstrUNITID & "' and S.Doc_No='" & txtChallanNo.Text & "'"))
                If strEway.ToUpper() = "I" Or strEway.ToUpper() = "B" Then
                    Dim strDeactivateDate As String = Convert.ToString(SqlConnectionclass.ExecuteScalar("Select ISNULL(S.IRN_DEACTIVATE_DATE,'') IRN_DEACTIVATE_DATE From SALESCHALLAN_DTL_IRN S (Nolock) Where S.Unit_code='" & gstrUNITID & "' and S.Doc_No='" & txtChallanNo.Text & "' and ISNULL(S.IRN_Deactivate,0)=1"))
                    If strDeactivateDate.Length = 0 Then
                        blnEway = True
                    Else
                        blnEway = False
                    End If
                Else
                    blnEway = False
                End If
            End If
            Return blnEway
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function
End Class