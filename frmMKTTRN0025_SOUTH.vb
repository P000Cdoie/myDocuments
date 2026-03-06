Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.IO
Friend Class frmMKTTRN0025_SOUTH
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
	'===============================================================================================================
	' Revise by         :   Davinder Singh
	' Revision date     :   13-Feb-2007
	' Revision history  :   Schedules are Reversed by using Stored Procedures
	'                       Barcode related tables are updated during Cancellation of invoice
	'===============================================================================================================
	'Revised By      : Manoj Kr. Vaish
	'Issue ID        : 22303
	'Revision Date   : 04 Feb 2008
	'History         : Addition of bar Code Functionaltiy for Chennai
    '*********************************************************************************
    'Revised By      : Manoj Kr.Vaish
    'Issue ID        : eMpro-20090209-27201
    'Revision Date   : 23 Mar 2009
    'History         : BatchWise Tracking of Invoices Made from 01M1 Location including BarCode Tracking
    '*******************************************************************************************************
    'Revised By      : Manoj Kr.Vaish
    'Issue ID        : eMpro-20090513-31282
    'Revision Date   : 18 May 2009
    'History         : Intergeration of Ford ASN File Generation for Mate South Units
    '*******************************************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : eMpro-20090709-33409
    'Revision Date   : 09 Jul 2009
    'History         : Cummulative Qunatity mismatch problem in FORD ASN Invoice for Mate 1,2 and 4
    '****************************************************************************************
    'Revised By        : Siddharth Ranjan
    'Issue ID          : eMpro-20100108-40881
    'Revision Date     : 09 Dec 2009
    'History           : New CSM FIFO KnockedOff functionality
    'Modified By Nitin Mehta on 17 May 2011
    'Modified to support MultiUnit functionality
    '****************************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Revision On     : 16 Nov 2011
    'Issue ID        : 10160094   
    'History         : Changes for ASN Path for Multi-Unit 
    '***********************************************************************************
    'Modified By Roshan Singh on 19 Dec 2011 for multiUnit change management    
    '****************************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Revision On     : 25 may 2012
    'Issue ID        : 10229016 
    'History         : Changes for invoice cancellation should be configurable 
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Revision On     : 06 Aug 2012
    'Issue ID        : 10259691 
    'History         : Changes in invoice cancellation 
    '***********************************************************************************
    'Revised By      : VINOD SINGH
    'Revision On     : 06 OCT 2014
    'Issue ID        : 10683802 
    'History         : "COMPONENTS" type invoices considereed for barcode invoice cancellation
    '*******************************************************************************************
    'Revised By         :   Vinod Singh
    'Revision Date      :   13 Jan 2015
    'Issue id           :   10736222 - eMPro - CT2 - ARE3 functionality
    '*******************************************************************************************************
    'REVISED BY     -  ASHISH SHARMA
    'REVISED ON     -  10 SEP 2018
    'PURPOSE        -  101375632 - REG Bar code implementation - BM1 UNIT
    '*******************************************************************************************************


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
    Dim mSchTypeArr() As String
    Dim mstrupdateBarBondedStockFlag As String
    Dim mstrupdateBarBondedStockQty As String
    Dim mstrConsigneeCode As String
    Dim mblnBatckTrackingAllowed As Boolean
    Dim mstrupdateASNdtl As String
    Dim mblnINVOICECANCELLATION_CURRENTDATE As Boolean
    Dim mblnRejTracking As Boolean = False
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
        'Revised By      : Manoj Kr. Vaish
        'Issue ID        : 22303
        'Revision Date   : 07 Feb 2008
        'History         : Check the Billflag while displaying the help
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
        ' issue id  : 10229016
        If mblnINVOICECANCELLATION_CURRENTDATE = True Then
            strInvoiceConditionDate = " AND INVOICE_DATE =CONVERT(char(12), GETDATE(), 106)"
        Else
            strInvoiceConditionDate = ""
        End If
        If Len(Trim(txtChallanNo.Text)) = 0 Then
            strHelpString = ShowList(1, (txtChallanNo.MaxLength), "", "Doc_No", DateColumnNameInShowList("Invoice_Date") & " as Invoice_Date ", "SalesChallan_Dtl ", "AND Location_Code='" & Trim(txtLocationCode.Text) & "' AND Cancel_Flag=0 and bill_flag=1 " & strInvoiceConditionDate)
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
            strHelpString = ShowList(1, (txtChallanNo.MaxLength), txtChallanNo.Text, "Doc_No", DateColumnNameInShowList("Invoice_Date") & " as Invoice_Date ", "SalesChallan_Dtl ", "AND Location_Code='" & Trim(txtLocationCode.Text) & "' AND Cancel_Flag=0 and Bill_flag=1 and invoice_date=getdate()")
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
        Dim StrItemCode As String 'Item Code
        Dim dbl_Quantity As Double 'Qty
        Dim mCust_Item_Code As String 'Drwg No
        Dim strupdatecustodtdtl As String 'SQL
        UpDateSalesOrder = True
        With SpChEntry
            For intCount = 1 To .MaxRows
                .Col = 1 : .Col2 = 1 : .Row = intCount : .Row2 = intCount : StrItemCode = Trim(.Text)
                .Col = 2 : .Col2 = 2 : .Row = intCount : .Row2 = intCount : mCust_Item_Code = Trim(.Text)
                .Col = 5 : .Col2 = 5 : .Row = intCount : .Row2 = intCount : dbl_Quantity = Val(.Text)
                'Make Updating Query
                strupdatecustodtdtl = strupdatecustodtdtl & "Update Cust_ord_dtl set Despatch_Qty = Despatch_Qty - "
                strupdatecustodtdtl = strupdatecustodtdtl & dbl_Quantity & " ,Active_Flag = 'A' where Account_code ='"
                strupdatecustodtdtl = strupdatecustodtdtl & Trim(txtCustCode.Text) & "' AND Item_Code = '" & StrItemCode & "'  And Cust_DrgNo = '"
                strupdatecustodtdtl = strupdatecustodtdtl & mCust_Item_Code & "' and Cust_ref = '" & Trim(txtRefNo.Text)
                strupdatecustodtdtl = strupdatecustodtdtl & "' And Amendment_No = '" & Trim(txtAmendNo.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'"
            Next
            If Len(strupdatecustodtdtl) > 0 Then mP_Connection.Execute(strupdatecustodtdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        End With
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        UpDateSalesOrder = False
    End Function
    Private Sub CmdGrpChEnt_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.cmdGrpAuthorise.ButtonClickEventArgs) Handles CmdGrpChEnt.ButtonClick
        On Error GoTo ErrHandler
        Dim StrItemCode As String
        Dim strDrgNo As String
        Dim strDSNo As String
        Dim rsDSTracking As New ClsResultSetDB
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim intRowCount As Short
        Dim strDespatchAdvise As String
        Dim strRemoveInvFromLoadingSlip As String

        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_AUTHORIZE 'CANCELLATION OF INVOICE

                ' Added by ABHIJIT on 27-July-2017
                If INVOICE_CHECK_IN_FRIEGT_OUTWARD() = True Then
                    MsgBox("This Invoice No. Can't be cancelled.." & vbNewLine & "Ourward Entry has been already done.", MsgBoxStyle.Information, "empower")
                    Exit Sub
                End If

                If INVOICE_CHECK_IN_OUTWARD() = True Then
                    MsgBox("This Invoice No. Can't be cancelled.." & vbNewLine & "Ourward Entry has been already done.", MsgBoxStyle.Information, "empower")
                    Exit Sub
                End If
                ' Added by ABHIJIT on 27-July-2017

                If CheckIRNCancel() Then
                    MsgBox("This Invoice No. Can't be cancelled.." & vbNewLine & "Please first cancel IRN No.", MsgBoxStyle.Information, "eMPower")
                    Exit Sub
                End If

                If Trim(txtremark.Text) = "" Then
                    MsgBox("Cancel Remark is not present", MsgBoxStyle.Information, "empower")
                    If txtremark.Enabled Then txtremark.Focus()
                    Exit Sub
                End If
                If ConfirmWindow(10054, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    mP_Connection.BeginTrans()
                    If MBlnDsTracking Then
                        With SpChEntry
                            For intRowCount = 1 To .MaxRows 'For All Items In Spread Revise Schedule
                                .Col = 1 : .Col2 = 1 : .Row = intRowCount : .Row2 = intRowCount : StrItemCode = Trim(.Text)
                                .Col = 2 : .Col2 = 2 : .Row = intRowCount : .Row2 = intRowCount : strDrgNo = Trim(.Text)
                                rsDSTracking.GetResult("Select Item_code,Cust_Part_Code,DSno from Mkt_InvDSHistory where Doc_No = '" & txtChallanNo.Text & "' and Item_code = '" & StrItemCode & "' and Cust_Part_Code = '" & strDrgNo & "' AND UNIT_CODE='" & gstrUNITID & "'")
                                intMaxLoop = rsDSTracking.RowCount
                                rsDSTracking.MoveFirst()
                                strUpdateDailyMktSchedule = ""
                                strUpdateInvDsHistory = ""
                                strUpdateMonthlyMktSchedule = ""
                                For intLoopCounter = 1 To intMaxLoop
                                    strDSNo = rsDSTracking.GetValue("DSNo")
                                    StrItemCode = rsDSTracking.GetValue("Item_code")
                                    strDrgNo = rsDSTracking.GetValue("Cust_Part_code")
                                    If DispatchQuantityFromDailyScheduleForDSTracking(txtCustCode.Text, strDrgNo, StrItemCode, lblDateDes.Text, strDSNo) = True Then
                                        If Not UpdateDailyMktSchedule(StrItemCode, strDrgNo, strDSNo) Then mP_Connection.RollbackTrans() : Exit Sub
                                    Else
                                        If DispatchQuantityFromMonthlyScheduleForDSTracking(txtCustCode.Text, strDrgNo, StrItemCode, lblDateDes.Text, strDSNo) = True Then
                                            If Not UpdateMonthlyMktSchedule(StrItemCode, strDrgNo, strDSNo) Then mP_Connection.RollbackTrans() : Exit Sub
                                        Else
                                            MsgBox("No Schedule available for Item Code " & StrItemCode & " and DSNo " & strDSNo & ".", MsgBoxStyle.Information, "Empower")
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
                        If ReviseSchedule() = False Then mP_Connection.RollbackTrans() : Exit Sub
                    End If
                    '        If Not ReviseSchedule Then mP_Connection.RollbackTrans: Exit Sub
                    If Not UpDateSalesOrder() Then mP_Connection.RollbackTrans() : Exit Sub
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
                    If AllowASNTextFileGeneration(txtCustCode.Text.Trim()) = True Then
                        If FordASNFileGeneration(txtChallanNo.Text.Trim(), txtCustCode.Text.Trim()) = False Then
                            mP_Connection.RollbackTrans()
                            Exit Sub
                        Else
                            If Len(mstrupdateASNdtl) > 0 Then
                                mP_Connection.Execute(mstrupdateASNdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                        End If
                    End If
                    If Not updatesalesconfandsaleschallan() Then mP_Connection.RollbackTrans() : Exit Sub
                    If Not UpdateinSale_Dtl() Then mP_Connection.RollbackTrans() : Exit Sub
                    If UCase(mstrInvType) = "INV" And UCase(mstrInvSubType) = "F" Then
                        Dim blnCSM_Knockingoff_req As Boolean = DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE CSM_KNOCKINGOFF_REQ = 1 AND UNIT_CODE='" & gstrUNITID & "'")
                        If blnCSM_Knockingoff_req Then
                            If Not CANCEL_CSM_DETAILS() Then mP_Connection.RollbackTrans() : Exit Sub
                        End If
                    End If
                    If mblnBatckTrackingAllowed = True Then ''And InvoiceAgainstBatch() = True  InvoiceAgainstBatch commented by priti on 23 Jan to solve batch issue on rejection invoice
                        If Not UpdateItemBatchDetail() Then mP_Connection.RollbackTrans() : Exit Sub
                    End If
                    If Not UpdateDespAdvise() Then mP_Connection.RollbackTrans() : Exit Sub

                    'ADDED AGAINST ISSUE ID 10736222 
                    mP_Connection.Execute("EXEC USP_CANCEL_CT2_INVOICE '" & gstrUNITID & "'," & Val(txtChallanNo.Text.Trim) & "", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    'END OF ADDITION

                    If UpdateFinance() Then
                        ''Code Added by priti to cancel rejection invoice on 02 Dec 2025
                        If mblnRejTracking = True And UCase(mstrInvType) = "REJ" Then
                            mP_Connection.Execute("Update MKT_INVREJ_DTL Set Cancel_flag=1,Upd_DT=Getdate(), Upd_UserID='" & mP_User & "'" & " where UNIT_CODE='" & gstrUNITID & "' AND Invoice_No=" & Trim(txtChallanNo.Text), , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If

                        'Code Added By Shubhra To Remove Cancelled Invoice from ILVS
                        'Begin
                        strRemoveInvFromLoadingSlip = "Update Loadingslip set InvoiceNo = NULL, ACT_INV_NO = NULL" &
                            " where Unit_Code = '" & gstrUNITID & "' and ACT_INV_NO = " & Val(txtChallanNo.Text.Trim) & ""
                        mP_Connection.Execute(strRemoveInvFromLoadingSlip, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        'End

                        '101375632
                        If GetPaletteStatus(mstrInvType, mstrInvSubType, txtCustCode.Text.Trim()) Then
                            mP_Connection.Execute("INSERT INTO INVOICE_PALETTE_DTL_LOG (TEMP_INVOICE_NO,ITEM_CODE,PALETTE_LABEL,QTY,INVOICE_NO,IS_SCANNED,UNIT_CODE,ENT_DT,ENT_USERID,UPD_DT,UPD_USERID,ENT_DT_LOG,ENT_USERID_LOG) SELECT TEMP_INVOICE_NO,ITEM_CODE,PALETTE_LABEL,QTY,INVOICE_NO,IS_SCANNED,UNIT_CODE,ENT_DT,ENT_USERID,UPD_DT,UPD_USERID,GETDATE(),'" & mP_User & "' FROM INVOICE_PALETTE_DTL WHERE UNIT_CODE='" & gstrUNITID & "' AND INVOICE_NO=" & Val(txtChallanNo.Text.Trim) & "", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            mP_Connection.Execute("DELETE FROM INVOICE_PALETTE_DTL WHERE UNIT_CODE='" & gstrUNITID & "' AND INVOICE_NO=" & Val(txtChallanNo.Text.Trim) & "", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If

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
                Call frmMKTTRN0025_SOUTH_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Function DispatchQuantityFromMonthlySchedule(ByVal pstrAccountCode As String, ByVal pstrCustomerDrawNo As String, ByVal pstrItemCode As String, ByVal pstrDate As String, ByVal pstrMode As String, ByVal pdblPrevQty As Double) As Boolean
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
        Dim ldblTotalDispatchQuantity As Double
        Dim ldblTotalScheduleQuantity As Double
        Dim lintLoopCounter As Short
        Dim strMakeDate As String
        On Error GoTo ErrHandler
        ldblTotalDispatchQuantity = 0
        ldblTotalScheduleQuantity = 0
        mP_Connection.Execute("SET DATEFORMAT 'mdy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        strMakeDate = Year(ConvertToDate(pstrDate)) & Month(ConvertToDate(pstrDate)).ToString.PadLeft(2, "0")

        'If Val(CStr(Month(CDate(pstrDate)))) < 10 Then
        '    strMakeDate = Year(CDate(pstrDate)) & "0" & Month(CDate(pstrDate))
        'Else
        '    strMakeDate = Year(CDate(pstrDate)) & Month(CDate(pstrDate))
        'End If
        If objRsForSchedule.State = 1 Then objRsForSchedule.Close()
        objRsForSchedule.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        strScheduleSql = "Select Schedule_Qty,isnull(Despatch_Qty,0) AS Despatch_qty  from MonthlyMktSchedule where Account_Code='" & pstrAccountCode & "' and "
        strScheduleSql = strScheduleSql & " Year_Month =" & Val(Trim(strMakeDate)) & ""
        strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND Item_code = '" & pstrItemCode & "' and status =1 AND UNIT_CODE='" & gstrUNITID & "'"
        objRsForSchedule.Open(strScheduleSql, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        If objRsForSchedule.EOF Or objRsForSchedule.BOF Then
            DispatchQuantityFromMonthlySchedule = False
            mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            objRsForSchedule.Close()
            Exit Function
        Else
            objRsForSchedule.MoveFirst()
            For lintLoopCounter = 1 To objRsForSchedule.RecordCount
                'Get Total Schedule Qty , Dispatch Qty
                ldblTotalScheduleQuantity = ldblTotalScheduleQuantity + Val(objRsForSchedule.Fields("Schedule_Qty").Value)
                ldblTotalDispatchQuantity = ldblTotalDispatchQuantity + Val(objRsForSchedule.Fields("Despatch_Qty").Value)
                objRsForSchedule.MoveNext()
            Next
            DispatchQuantityFromMonthlySchedule = True 'Dispatch Against Monthly Schedule
            mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            objRsForSchedule.Close()
            Exit Function
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        DispatchQuantityFromMonthlySchedule = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function ReviseSchedule() As Boolean
        '****************************************************
        'Created By     -  Nitin Sood
        'Description    -  Decrease Dispatch Quantity for a Particular Item
        '                   from Marketing Schedule (Daily / Monthly)
        '
        '******************************************************************************************************
        'Revised by       - Davinder Singh
        'Revision date    - 13 Feb 2007
        'Revision History - To Revert the schedules by using Stored procedured on knockoff basis
        '******************************************************************************************************
        '''Dim dblNetDispatchQty As Double         'Net Dispatch Qty
        '''Dim intRowCount As Integer              'For Next Loop . . .
        '''Dim strDrgNo As String                  'Cust DRWG no.
        '''Dim StrItemCode As String               'Item Code
        '''Dim dblqty As Double                    'Qty
        '''Dim mstrUpdDispatchSql As String        'SQL
        Dim strCheckSchedule As String 'To Make Date in Case of MonthlySchedule
        On Error GoTo ErrHandler
        ReviseSchedule = True
        strCheckSchedule = CheckMktSchedules()
        If Trim(strCheckSchedule) = "Error" Then
            ReviseSchedule = False
            Exit Function
        End If
        If UpdateMktSchedules("-") = False Then
            ReviseSchedule = False
            Exit Function
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        ReviseSchedule = False
    End Function
    Private Function DispatchQuantityFromDailySchedule(ByVal pstrAccountCode As String, ByVal pstrCustomerDrawNo As String, ByVal pstrItemCode As String, ByVal pstrDate As String, ByVal pstrMode As String, ByVal pdblPrevQty As Double) As Boolean
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
        mP_Connection.Execute("SET DATEFORMAT 'mdy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        strScheduleSql = "Select Schedule_Quantity,Despatch_Qty from DailyMktSchedule where Account_Code='" & pstrAccountCode & "' and "
        strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(pstrDate)) & "'"
        strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(pstrDate)) & "'"
        strScheduleSql = strScheduleSql & " and Trans_Date <='" & getDateForDB(pstrDate) & "'"
        strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND ITEM_CODE  = '" & pstrItemCode & "' and Status =1 AND UNIT_CODE='" & gstrUNITID & "'  ORDER BY Trans_Date DESC" '''and Schedule_Flag =1   ( Now Not Consider)
        If objRsForSchedule.State = 1 Then objRsForSchedule.Close()
        objRsForSchedule.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        objRsForSchedule.Open(strScheduleSql, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        If objRsForSchedule.EOF Or objRsForSchedule.BOF Then
            DispatchQuantityFromDailySchedule = False
            mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            objRsForSchedule.Close()
            Exit Function
        Else
            objRsForSchedule.MoveFirst()
            For lintLoopCounter = 1 To objRsForSchedule.RecordCount
                ldblTotalScheduleQuantity = ldblTotalScheduleQuantity + Val(objRsForSchedule.Fields("Schedule_Quantity").Value)
                ldblTotalDispatchQuantity = ldblTotalDispatchQuantity + Val(objRsForSchedule.Fields("Despatch_Qty").Value)
                objRsForSchedule.MoveNext()
            Next
            If ldblTotalScheduleQuantity > 0 And ldblTotalDispatchQuantity > 0 Then
                DispatchQuantityFromDailySchedule = True
            End If
            mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            objRsForSchedule.Close()
            Exit Function
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        DispatchQuantityFromDailySchedule = False
        mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub Cmditems_Click()
        Dim frmMKTTRN0021 As Object
        '****************************************************
        'Created By     -  Nisha
        'Description    -  Display Another Form for User To Select Item Code >From CustOrd_Dtl
        '                  And After Selecting Item Code Select Data From Sales_Dtl and Display
        '                  That Details In The Spread
        '****************************************************
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
        With Me.SpChEntry
            .MaxRows = 1
            .Row = 1 : .Row2 = .MaxRows : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
        End With
        strStockLocation = StockLocationSalesConf(mstrInvType, mstrInvSubType, "TYPE")
        If Len(Trim(strStockLocation)) > 0 Then
            If (UCase(mstrInvType) = "INV") Or (UCase(mstrInvType) = "EXP") Then
                mstrItemCode = SelectItemfromsaleDtl(Trim(txtChallanNo.Text))
                If Len(Trim(mstrItemCode)) = 0 Then
                    SpChEntry.MaxRows = 0
                End If
            Else
                mstrItemCode = SelectItemfromsaleDtl(Trim(txtChallanNo.Text))
                If Len(Trim(mstrItemCode)) = 0 Then
                    SpChEntry.MaxRows = 0
                End If
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
            strHelp = ShowList(1, (txtLocationCode.MaxLength), "", "s.Location_Code", "l.Description", "Location_mst l,SaleConf s", "and s.Location_Code=l.Location_Code and s.UNIT_CODE = l.UNIT_CODE", , , , , , "s.UNIT_CODE")
            If strHelp = "-1" Then 'If No Record Exists In The Table
                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Exit Sub
            Else
                txtLocationCode.Text = strHelp
            End If
        Else
            'To Display All Possible Help Starting With Text in TextField
            strHelp = ShowList(1, (txtLocationCode.MaxLength), txtLocationCode.Text, "s.Location_Code", "l.Description", "Location_mst l,SaleConf s", "and s.Location_Code=l.Location_Code and s.UNIT_CODE = l.UNIT_CODE", , , , , , "s.UNIT_CODE")
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
        Call ShowHelp("HLPMKTTRN0025.HTM") ' HLPMKTTRN0005.htm
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0025_SOUTH_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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
    Private Sub frmMKTTRN0025_SOUTH_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
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
    Private Sub frmMKTTRN0025_SOUTH_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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
        CmdGrpChEnt.Picture(UCActXCtl.clsDeclares.ButtonEnabledEnum.NEW_BUTTON) = My.Resources.resEmpower.ico123.ToBitmap
        CmdGrpChEnt.Picture(UCActXCtl.clsDeclares.ButtonEnabledEnum.UPDATE_BUTTON) = My.Resources.resEmpower.ico121.ToBitmap
        'Set Help Pictures At Command Button
        CmdLocCodeHelp.Image = My.Resources.ico111.ToBitmap
        CmdChallanNo.Image = My.Resources.ico111.ToBitmap
        'Check If Company is 100% EOU then CVD SVD fields are SHOWN - NITIN SOOD
        gobjDB = New ClsResultSetDB
        If gobjDB.GetResult("Select EOU_FLAG From Company_Mst where UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic) Then
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
            .Row = 0 : .Col = 3 : .Text = "Rate"
            .Row = 0 : .Col = 4 : .Text = "Cust Material"
            .Row = 0 : .Col = 5 : .Text = "Quantity"
            .Row = 0 : .Col = 6 : .Text = "Packing(%)"
            .Row = 0 : .Col = 7 : .Text = "EXC(%)"
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
        Call addRowAtEnterKeyPress(1)
        lblRGPDes.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)

        mblnRejTracking = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("Select ISNULL(REJINV_Tracking,0) REJINV_Tracking From sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'"))

        '----------Checks the flag whether Schedule is DSwise or old procedure
        If RsobjSchedules.State = ADODB.ObjectStateEnum.adStateOpen Then RsobjSchedules.Close()
        RsobjSchedules.Open("SELECT DSWiseTracking FROM sales_parameter where UNIT_CODE='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not RsobjSchedules.EOF Then
            If IIf(IsDBNull(RsobjSchedules.Fields(0).Value), 1, 0) = 1 Then
                MBlnDsTracking = True
            Else
                MBlnDsTracking = False
            End If
        End If
        If RsobjSchedules.State = ADODB.ObjectStateEnum.adStateOpen Then RsobjSchedules.Close()
        RsobjSchedules = Nothing
        ' issue id  : 10229016
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
        ' issue id  : 10229016
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0025_SOUTH_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        CmdGrpChEnt.Enabled(1) = False
        mdifrmMain.CheckFormName = mintIndex
        txtLocationCode.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0025_SOUTH_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0025_SOUTH_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        On Error GoTo ErrHandler
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
            Me.Dispose()
        End If
        'Checking The Status
        If gblnCancelUnload = True Then Cancel = 1
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTTRN0025_SOUTH_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
                'Get Data From Challan_Dtl,Cust_Ord_Dtl,Sales_Dtl
                If Len(txtLocationCode.Text) > 0 Then
                    If GetDataInViewMode() Then 'if record found
                        Call CheckMultipleSOAllowed(lblInvoiceType.Text, lblInvoiceSubType.Text)
                        Call Cmditems_Click()
                        With CmdGrpChEnt
                            .Enabled(0) = True 'CANCEL
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
        If Len(Trim(pstrCondition)) > 0 Then
            strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and " & pstrCondition
        Else
            strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' AND cancel_flag=0 and invoice_date=CONVERT(char(12), GETDATE(), 106)"
        End If
        ' issue id  : 10229016
        If mblnINVOICECANCELLATION_CURRENTDATE = True Then
            strTableSql = strTableSql + " AND INVOICE_DATE =CONVERT(char(12), GETDATE(), 106)"
        End If
        ' issue id  : 10229016
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
        strSalesChallanDtl = strSalesChallanDtl & "Insurance,Invoice_Date,consignee_code,"
        strSalesChallanDtl = strSalesChallanDtl & "Invoice_Type,Sub_Category,Cust_Name,Carriage_Name,Frieght_Amount, "
        strSalesChallanDtl = strSalesChallanDtl & "Surcharge_salesTaxType,Amendment_No,ref_doc_no,Currency_Code,Exchange_Rate,Remarks,PerValue From Saleschallan_dtl WHERE Location_Code ='"
        strSalesChallanDtl = strSalesChallanDtl & Trim(txtLocationCode.Text) & "' and UNIT_CODE='" & gstrUNITID & "' and Doc_No = " & Val(txtChallanNo.Text)
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
            mstrConsigneeCode = rsGetData.GetValue("consignee_code")
            ''Commented by priti on 23 Jan 2026 to impact barcode issue in Raw material invoice
            'lblInvoiceSubType.Text = SelectDataFromTable("Sub_Type_Description", "Saleconf", " WHERE UNIT_CODE='" & gstrUNITID & "' and Sub_type='" & Trim(mstrInvSubType) & "'")
            lblInvoiceSubType.Text = SelectDataFromTable("Sub_Type_Description", "Saleconf", " WHERE UNIT_CODE='" & gstrUNITID & "' and Invoice_type='" & Trim(mstrInvType) & "' and Sub_type='" & Trim(mstrInvSubType) & "'")

            lblCurrency.Visible = True : lblCurrencyDes.Visible = True
            lblCurrencyDes.Text = rsGetData.GetValue("Currency_code")
            lblExchangeRateLable.Visible = True : lblExchangeRateValue.Visible = True
            ''Have to Add Remarks & Per Value
        Else
            GetDataInViewMode = False
        End If
        '***To Display invoice Address of Customer
        If Len(Trim(txtCustCode.Text)) > 0 Then
            rsCustMst = New ClsResultSetDB
            strCustMst = "SELECT Bill_Address1 + ', '  +  Bill_Address2 + ', ' + Bill_City + ' - ' + Bill_Pin as  invoiceAddress from Customer_Mst where Customer_code ='" & txtCustCode.Text & "' and UNIT_CODE='" & gstrUNITID & "'"
            rsCustMst.GetResult(strCustMst)
            If rsCustMst.GetNoRows > 0 Then
                lblAddressDes.Text = rsCustMst.GetValue("InvoiceAddress")
            End If
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
        Dim rsTariffMst As ClsResultSetDB
        Dim intDecimal As Short
        strsaledtl = ""
        strsaledtl = "SELECT * from Sales_Dtl WHERE Location_Code='" & Trim(txtLocationCode.Text) & "'"
        strsaledtl = strsaledtl & " and Doc_No=" & Val(txtChallanNo.Text) & " and Cust_Item_Code in(" & Trim(mstrItemCode) & ") and UNIT_CODE='" & gstrUNITID & "'"
        rsSalesDtl = New ClsResultSetDB
        rsSalesDtl.GetResult(strsaledtl, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim intloopcount As Short
        Dim varCumulative As Object
        If rsSalesDtl.GetNoRows > 0 Then
            intRecordCount = rsSalesDtl.GetNoRows
            ReDim mdblPrevQty(intRecordCount - 1) ' To get value of Quantity in Arrey for updation in despatch
            ReDim mdblToolCost(intRecordCount - 1) ' To get value of Quantity i
            Call addRowAtEnterKeyPress(intRecordCount - 1)
            rsSalesDtl.MoveFirst()
            If Trim(UCase(lblInvoiceType.Text)) = "NORMAL INVOICE" Or Trim(UCase(lblInvoiceType.Text)) = "JOBWORK INVOICE" Or Trim(UCase(lblInvoiceType.Text)) = "EXPORT INVOICE" Then
                If UCase(Trim(lblInvoiceSubType.Text)) <> "SCRAP" Then
                    For intloopcount = 1 To intRecordCount
                        mdblToolCost(intloopcount - 1) = Val(rsSalesDtl.GetValue("Tool_Cost"))
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
                lblInvoiceType.Text = "" : lblInvoiceSubType.Text = ""
                Me.CmdGrpChEnt.Enabled(1) = False
                Me.CmdGrpChEnt.Enabled(2) = False
            Case "CHALLAN"
                txtChallanNo.Text = "" : txtCustCode.Text = "" : lblCustCodeDes.Text = "" : lblAddressDes.Text = ""
                txtCarrServices.Text = "" : txtVehNo.Text = ""
                txtFreight.Text = "" : txtSaleTaxType.Text = "" : lblSaltax_Per.Text = "0.00"
                txtSurchargeTaxType.Text = "" : lblSurcharge_Per.Text = "0.00"
                ctlInsurance.Text = "" : lblRGPDes.Text = ""
                lblInvoiceType.Text = "" : lblInvoiceSubType.Text = "" : lblCurrencyDes.Text = "" : txtRefNo.Text = ""
                Me.CmdGrpChEnt.Enabled(1) = False
                Me.CmdGrpChEnt.Enabled(2) = False
        End Select
        With Me.SpChEntry
            .MaxRows = 1
            .Row = 1 : .Row2 = 1 : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
        End With
        txtremark.Enabled = False
        txtremark.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
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
                rsSalesConf.GetResult("Select Stock_Location from SaleConf Where Description ='" & Trim(pstrInvType) & "' and Sub_type_Description ='" & Trim(pstrInvSubtype) & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate()) and UNIT_CODE='" & gstrUNITID & "'")
            Case "TYPE"
                rsSalesConf.GetResult("Select Stock_Location from SaleConf Where Invoice_type ='" & Trim(pstrInvType) & "' and Sub_type ='" & Trim(pstrInvSubtype) & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate()) and UNIT_CODE='" & gstrUNITID & "'")
        End Select
        If rsSalesConf.GetNoRows > 0 Then
            StockLocation = rsSalesConf.GetValue("Stock_Location")
        End If
        StockLocationSalesConf = StockLocation
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    '    Private Function GetServerDate() As Date
    '        Dim objServerDate As ClsResultSetDB 'Class Object
    '        Dim strSql As String 'Stores the SQL statement
    '        On Error GoTo ErrHandler
    '        'Build the SQL statement
    '        strSql = "SELECT CONVERT(datetime,getdate(),103)"
    '        'Creating the instance
    '        objServerDate = New ClsResultSetDB
    '        With objServerDate
    '            'Open the recordset
    '            Call .GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '            'If we have a record, then getting the financial year else exiting
    '            If .GetNoRows <= 0 Then Exit Function
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
        rscurrency.GetResult("Select Decimal_Place from Currency_Mst where Currency_code ='" & pstrCurrency & "' and UNIT_CODE='" & gstrUNITID & "'")
        ToGetDecimalPlaces = Val(rscurrency.GetValue("Decimal_Place"))
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
            strTableSql = "SELECT CExch_MultiFactor FROM Gen_CurExchMaster WHERE UNIT_CODE='" & gstrUNITID & "' and CExch_CurrencyTo='" & Trim(pstrCurrencyCode) & "' AND CExch_InOut=1 AND '" & Trim(pstrDate) & "' BETWEEN CExch_DateFrom AND CExch_DateTo"
        Else
            strTableSql = "SELECT CExch_MultiFactor FROM Gen_CurExchMaster WHERE UNIT_CODE='" & gstrUNITID & "' and CExch_CurrencyTo='" & Trim(pstrCurrencyCode) & "' AND CExch_InOut=0 AND '" & Trim(pstrDate) & "' BETWEEN CExch_DateFrom AND CExch_DateTo"
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
    Private Sub txtRemark_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtremark.KeyPress
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
            If CheckExistanceOfFieldData((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST') and UNIT_CODE='" & gstrUNITID & "'") Then
                lblSaltax_Per.Text = CStr(GetTaxRate((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST') and UNIT_CODE='" & gstrUNITID & "'"))
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
        Dim rsInvType As ClsResultSetDB
        Dim intRecordCount As Short 'To Hold Record Count
        Dim intCount As Short
        Dim strItemText As String
        'changed due to more then 4 item selection in case of Export in voice.
        '****************************
        strsaledtl = ""
        strsaledtl = "Select a.Item_Code,a.Cust_ITem_Code,a.Cust_Item_Desc,b.Tariff_Code from Sales_Dtl a,Item_Mst b where a.ITem_code = b.ITem_code and a.UNIT_CODE=b.UNIT_CODE and a.UNIT_CODE='" & gstrUNITID & "' and Doc_No ="
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
        Dim strSql As String
        Dim rsSalesChallan As ClsResultSetDB
        Dim dblInvoiceAmt As Double
        Dim saleschallan As String
        Dim salesconf As String
        On Error GoTo Err_Handler
        updatesalesconfandsaleschallan = True
        strSql = "select *  from Saleschallan_dtl where UNIT_CODE='" & gstrUNITID & "' AND Doc_No = " & Trim(txtChallanNo.Text)
        strSql = strSql & " and Invoice_type = '" & mstrInvType & "'  and  sub_category =  '" & mstrInvSubType & "' and Location_Code='" & Trim(txtLocationCode.Text) & "'"
        rsSalesChallan = New ClsResultSetDB
        rsSalesChallan.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
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
        Dim StrItemCode As String
        Dim dblQty As Double
        Dim intloopcount As Short
        Dim strupdateGrinhdr As String
        On Error GoTo ErrHandler
        UpdateGrnHdr = True
        rsSalesDtl = New ClsResultSetDB
        rsSalesDtl.GetResult("select * from sales_dtl where UNIT_CODE='" & gstrUNITID & "' AND Doc_No = " & pdblinvoiceNo & " and Location_Code='" & Trim(txtLocationCode.Text) & "'")
        If rsSalesDtl.GetNoRows > 0 Then
            intMaxLoop = rsSalesDtl.GetNoRows
            rsSalesDtl.MoveFirst()
            strupdateGrinhdr = ""
            For intloopcount = 1 To intMaxLoop
                StrItemCode = rsSalesDtl.GetValue("ITem_code")
                dblQty = rsSalesDtl.GetValue("Sales_Quantity")
                If Len(Trim(strupdateGrinhdr)) = 0 Then
                    strupdateGrinhdr = "Update Grn_Dtl Set Despatch_Quantity = isnull(Despatch_Quantity,0) -" & dblQty
                    strupdateGrinhdr = strupdateGrinhdr & " Where UNIT_CODE='" & gstrUNITID & "' AND ITem_Code = '" & StrItemCode & "' and Doc_No = " & pdblGrinNo
                Else
                    strupdateGrinhdr = strupdateGrinhdr & vbCrLf & "Update Grn_Dtl Set Despatch_Quantity = isnull(Despatch_Quantity,0) - " & dblQty
                    strupdateGrinhdr = strupdateGrinhdr & " Where UNIT_CODE='" & gstrUNITID & "' AND ITem_Code = '" & StrItemCode & "' and Doc_No = " & pdblGrinNo
                End If
                rsSalesDtl.MoveNext()
            Next
            If Len(strupdateGrinhdr) > 0 Then mP_Connection.Execute(strupdateGrinhdr, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Else
            MsgBox("No Items Available in Invoice " & txtChallanNo.Text)
            UpdateGrnHdr = False
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        UpdateGrnHdr = False
    End Function
    Private Function CANCEL_CSM_DETAILS() As Boolean
        On Error GoTo Err_Handler
        CANCEL_CSM_DETAILS = True
        Dim blnCSM_Knockingoff_req As Boolean = DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE UNIT_CODE='" & gstrUNITID & "' AND CSM_KNOCKINGOFF_REQ = 1")
        If blnCSM_Knockingoff_req Then
            Dim objComm As New ADODB.Command
            With objComm
                .ActiveConnection = mP_Connection
                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                .CommandText = "USP_CANCEL_CSM_DETAILS"
                .CommandTimeout = 0
                .Parameters.Append(.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                .Parameters.Append(.CreateParameter("@INV_NO", ADODB.DataTypeEnum.adBigInt, ADODB.ParameterDirectionEnum.adParamInput, , txtChallanNo.Text.Trim))
                .Parameters.Append(.CreateParameter("@RETURN", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, , 0))
                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                If .Parameters(.Parameters.Count - 1).Value <> 0 Then
                    MsgBox("Unable To Cancel CSM Knocking Off Details.", MsgBoxStyle.Information, ResolveResString(100))
                    CANCEL_CSM_DETAILS = False
                    Exit Function
                End If
            End With
            objComm = Nothing
        End If
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        CANCEL_CSM_DETAILS = False
    End Function
    Private Function UpdateinSale_Dtl() As Boolean
        Dim rssaledtl As ClsResultSetDB
        Dim strupdateitbalmst As String
        Dim strSql As String
        Dim strStockLocCode As String
        Dim intRow, intloopcount As Short
        Dim mItem_Code As String
        Dim mSales_Quantity As Double
        strupdateitbalmst = ""
        On Error GoTo Err_Handler
        UpdateinSale_Dtl = True
        strStockLocCode = StockLocationSalesConf(mstrInvType, mstrInvSubType, "TYPE")
        strSql = "Select * from sales_Dtl where UNIT_CODE='" & gstrUNITID & "' AND Doc_No = " & Trim(txtChallanNo.Text) & " and Location_Code='" & Trim(txtLocationCode.Text) & "'"
        rssaledtl = New ClsResultSetDB
        rssaledtl.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rssaledtl.GetNoRows > 0 Then
            intRow = rssaledtl.GetNoRows
            rssaledtl.MoveFirst()
            For intloopcount = 1 To intRow
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
        On Error GoTo ErrHandler
        strRemark = QuoteRem(Trim(txtremark.Text))
        'For Cancellation of Dr Note raised against Rejection Invoice
        If ChkForRejInv() Then
            strResult = objDrCrNote.ReverseAPDocument(gstrUNITID, mstrInvNo, mP_User, getDateForDB(lblDateDes.Text), strRemark, gstrCURRENCYCODE, , gstrCONNECTIONSTRING)
        Else
            strResult = objDrCrNote.ReverseARInvoiceDocument(gstrUNITID, Trim(txtChallanNo.Text), mP_User, getDateForDB(lblDateDes.Text), strRemark, gstrCURRENCYCODE, , gstrCONNECTIONSTRING)
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
        Dim strSql As String
        Dim lintLoop As Short
        Dim StrItemCode As String
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
            StrItemCode = Trim(SpChEntry.Text)
            strSql = "SELECT Quantity,Ref57F4_no,Customer_code FROM CustAnnex_Dtl WHERE UNIT_CODE='" & gstrUNITID & "' AND Invoice_No = " & Trim(txtChallanNo.Text) & " and Item_Code='" & Trim(StrItemCode) & "'"
            rsCustAnnexDtl.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsCustAnnexDtl.GetNoRows > 0 Then
                dblItemQty = rsCustAnnexDtl.GetValue("Quantity")
                strCustCode = rsCustAnnexDtl.GetValue("Customer_code")
                strRef57F4 = rsCustAnnexDtl.GetValue("Ref57F4_no")
                strUpdtCustAnnexHdr = "UPDATE custannex_hdr SET balance_Qty=balance_Qty +" & dblItemQty & " WHERE UNIT_CODE='" & gstrUNITID & "' AND Customer_Code='" & strCustCode & "' AND Item_Code='" & StrItemCode & "' AND ref57f4_no='" & strRef57F4 & "'" & vbCrLf
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
        '--------The Retrieved Document no will further be called for Cancellation by ClsDrCrNote using Function ReverseAPDocument
        Dim rsObjInv As New ADODB.Recordset
        ChkForRejInv = False
        On Error GoTo ErrHandler
        If rsObjInv.State = ADODB.ObjectStateEnum.adStateOpen Then rsObjInv.Close()
        rsObjInv.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rsObjInv.Open("SELECT apdocm_vono FROM ap_docmaster WHERE apdocM_unit='" & gstrUNITID & "' AND apdocm_venprdocno='" & Trim(txtChallanNo.Text) & "' AND apdocm_open=1 AND apdocm_cancel=0 ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not rsObjInv.EOF Then
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
        Dim strSql As String
        Dim dblQty As String
        Dim strYearMonth As String
        On Error GoTo ErrHandler
        UpdateMonthlyMktSchedule = True
        RsObjUpdateSchedules = New ADODB.Recordset
        RsobjSchedules = New ADODB.Recordset
        If RsObjUpdateSchedules.State = ADODB.ObjectStateEnum.adStateOpen Then RsObjUpdateSchedules.Close()
        RsObjUpdateSchedules.Open("SELECT doc_no ,item_code,cust_part_code,dsno,customer_code,quantityknockedoff FROM mkt_invdshistory WHERE UNIT_CODE='" & gstrUNITID & "' AND cancellation_flag=0 AND doc_no='" & txtChallanNo.Text & "' and Item_code = '" & pstrItemCode & "' and cust_Part_code = '" & pstrCustPartCode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        'If Val(Mid(lblDateDes.Text, 4, 2)) < 10 Then
        '    strYearMonth = Mid(lblDateDes.Text, 7, 4) & "0" & Val(Mid(lblDateDes.Text, 4, 2))
        'Else
        '    strYearMonth = Mid(lblDateDes.Text, 7, 4) & Val(Mid(lblDateDes.Text, 4, 2))
        'End If
        strYearMonth = Year(ConvertToDate(lblDateDes.Text)) & VB.Right("0" & Month(ConvertToDate(lblDateDes.Text)), 2)

        dblQty = RsObjUpdateSchedules.Fields(5).Value
        If CDbl(dblQty) > 0 Then
            strSql = "SELECT dsno,isNULL(despatch_Qty,0) as despatch_Qty ,item_code,CUST_DRGNO FROM Monthlymktschedule WHERE UNIT_CODE='" & gstrUNITID & "' AND account_code='" & RsObjUpdateSchedules.Fields(4).Value & "' AND item_code='" & RsObjUpdateSchedules.Fields(1).Value & "' AND status=1 AND despatch_qty > 0 AND dsno ='" & RsObjUpdateSchedules.Fields(3).Value & "' and year_Month <= '" & strYearMonth & "' and Cust_drgNo = '" & RsObjUpdateSchedules.Fields(2).Value & "' ORDER BY Year_Month DESC,dsdatetime DESC"
            mP_Connection.Execute("Set Dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            RsobjSchedules.Open(strSql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
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

        strMakeDate = Year(ConvertToDate(pstrDate)) & VB.Right("0" & Month(ConvertToDate(pstrDate)), 2)
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
        Dim strSql As String
        Dim dblQty As String
        On Error GoTo ErrHandler
        UpdateDailyMktSchedule = True
        RsObjUpdateSchedules = New ADODB.Recordset
        RsobjSchedules = New ADODB.Recordset
        If RsObjUpdateSchedules.State = 1 Then
            RsObjUpdateSchedules.Close()
        End If
        strSql = "SELECT doc_no ,item_code,cust_part_code,dsno,customer_code,quantityknockedoff FROM mkt_invdshistory WHERE UNIT_CODE='" & gstrUNITID & "' AND cancellation_flag=0 AND doc_no='" & txtChallanNo.Text & "' and ITem_code = '" & pstrItemCode & "' and Cust_part_code = '" & pstrCustPartCode & "' and DSNo = '" & pstrDSNo & "'"
        RsObjUpdateSchedules.Open(strSql, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        dblQty = RsObjUpdateSchedules.Fields(5).Value
        If CDbl(dblQty) > 0 Then
            strSql = "SELECT dsno,isNULL(despatch_qty,0) as despatch_qty ,item_code,CUST_DRGNO,trans_date FROM dailymktschedule WHERE UNIT_CODE='" & gstrUNITID & "' AND account_code='" & RsObjUpdateSchedules.Fields(4).Value & "' AND item_code='" & RsObjUpdateSchedules.Fields(1).Value & "' AND status=1 AND despatch_qty > 0 AND dsno ='" & RsObjUpdateSchedules.Fields(3).Value & "' and trans_date <= '" & getDateForDB(lblDateDes.Text) & "' and Cust_drgNo = '" & RsObjUpdateSchedules.Fields(2).Value & "' ORDER BY trans_date DESC,dsdatetime DESC"
            mP_Connection.Execute("Set Dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            If RsobjSchedules.State = 1 Then RsobjSchedules.Close()
            RsobjSchedules.Open(strSql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
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
        Dim strSql As String
        Dim intCtr As Short
        Dim strMSG As String
        Dim strYYYYmm As String
        ReDim mSchTypeArr(0)
        CheckMktSchedules = ""
        strYYYYmm = Year(ConvertToDate(lblDateDes.Text)) & VB.Right("0" & Month(ConvertToDate(lblDateDes.Text)), 2)
        With SpChEntry
            If (UCase(mstrInvType) = "EXP") Then
                For intCtr = 1 To .MaxRows Step 1
                    ReDim Preserve mSchTypeArr(intCtr)
                    Com = New ADODB.Command
                    Com.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Com.CommandText = "MKT_SCHUDULE_CHECK"
                    .Row = intCtr
                    Com.Parameters.Append(Com.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                    Com.Parameters.Append(Com.CreateParameter("@CUSTOMER_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(txtCustCode.Text)))
                    Com.Parameters.Append(Com.CreateParameter("@CONSIGNEE_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(mstrConsigneeCode)))
                    .Col = 1
                    Com.Parameters.Append(Com.CreateParameter("@ITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(.Text)))
                    .Col = 2
                    Com.Parameters.Append(Com.CreateParameter("@CUSTDRG_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(.Text)))
                    Com.Parameters.Append(Com.CreateParameter("@REQ_QTY", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adCurrency, 0))
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
                    mSchTypeArr(intCtr) = Com.Parameters(7).Value
                    Com = Nothing
                Next intCtr
            Else
                For intCtr = 1 To .MaxRows Step 1
                    ReDim Preserve mSchTypeArr(intCtr)
                    Com = New ADODB.Command
                    Com.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Com.CommandText = "MKT_SCHUDULE_CHECK_INV"
                    .Row = intCtr
                    Com.Parameters.Append(Com.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                    Com.Parameters.Append(Com.CreateParameter("@CUSTOMER_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(txtCustCode.Text)))
                    .Col = 1
                    Com.Parameters.Append(Com.CreateParameter("@ITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(.Text)))
                    .Col = 2
                    Com.Parameters.Append(Com.CreateParameter("@CUSTDRG_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(.Text)))
                    Com.Parameters.Append(Com.CreateParameter("@REQ_QTY", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adCurrency, 0))
                    Com.Parameters.Append(Com.CreateParameter("@YYYYMM", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adInteger, strYYYYmm))
                    Com.Parameters.Append(Com.CreateParameter("@SCH_TYPE", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamOutput, 1))
                    Com.Parameters.Append(Com.CreateParameter("@MSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 500))
                    Com.Parameters.Append(Com.CreateParameter("@ERR", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 100))
                    Com.let_ActiveConnection(mP_Connection)
                    Com.Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    If Len(Com.Parameters(8).Value) > 0 Then
                        MsgBox(Com.Parameters(8).Value, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        CheckMktSchedules = "Error"
                        Com = Nothing
                        Exit Function
                    End If
                    mSchTypeArr(intCtr) = Com.Parameters(6).Value
                    Com = Nothing
                Next intCtr
            End If
        End With
        CheckMktSchedules = strMSG
        Exit Function
ErrHandler:
        Com = Nothing
        CheckMktSchedules = "Error"
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
        Dim strSql As String
        Dim intCtr As Short
        Dim strMSG As String
        Dim strYYYYmm As String
        Dim curQty As Decimal
        UpdateMktSchedules = True
        strYYYYmm = Year(ConvertToDate(lblDateDes.Text)) & VB.Right("0" & Month(ConvertToDate(lblDateDes.Text)), 2)
        With SpChEntry
            If (UCase(mstrInvType) = "EXP") Then
                For intCtr = 1 To .MaxRows Step 1
                    Com = New ADODB.Command
                    Com.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Com.CommandText = "MKT_SCHUDULE_KNOCKOFF"
                    .Row = intCtr
                    Com.Parameters.Append(Com.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                    Com.Parameters.Append(Com.CreateParameter("@CUSTOMER_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(txtCustCode.Text)))
                    Com.Parameters.Append(Com.CreateParameter("@CONSIGNEE_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(mstrConsigneeCode)))
                    .Col = 1
                    Com.Parameters.Append(Com.CreateParameter("@ITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(.Text)))
                    .Col = 2
                    Com.Parameters.Append(Com.CreateParameter("@CUSTDRG_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(.Text)))
                    Com.Parameters.Append(Com.CreateParameter("@FLAG", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, pstrUpdType))
                    Com.Parameters.Append(Com.CreateParameter("@SCH_TYPE", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(mSchTypeArr(intCtr))))
                    Com.Parameters.Append(Com.CreateParameter("@YYYYMM", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adInteger, strYYYYmm))
                    .Col = 5
                    Com.Parameters.Append(Com.CreateParameter("@REQ_QTY", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adCurrency, Trim(.Text)))
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
            Else
                For intCtr = 1 To .MaxRows Step 1
                    Com = New ADODB.Command
                    Com.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    Com.CommandText = "MKT_SCHUDULE_KNOCKOFF_INV"
                    .Row = intCtr
                    Com.Parameters.Append(Com.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                    Com.Parameters.Append(Com.CreateParameter("@CUSTOMER_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(txtCustCode.Text)))
                    .Col = 1
                    Com.Parameters.Append(Com.CreateParameter("@ITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(.Text)))
                    .Col = 2
                    Com.Parameters.Append(Com.CreateParameter("@CUSTDRG_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(.Text)))
                    Com.Parameters.Append(Com.CreateParameter("@FLAG", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, pstrUpdType))
                    Com.Parameters.Append(Com.CreateParameter("@SCH_TYPE", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(mSchTypeArr(intCtr))))
                    Com.Parameters.Append(Com.CreateParameter("@YYYYMM", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adInteger, strYYYYmm))
                    .Col = 5
                    Com.Parameters.Append(Com.CreateParameter("@REQ_QTY", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adCurrency, Trim(.Text)))
                    Com.Parameters.Append(Com.CreateParameter("@ERR", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 100))
                    Com.let_ActiveConnection(mP_Connection)
                    Com.Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    If Len(Com.Parameters(8).Value) > 0 Then
                        MsgBox(Com.Parameters(8).Value, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        UpdateMktSchedules = False
                        Com = Nothing
                        Exit Function
                    End If
                    Com = Nothing
                Next intCtr
            End If
        End With
        Exit Function
ErrHandler:
        Com = Nothing
        UpdateMktSchedules = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function UpdateDespAdvise() As Boolean
        '-------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Datetime      : 12-Feb-2007
        ' Function      : To revert the Invoice No from Bar Code related tables
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim Com As ADODB.Command
        UpdateDespAdvise = True
        Com = New ADODB.Command
        Com.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Com.CommandText = "BAR_INV_CANLELLATION"
        Com.Parameters.Append(Com.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
        Com.Parameters.Append(Com.CreateParameter("@INVNO", ADODB.DataTypeEnum.adBigInt, ADODB.ParameterDirectionEnum.adParamInput, , Trim(txtChallanNo.Text)))
        Com.Parameters.Append(Com.CreateParameter("@CUST_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(txtCustCode.Text)))
        Com.Parameters.Append(Com.CreateParameter("@ERR", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 100))
        Com.let_ActiveConnection(mP_Connection)
        Com.Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        If Len(Com.Parameters(3).Value) > 0 Then
            MsgBox(Com.Parameters(3).Value, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            UpdateDespAdvise = False
            Com = Nothing
            Exit Function
        End If
        Com = Nothing
        Exit Function
ErrHandler:
        Com = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function InvAgstBarCode() As Boolean
        '----------------------------------------------------------------------------
        'Author         :   Manoj Kr. Vaish
        'Function       :   Get the BarCodefor Invoice from sales_parameter
        'Comments       :   Date: 04 Feb 2008 ,Issue Id: 22303
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
            strQry = strQry & " from SaleConf a,SalesChallan_Dtl b where a.UNIT_CODE = b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "' AND Doc_No ='" & Trim(txtChallanNo.Text) & "'"
            strQry = strQry & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and"
            strQry = strQry & " a.Location_Code = b.Location_Code And (Fin_Start_Date <= getDate() And Fin_End_Date >= getDate())"
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
        'Comments       :   Date: 04 Feb 2008 ,Issue Id: 22303
        'Revision Date  :   23 Mar 2009 Issue ID : eMpro-20090209-27201
        'History        :   Functionality of Raw Material,Input Invoice through Bar Code
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rsGetQty As ClsResultSetDB
        Dim strsql As String
        Dim CurInvoiceQty As Decimal
        Dim CurBarBondedQty As Decimal
        rsGetQty = New ClsResultSetDB
        mstrupdateBarBondedStockQty = ""
        ' "COMPONENTS" case added in if condition against issue id : 10683802 
        If UCase(Trim(lblInvoiceSubType.Text)) = "RAW MATERIAL" Or UCase(Trim(lblInvoiceSubType.Text)) = "INPUTS" Or UCase(Trim(lblInvoiceSubType.Text)) = "COMPONENTS" Then
            strsql = "select A.CRef_PacketNo,isnull(sum(A.CRef_BalQty),0)as BarQuantity,Isnull(sum(Convert(numeric(16,4),Issue_Qty)),0)as SalesQuantity "
            strsql = strsql & "from Bar_CrossReference A,Bar_Invoice_Issue B where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" & gstrUNITID & "' AND A.CRef_PacketNo=substring(B.Issue_PartbarCode,9,len(CRef_PacketNo)) and "
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
            strsql = strsql & "A.Box_Label=B.Box_label and A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" & gstrUNITID & "' and B.Invoice_No='" & Trim(pstrInvNo) & "' and B.Status_Flag='L' Group By B.Box_label"
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
        'Comments       :   Date: 01 Oct 2008 ,Issue Id: eMpro-20090209-27201
        '----------------------------------------------------------------------------
        Dim rsGetRecord As ClsResultSetDB
        Dim strsql As String
        On Error GoTo ErrHandler
        rsGetRecord = New ClsResultSetDB
        strsql = "select isnull(Batch_Tracking,0) as Batch_Tracking from sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'"
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
        'Comments       :   Date: 23 Mar 2009 ,Issue Id:eMpro-20090209-27201
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
        strStockLocation = StockLocationSalesConf(mstrInvType, mstrInvSubType, "TYPE")
        strsql = "select Item_Code,Batch_no,Batch_Qty from Itembatch_dtl where UNIT_CODE='" & gstrUNITID & "' AND doc_no='" & Trim(txtChallanNo.Text) & "' and Doc_Type='9999'"
        rsgetItemCode = New ClsResultSetDB
        rsgetItemCode.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsgetItemCode.GetNoRows > 0 Then
            rsgetItemCode.MoveFirst()
            Do While Not rsgetItemCode.EOFRecord
                strupdateItemBatchDtl = "Update ItemBatch_Dtl set cancel_flag=1 ,upd_userid='" & mP_User & "',upd_dt=getdate() where UNIT_CODE='" & gstrUNITID & "' and doc_no='" & Trim(txtChallanNo.Text) & "' and doc_type='9999'"
                strupdateItemBatchMst = strupdateItemBatchMst & " Update ItemBatch_Mst set Current_Batch_Qty = Current_Batch_Qty + " & Convert.ToDouble(rsgetItemCode.GetValue("Batch_Qty"))
                strupdateItemBatchMst = strupdateItemBatchMst & " where UNIT_CODE='" & gstrUNITID & "' AND Item_code='" & rsgetItemCode.GetValue("Item_Code") & "' and Batch_no='" & rsgetItemCode.GetValue("Batch_no") & "'"
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
    Public Sub CheckMultipleSOAllowed(ByVal pInvType As String, ByVal pInvSubType As String)
        '-----------------------------------------------------------------------------------
        'Created By      : Manoj Kr.Vaish
        'Issue ID        : eMpro-20090209-27201
        'Creation Date   : 23 Mar 2009
        'Procedure       : To Check Batch Tracking allowed for Any Invoice Type
        '-----------------------------------------------------------------------------------
        Dim rsCheckSo As ClsResultSetDB
        Dim strsql As String
        On Error GoTo ErrHandler
        rsCheckSo = New ClsResultSetDB
        strsql = "select isnull(BatchTrackingAllowed,0) as BatchTrackingAllowed from saleconf where UNIT_CODE='" & gstrUNITID & "' AND description='" & Trim(pInvType) & "' and sub_type_description='" & Trim(pInvSubType) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())"
        rsCheckSo.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsCheckSo.GetNoRows > 0 Then
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
        'Revised On     : 18 May 2009
        'Arguments      : Account Code
        'Return Value   : True/False
        'Issue ID       : eMpro-20090513-31282
        'Reason         : Check ASNTextFileGeneration from Customer Master
        '--------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strQry As String
        Dim Rs As ClsResultSetDB
        AllowASNTextFileGeneration = False
        If (Trim(UCase(lblInvoiceType.Text)) = "NORMAL INVOICE" And UCase(Trim(lblInvoiceSubType.Text)) = "FINISHED GOODS") Then
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
        End If
        Exit Function
ErrHandler:
        Rs = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function FordASNFileGeneration(ByVal pintdocno As Integer, ByVal pstraccountcode As String) As Boolean
        'Revised By     : Manoj Kr. Vaish
        'Revised On     : 18 May 2009
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
        FordASNFileGeneration = True
        strASNdata = ""
        '10259691 
        'strquery = "select * from dbo.FN_GETASNDETAIL(" & pintdocno & "," & gstrUNITID & ")"
        strquery = "select * from dbo.FN_GETASNDETAIL(" & pintdocno & ",'" & gstrUNITID & "')"
        '10259691 
        rsgetData.GetResult(strquery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsgetData.GetNoRows > 0 Then
            If rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Length = 0 Then
                MessageBox.Show("Customer vendor code is not defined for Customer : " & pstraccountcode, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                FordASNFileGeneration = False
                Exit Function
            Else
                strASNdata = "856HD20200000000000000000" & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) + rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & ",     ,     *" & vbCrLf
                strASNdata = strASNdata & "856A M" & txtChallanNo.Text.Trim() & Space(10 - txtChallanNo.Text.Trim().Length) & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & "01" & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & Space(10) & "+00000000100KG+00000000080KG" & "AE0N" & Space(8) & rsgetData.GetValue("TRANSPORT_TYPE").ToString() & Space("12") & "M" & VB.Right(txtChallanNo.Text.Trim(), 5) & Space(4) & Space(35) & "M" & VB.Right(txtChallanNo.Text.Trim(), 5) & Space(5) & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim() & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & Space(6) & rsgetData.GetValue("ARL_CODE").ToString.Trim() & Space(5 - rsgetData.GetValue("ARL_CODE").ToString.Trim().Length) & VB6.Format(GetServerDateTime, "mmddhhmm") & Space(3) & "0000000.00" & vbCrLf
                Dcount = 2
                strcontainerdespQty = Find_Value("select sum(isnull(to_box,0)-isnull(from_box,0)+1) as Desp_Qty from sales_dtl where UNIT_CODE='" & gstrUNITID & "' AND doc_no=" & pintdocno)
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
                    dblcummulativeQty = Find_Value("SELECT CUMMULATIVE_QTY FROM MKT_ASN_CUMFIG WHERE UNIT_CODE='" & gstrUNITID & "' AND CUST_PART_CODE='" & rsgetData.GetValue("CUST_PART_CODE").ToString() & "' AND CUST_PLANTCODE='" & rsgetData.GetValue("CUST_PLANTCODE").ToString() & "'")
                    dblSalesQty = rsgetData.GetValue("SALES_QUANTITY")
                    dblcummulativeQty = dblcummulativeQty - dblSalesQty
                    dblContainerQty = rsgetData.GetValue("CONTAINER_QTY")
                    strASNdata = strASNdata & "856P "
                    strASNdata = strASNdata & rsgetData.GetValue("CUST_PART_CODE").ToString().Trim & Space(30 - rsgetData.GetValue("CUST_PART_CODE").ToString().Length())
                    dblSalesQty = rsgetData.GetValue("Sales_Quantity")
                    strASNdata = strASNdata & "BP+" & Mid("0000000", dblSalesQty.ToString.Length(), 8) & dblSalesQty & "EA+"
                    strTotalQty = Val(strTotalQty) + Val(rsgetData.GetValue("SALES_QUANTITY"))
                    Dcount = Dcount + 1
                    strASNdata = strASNdata & Mid("000000000", dblcummulativeQty.ToString().Length(), 10) & dblcummulativeQty
                    strASNdata = strASNdata & "+0000000000" & Space(10) & txtChallanNo.Text.Trim() & Space(11 - txtChallanNo.Text.Trim().Length()) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString & VB6.Format(rsgetData.GetValue("INVOICE_DATE").ToString(), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE").ToString(), "hhmm") & vbCrLf
                    strASNdata = strASNdata & "856PA" & Space(30) & "+00000000000  +00000000000  " & vbCrLf
                    strASNdata = strASNdata & "856V " & "+000000000000000" & vbCrLf
                    Dcount = Dcount + 2
                    strASNdata = strASNdata & "856C +" & Mid("0000000", dblContainerQty.ToString.Length(), 8) & dblContainerQty & "+" & Mid("0000", rsgetData.GetValue("CONTAINER_DESP_QTY").ToString.Length, 5) & rsgetData.GetValue("CONTAINER_DESP_QTY").ToString & rsgetData.GetValue("CONTAINER").ToString & "90" & vbCrLf
                    Dcount = Dcount + 1
                    mstrupdateASNdtl = Trim(mstrupdateASNdtl) & "UPDATE MKT_ASN_CUMFIG SET CUMMULATIVE_QTY=" & dblcummulativeQty & " WHERE UNIT_CODE='" & gstrUNITID & "' AND CUST_PART_CODE='" & rsgetData.GetValue("CUST_PART_CODE").ToString().Trim() & "' AND CUST_PLANTCODE='" & rsgetData.GetValue("CUST_PLANTCODE").ToString().Trim & "'" & vbCrLf
                    rsgetData.MoveNext()
                Loop
                Dcount = Dcount + 1
                strASNdata = strASNdata & "856T " & Mid("0000", Dcount.ToString.Length, 5) & Dcount & Mid("00000000", strTotalQty.ToString.Length(), 9) & strTotalQty
                gstrASNPath = ReadValueFromINI(Application.StartupPath & "\mind.cfg", "ASNPATH-" & gstrUNITID, "Filepath")
                gstrASNPathForEDI = ReadValueFromINI(Application.StartupPath & "\mind.cfg", "ASNPATH-" & gstrUNITID, "FilepathforEDI")
                If Directory.Exists(gstrASNPath) = False Then
                    Directory.CreateDirectory(gstrASNPath)
                End If
                If Directory.Exists(gstrASNPathForEDI) = False Then
                    Directory.CreateDirectory(gstrASNPathForEDI)
                End If
                strASNFilepath = gstrASNPath & "\C" & txtChallanNo.Text.Trim() & ".dat"
                strASNFilepathforEDI = gstrASNPathForEDI & "\C" & txtChallanNo.Text.Trim() & ".dat"
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
            End If
        End If
        Exit Function
ErrHandler:
        FordASNFileGeneration = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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

        strQry = "select b.Trip_Doc_No from Freight_Gate_Outward_Reg_Hdr a inner join  Freight_Gate_Outward_Trip_Doc_Dtl b on a.Doc_No=b.Doc_No " & _
        "and a.UNIT_CODE=b.UNIT_CODE where b.Trip_Doc_No='" & txtChallanNo.Text & "' and a.UNIT_CODE='" & gstrUNITID & "' and " & _
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

    Public Function CheckIRNCancel() As Boolean
        Try
            Dim blnEway As Boolean = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("Select ISNULL(EWAY_BILL_FUNCTIONALITY,0) EWAY_BILL_FUNCTIONALITY From SaleConf (Nolock) Where Unit_code='" & gstrUnitId & "' and Invoice_Type='" & mstrInvType & "' and Sub_Type='" & mstrInvSubType & "' and datediff(dd,'" & Convert.ToDateTime(lblDateDes.Text).ToString("dd MMM yyyy") & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & Convert.ToDateTime(lblDateDes.Text).ToString("dd MMM yyyy") & "')<=0"))
            If blnEway Then
                Dim strEway As String = Convert.ToString(SqlConnectionclass.ExecuteScalar("Select ISNULL(S.EWAY_IRN_REQUIRED,'') EWAY_IRN_REQUIRED From SalesChallan_Dtl S (Nolock) Where S.Unit_code='" & gstrUnitId & "' and S.Doc_No='" & txtChallanNo.Text & "'"))
                If strEway.ToUpper() = "I" Or strEway.ToUpper() = "B" Then
                    Dim strDeactivateDate As String = Convert.ToString(SqlConnectionclass.ExecuteScalar("Select ISNULL(S.IRN_DEACTIVATE_DATE,'') IRN_DEACTIVATE_DATE From SALESCHALLAN_DTL_IRN S (Nolock) Where S.Unit_code='" & gstrUnitId & "' and S.Doc_No='" & txtChallanNo.Text & "' and ISNULL(S.IRN_Deactivate,0)=1"))
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