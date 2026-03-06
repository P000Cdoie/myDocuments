Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.Data.SqlClient
Friend Class frmMKTMST0003_SOUTH
    Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
	' File Name         :   FRMMKTMST0003.frm
    ' Function          :   Used to add customer-Items
	' Created By        :   Nisha Rai
	' Created On        :   20 April, 2001
	' Revision History  :   Nisha Rai
	'21/09/2001 MARKED CHECKED BY BCs changed on version 4
	'15/02/2002 changed on version 4 on checked out form no 4047
	'===================================================================================
	'Revised By         : Arul Mozhi Varman
	'Revised On         : 05-02-2005
	'Revision History   : To add the Model Name
    '-----------------------------------------------------------------------------------
    'Revised By         : Shubhra Verma
    'Revised On         : 22 Jan 2009
    'Issue ID           : eMpro-20090122-26322
    'REVISED BY         : MANOJ VAISH
    'REVISED ON         : 12 MAY 2009
    'ISSUE ID           : eMpro-20090512-31252
    'DECRIPTION         : ADD Commodity Column in the GRID of Customer Item Master 
    'REVISED BY         : Shubhra Verma
    'REVISED ON         : 21 Dec 2010
    'ISSUE ID           : 1054312
    'DECRIPTION         : while updating any record, it gives primary key violation error.
    'REVISED BY         : SAURAV KUMAR 
    'REVISED ON         : 19 JULY 2011
    'ISSUE ID           : 10117300
    'DESCRIPTION        : Addition of GEM and EAN functionality
    'REVISED BY         : Prashant Rajpal
    'REVISED ON         : 19 Dec 2011
    'ISSUE ID           : 10153377 and 10172991 
    'DESCRIPTION        : Shop code functionality enabled in customer item master 
    'REVISED BY         : Prashant Rajpal
    'REVISED ON         : 30 Dec 2011
    'ISSUE ID           : 10164345 
    'DECRIPTION         : Bin Quantity Field Added 
    '-----------------------------------------------------------------------------------
    ' Revised By     :   Roshan Singh
    ' Revision Date  :   30 Dec 2011
    ' Description    :   Modified for MultiUnit Change Management
    '-----------------------------------------------------------------------------------
    'Revised By      : Neha Ghai
    'Revised On      : 27 April 2012
    'Issue Id        : 10216897 
    'Description     : Primary key violation of CustItem_mst
    '-----------------------------------------------------------------------------------
    'Revised By      : Prashant Rajpal
    'Revised On      : 17 june 2013
    'Issue Id        : 10405554 
    'Description     : Flag thru which  shop code check :is changed now it depends upon allow_shopcode 
    '----------------------------------------------------------------------------------------------------
    'Revised By      : Parveen Kumar
    'Revised On      : 03 Apr 2014
    'Issue Id        : 10571076  
    'Description     : Cust Drawing no added in item help from customer item master
    '----------------------------------------------------------------------------------------------------
    'Created By     : Parveen Kumar
    'Created On     : 16 FEB 2015
    'Description    : eMPro Vehicle BOM
    'Issue ID       : 10737738 
    '-------------------------------------------------------------------------------------------
    'Created By     : Vinod Singh
    'Created On     : 22 June 2015
    'Issue ID       : 10808160 - EOP Changes 
    '-------------------------------------------------------------------------------------------
    'REVISED BY     -  PRASHANT RAJPAL
    'REVISED ON     -  15 JAN 2015
    'PURPOSE        -  10856126 -ASN DOCK CODE FUNCTIONALITY
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'REVISED BY     -  PRASHANT RAJPAL
    'REVISED ON     -  18 SEP 2015
    'PURPOSE        -  10869290 -SERVICE INVOICE 
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'REVISED BY     -  MILIND MISHRA
    'REVISED ON     -  27 OCT 2016
    'PURPOSE        -  101157667 -Auto Invoice part functionality  
    '-------------------------------------------------------------------------------------------------------------------------------------------------------

    Dim clsADOrs As New ClsResultSetDB
    Dim strSQL, mcustItemmst, mStrCustMst, strBUDGETSQL As String
    Dim StrCustItemMstVeiw, mStrItemmst, strPrevAccode, mcustitemmst2 As String
    Dim mintFormIndex As Short
    Dim mactfg As Short
    Dim mactfg1 As Short
    Dim mboolfg As Boolean
    Dim strProdPrice As String
    Dim mvalid As Boolean
    Dim blnAllowBudget As Boolean
    Private Sub chkflag_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles chkflag.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case ctlCustItem.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        txtPrice.Focus()
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub cmdHelp1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdhelp1.Click
        Dim lstrSqL, lstrSQL1 As String
        Dim strhelp1 As String
        Dim strItem() As String
        On Error GoTo Err_Handler
        Dim gobjDB1 As New ClsResultSetDB

        Select Case ctlCustItem.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                With Me.txtItemCode
                    'strhelp1 = ShowList(1, .Maxlength, "", "Item_code", "item_desc", "custItem_mst", " And Item_grp Not In('M','O')")
                    ''--------query changed by Jasmeet on 7/05/2003-----------------------
                    '10869290
                    'strhelp1 = ShowList(1, .MaxLength, "", "Item_code", "description", "Item_mst", "and status='A' AND Hold_flag='0' AND Item_main_grp not in ('M','o')")
                    strhelp1 = ShowList(1, .MaxLength, "", "Item_code", "description", "Item_mst", "and status='A' AND Hold_flag='0' AND Item_main_grp not in ('o')")
                    ''------------------------------------------------------------------
                    'If CheckForItemMainGroup(txtItemCode.Text) = True Then

                    '    If IsRecordExists("select * from budgetitem_mst  where  unit_Code='" & gstrUNITID & "' and  account_code ='" & txtCustCode.Text & "' and item_code='" & txtItemCode.Text & "' and cust_drgno='" & txtdrgno.Text.Trim & "'") Then
                    '        DTPEndDt.Enabled = False
                    '    Else
                    '        DTPEndDt.Enabled = True
                    '        'Dim a As String = "select * from budgetitem_mst  where  unit_Code='" & gstrUNITID & "' and  account_code ='" & txtCustCode.Text & "' and item_code='" & txtItemCode.Text & "' and cust_drgno='" & txtdrgno.Text.Trim & "'"
                    '    End If
                    'End If
                    .Focus()
                End With
                If Val(strhelp1) = -1 Then ' No record
                    Call ConfirmWindow(10430, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                Else
                    Me.txtItemCode.Text = strhelp1
                    gobjDB1.GetResult("SELECT Item_code,Description FROM Item_Mst where unit_Code='" & gstrUNITID & "' and Item_code = '" & strhelp1 & "' ")
                    If gobjDB1.GetNoRows > 0 Then 'RECORD FOUND
                        Me.lbldescription.Text = gobjDB1.GetValue("Description")
                    End If
                    If CheckForItemMainGroup(txtItemCode.Text) = True Then

                        If IsRecordExists("select * from budgetitem_mst  where  unit_Code='" & gstrUNITID & "' and  account_code ='" & txtCustCode.Text & "' and item_code='" & txtItemCode.Text & "' and cust_drgno='" & txtdrgno.Text.Trim & "'") Then
                            DTPEndDt.Enabled = False
                        Else
                            DTPEndDt.Enabled = True
                            'Dim a As String = "select * from budgetitem_mst  where  unit_Code='" & gstrUNITID & "' and  account_code ='" & txtCustCode.Text & "' and item_code='" & txtItemCode.Text & "' and cust_drgno='" & txtdrgno.Text.Trim & "'"
                        End If
                    End If
                End If
               
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                'COMMENTED AND REWRITTEN AGAINST ISSUE ID : 10571076 

                'lstrSqL = "and b.Account_Code ='" & Me.txtCustCode.Text & "'"
                'With Me.txtItemCode
                '    strhelp1 = ShowList(1, .MaxLength, "", "b.Item_code", "a.Description,b.Drg_Desc", "Item_mst a,custItem_mst b", " and a.unit_Code=b.unit_Code and a.unit_Code='" & gstrUNITID & "' and a.item_code = b.item_code " & lstrSqL, , , , , , "a.Unit_Code")
                '    .Focus()
                'End With
                'If Val(strhelp1) = -1 Then ' No record
                '    Call ConfirmWindow(10430, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                'Else
                '    Me.txtItemCode.Text = strhelp1
                '    gobjDB1.GetResult("SELECT Item_code,item_Desc FROM CustItem_Mst where unit_Code='" & gstrUNITID & "' and  Item_code = '" & strhelp1 & "' ")
                '    If gobjDB1.GetNoRows > 0 Then 'RECORD FOUND
                '        Me.lbldescription.Text = gobjDB1.GetValue("item_Desc")
                '    End If
                '    Call txtItemCode_Validating(txtItemCode, New System.ComponentModel.CancelEventArgs(False))
                'End If

                lstrSqL = "select a.Item_code,a.description,b.Cust_Drgno from Item_mst a,custItem_mst b where a.unit_Code=b.unit_Code and a.unit_Code='" & gstrUNITID & "' and a.item_code = b.item_code"
                lstrSqL += " and b.Account_Code ='" & Me.txtCustCode.Text & "'"
                With Me.txtItemCode
                    'strhelp1 = ShowList(1, .Maxlength, "", "Item_code", "item_desc", "custItem_mst", " And Item_grp Not In('M','O')")
                    ''--------query changed by Jasmeet on 7/05/2003-----------------------
                    'strhelp1 = ShowList(1, .MaxLength, "", "Item_code", "description", "Item_mst", "and status='A' AND Hold_flag='0' AND Item_main_grp not in ('M','o')")
                    'strhelp1 = ShowList(1, .MaxLength, "", "a.Item_code", "a.description", "Item_mst a,custItem_mst b", "and a.unit_Code=b.unit_Code and a.unit_Code='" & gstrUNITID & "' and a.item_code = b.item_code and a.status='A' AND a.Hold_flag='0' AND a.Item_main_grp not in ('M','o')", , , , , "b.Drg_Desc", "a.unit_code")
                    strItem = Me.CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, lstrSqL)
                    ''------------------------------------------------------------------
                    .Focus()
                End With
                If IsNothing(strItem) = True Then Exit Sub
                If strItem.GetUpperBound(0) <> -1 Then
                    If (Len(strItem(0)) >= 1) And strItem(0) = "0" Then
                        MsgBox("No Record found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                        Exit Sub
                    Else
                        txtItemCode.Text = strItem(0)
                        gobjDB1.GetResult("SELECT Item_code,item_Desc FROM CustItem_Mst where unit_Code='" & gstrUNITID & "' and  Item_code = '" & strhelp1 & "' ")
                        If gobjDB1.GetNoRows > 0 Then 'RECORD FOUND
                            Me.lbldescription.Text = gobjDB1.GetValue("item_Desc")
                        End If
                        Call txtItemCode_Validating(txtItemCode, New System.ComponentModel.CancelEventArgs(False))
                    End If
                End If
        End Select
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdhelp.Click
        '----------------------------------------------------------------------------
        'Argument       :   NIL
        'Return Value   :   NIL
        'Function       :   To show help on Account Master
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        Dim varRetVal As Object
        Dim strHelp As String
        On Error GoTo Err_Handler
        Dim gobjDB1 As New ClsResultSetDB
        Select Case ctlCustItem.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                With Me.txtCustCode
                    strHelp = ShowList(1, .MaxLength, "", "customer_code", "cust_name", "customer_mst")
                    .Focus()
                End With
                If Val(strHelp) = -1 Then ' No record
                    Call ConfirmWindow(10170, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                Else
                    Me.txtCustCode.Text = strHelp
                    If AllowShopCodeflag(Me.txtCustCode.Text) = True Then
                        lblShopcode.Visible = True
                        txtShopCode.Visible = True
                        txtShopCode.Enabled = True : txtShopCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    Else
                        lblShopcode.Visible = False
                        txtShopCode.Visible = False
                        txtShopCode.Enabled = True : txtShopCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    End If
                    gobjDB1.GetResult("SELECT customer_Code,cust_name FROM Customer_mst Where  unit_Code='" & gstrUNITID & "' and  customer_code = '" & strHelp & "' ")
                    If gobjDB1.GetNoRows > 0 Then 'RECORD FOUND
                        Me.lblcustcode.Text = gobjDB1.GetValue("cust_name")
                    End If
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                With Me.txtCustCode
                    strHelp = ShowList(1, .MaxLength, "", "customer_code", "cust_name", "Customer_Mst")
                    .Focus()
                End With
                If Val(strHelp) = -1 Then ' No record
                    Call ConfirmWindow(10170, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                    txtCustCode.Focus()
                Else
                    Me.txtCustCode.Text = strHelp
                    If AllowShopCodeflag(Me.txtCustCode.Text) = True Then
                        lblShopcode.Visible = True
                        txtShopCode.Visible = True
                    Else
                        lblShopcode.Visible = False
                        txtShopCode.Visible = False
                    End If
                    gobjDB1.GetResult("SELECT customer_code,cust_name FROM Customer_mst where  unit_Code='" & gstrUNITID & "' and  customer_code='" & Trim(strHelp) & "' ")
                    If gobjDB1.GetNoRows > 0 Then 'RECORD FOUND
                        Me.lblcustcode.Text = gobjDB1.GetValue("cust_name")
                    End If
                    If Len(Trim(txtCustCode.Text)) > 0 Then
                        txtItemCode.Focus()
                    Else
                        txtCustCode.Focus()
                    End If
                End If
        End Select
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtcontainerqty_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtcontainerqty.TextChanged
        txtcontainerqty.Text = Number_Chk(txtcontainerqty.Text)
    End Sub
    Private Sub txtcontainerqty_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtcontainerqty.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                'Me.ctlCustItem.SetFocus
                'Code changed by Arul on 05-02-2005
                Me.TxtModel.Focus()
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub ctlCustItem_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles ctlCustItem.ButtonClick
        '--------------------------------------------------------------
        'PURPOSE-FOR ADD/EDIT/DELETE/SAVE RECORD TO bus_sup_mst
        'PARAMETER-GIVES YOU VALUE FOR EACH BUTTON CLICKED ON CONTROL
        '--------------------------------------------------------------
        Dim strDelProd As String
        
On Error GoTo Err_Handler
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD ' for add
                Call EnableControls(True, Me, True)
                LineItems.Enabled = True
                lblcustcode.Text = "" : lbldescription.Text = ""
                lblDate.Text = "" : lblEff_dt1.Text = ""
                lblEff_dt2.Text = "" : lblPrice1.Text = "" : lblPrice2.Text = ""
                lblDate.Text = VB6.Format(GetServerDate(), gstrDateFormat)
                chkflag.CheckState = System.Windows.Forms.CheckState.Checked
                txtPrice.Text = "0.00"
                txtCustSupp.Text = "0.00"
                txtTool.Text = "0.00"
                cmbDeliveryPattrn.SelectedIndex = 0
                DTPPOFTime.Value = GetServerDateTime()

                DTPEndDt.Enabled = False
                For i As Integer = 0 To LineItems.Items.Count - 1
                    LineItems.SetItemChecked(i, False)
                Next
                'If CheckForItemMainGroup(txtItemCode.Text) = True Then

                '    If IsRecordExists("select * from budgetitem_mst  where  unit_Code='" & gstrUNITID & "' and  account_code ='" & txtCustCode.Text & "' and item_code='" & txtItemCode.Text & "' and cust_drgno='" & txtdrgno.Text.Trim & "'") Then
                '        DTPEndDt.Enabled = False
                '    Else
                '        DTPEndDt.Enabled = True
                '        'Dim a As String = "select * from budgetitem_mst  where  unit_Code='" & gstrUNITID & "' and  account_code ='" & txtCustCode.Text & "' and item_code='" & txtItemCode.Text & "' and cust_drgno='" & txtdrgno.Text.Trim & "'"
                '    End If
                'End If

                '-----------------------------------------------------------------------------------
                'Modified BY         : SAURAV KUMAR
                'Modified ON         : 19 JULY 2011
                'ISSUE ID            : 10117300
                'DESCRIPTION         : Addition of GEM and EAN functionality
                '-----------------------------------------------------------------------------------
                txtEanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtEanNo.Enabled = False
                Me.txtCustCode.Focus()
                If blnAllowBudget = True Then
                    mP_Connection.Execute("delete from tmp_budgetitem_mst where  unit_Code='" & gstrUNITID & "' and  ip_address = '" & gstrIpaddressWinSck & "' And Account_Code='" & txtCustCode.Text.Trim & "' And Item_Code='" & txtItemCode.Text.Trim & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
                Me.lblShopcode.Visible = False
                Me.txtShopCode.Visible = False
                cboItmtype.SelectedIndex = -1
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT 'for editf
                If Len(txtCustCode.Text) <= 0 Then
                    Call ConfirmWindow(10058)
                    ctlCustItem.Revert()
                    Call EnableControls(False, Me)
                    LineItems.Enabled = False
                    txtCustCode.Enabled = True
                    cmdhelp.Enabled = True
                    txtItemCode.Enabled = True
                    txtItemCode.BackColor = System.Drawing.Color.White
                    cmdhelp1.Enabled = True


                    txtCustCode.Focus()
                    '10856126
                    If txtdockcode.Text.Trim.ToString.Trim.Length > 0 Then
                        txtdockcode.Enabled = False
                        txtdockcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    End If
                    '10856126

                Else
                    Call EnableControls(True, Me)
                    LineItems.Enabled = True
                    Me.txtCustCode.Enabled = False
                    Me.txtItemCode.Enabled = False
                    Me.txtdrgno.Enabled = False
                    Me.txtdrgdes.Enabled = False
                    txtdrgno.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    cmdhelp.Enabled = False
                    cmdhelp1.Enabled = False
                    Me.chkflag.Focus()
                    '10856126
                    If txtdockcode.Text.Trim.ToString.Trim.Length > 0 Then
                        txtdockcode.Enabled = False
                        txtdockcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    End If
                    '10856126
                    If IsRecordExists("select * from budgetitem_mst  where  unit_Code='" & gstrUNITID & "' and  account_code ='" & txtCustCode.Text & "' and item_code='" & txtItemCode.Text & "' and cust_drgno='" & txtdrgno.Text.Trim & "'") Then
                        DTPEndDt.Enabled = False
                    Else
                        Dim a As String = "select * from budgetitem_mst  where  unit_Code='" & gstrUNITID & "' and  account_code ='" & txtCustCode.Text & "' and item_code='" & txtItemCode.Text & "' and cust_drgno='" & txtdrgno.Text.Trim & "'"
                        DTPEndDt.Enabled = True
                        'Dim a As String = "select * from budgetitem_mst  where  unit_Code='" & gstrUNITID & "' and  account_code ='" & txtCustCode.Text & "' and item_code='" & txtItemCode.Text & "' and cust_drgno='" & txtdrgno.Text.Trim & "'"
                    End If
                    'checkstate will me enable/disable at Auto Invoice Part
                    If lblAutoInvPart.Checked = True Then
                        lblAutoInvPart.Enabled = False
                    Else
                        lblAutoInvPart.Enabled = True

                    End If
                End If

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE 'for delete
                If Len(txtCustCode.Text) <= 0 Then
                    Call ConfirmWindow(10058)
                    txtCustCode.Focus()
                    Exit Sub
                Else
                    strSQL = "select a.*,b.* from Cust_ord_dtl a,Cust_ord_hdr b where"
                    strSQL = strSQL & " a.unit_Code=b.unit_code and a.unit_Code='" & gstrUNITID & "' and b.Account_Code=A.Account_Code and a.Active_Flag in('A','L') and "
                    strSQL = strSQL & "a.Account_Code= '" & Trim(txtCustCode.Text) & "' and a.Item_Code = '" & Trim(txtItemCode.Text) & "' and a.cust_drgno = '" & Me.txtdrgno.Text & "' "
                    clsADOrs.GetResult(strSQL)
                    If clsADOrs.GetNoRows <= 0 Then
                        strSQL = "DELETE FROM CustItem_mst WHERE  unit_Code='" & gstrUNITID & "' and  Account_Code = '" & Me.txtCustCode.Text & "' and Item_code ='" & Me.txtItemCode.Text & "' "
                        strDelProd = "DELETE FROM prod_price_mst WHERE  unit_Code='" & gstrUNITID & "' and  Cust_C = '" & Me.txtCustCode.Text & "' and Product_no ='" & Me.txtItemCode.Text & "'"
                        If ConfirmWindow(10054, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_CRITICAL, 60096) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                            With mP_Connection
                                .Close()
                                .Open()
                                .BeginTrans()
                                .Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                .Execute(strDelProd, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                .CommitTrans()
                            End With
                            Call EnableControls(False, Me, True)

                            For i As Integer = 0 To LineItems.Items.Count - 1
                                LineItems.SetItemChecked(i, False)
                            Next
                            LineItems.Enabled = False
                            lblcustcode.Text = "" : lbldescription.Text = ""
                            lblDate.Text = "" : lblEff_dt1.Text = ""
                            lblEff_dt2.Text = "" : lblPrice1.Text = "" : lblPrice2.Text = ""
                            txtCustCode.Enabled = True : txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdhelp.Enabled = True
                            txtCustCode.Focus()
                            Exit Sub
                        End If
                    Else
                        Call ConfirmWindow(10426, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO) 'Message Cust Order is Active can not delete
                        txtCustCode.Focus()
                        txtItemCode.Enabled = False
                        cmdhelp1.Enabled = False
                        Exit Sub
                    End If
                    Exit Sub
                End If
                Call Me.ctlCustItem.Revert()
                Me.ctlCustItem.Enabled(5) = False
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE 'Save button clicked
                mactfg = IIf(Me.chkflag.CheckState = System.Windows.Forms.CheckState.Checked, 1, 0)
                mactfg1 = IIf(Me.lblAutoInvPart.CheckState = System.Windows.Forms.CheckState.Checked, 1, 0)
                'Checking for all required validations

                If ValidRecord() Then
                    '10737738
                    If ValidateVEHBOM() = False Then
                        Exit Sub
                    End If

                    Select Case ctlCustItem.Mode
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                            If PrimaryKeycheck() Then
                                If SaveData(Me.txtCustCode.Text, Me.txtItemCode.Text) Then 'string for insert/update
                                    '-----------------------------------------------------------------------------------
                                    'Modified BY         : SAURAV KUMAR
                                    'Modified ON         : 19 JULY 2011
                                    'ISSUE ID            : 10117300
                                    'DESCRIPTION         : Addition of GEM and EAN functionality
                                    '-----------------------------------------------------------------------------------
                                    Call aerobinvaluechanged()
                                    Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    ctlCustItem.Revert() : gblnCancelUnload = False : gblnFormAddEdit = False
                                    Me.ctlCustItem.Enabled(5) = False
                                    Call EnableControls(False, Me, True)
                                    For i As Integer = 0 To LineItems.Items.Count - 1
                                        LineItems.SetItemChecked(i, False)
                                    Next
                                    LineItems.Enabled = False
                                    txtCustCode.Enabled = True : txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdhelp.Enabled = True
                                    txtItemCode.Enabled = True : txtItemCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdhelp1.Enabled = True
                                    txtCustCode.Focus()
                                    cboItmtype.SelectedIndex = -1
                                End If
                            End If

                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                            If SaveData(Me.txtCustCode.Text, Me.txtItemCode.Text) Then 'string for insert/update
                                gblnCancelUnload = False : gblnFormAddEdit = False
                                Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                ctlCustItem.Revert()
                                Me.ctlCustItem.Enabled(5) = False
                                Call EnableControls(False, Me, True)
                                LineItems.Enabled = False
                                For i As Integer = 0 To LineItems.Items.Count - 1
                                    LineItems.SetItemChecked(i, False)
                                Next
                                DTPEndDt.Value = New Date(1990, 1, 1)
                                txtCustCode.Enabled = True
                                Me.txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                                txtItemCode.Enabled = True : txtItemCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                                cmdhelp.Enabled = True : cmdhelp1.Enabled = True
                                cboItmtype.SelectedIndex = -1

                                txtCustCode.Focus()
                            End If
                    End Select
                Else
                    gblnCancelUnload = True : gblnFormAddEdit = True
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL 'for cancel
                strPrevAccode = txtCustCode.Text
                If (Me.ctlCustItem.Mode > 0) Or (Me.ctlCustItem.Button > 0) Then
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION, 60095) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        gblnCancelUnload = False
                        gblnFormAddEdit = False
                        ctlCustItem.Focus()
                        If Len(txtCustCode.Text) = 0 Then
                            Call RefreshForm(False)
                            txtCustCode.Enabled = True
                            cmdhelp.Enabled = True
                            txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            txtCustCode.Focus()
                            DTPEndDt.Value = New Date(1990, 1, 1)
                       
                        Else
                            'Code Changed By Arul on 05-02-2005 to Adding OR part for MODE_EDIT
                            If Me.ctlCustItem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or Me.ctlCustItem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                                Call RefreshForm(False)
                                txtCustCode.Enabled = True
                                cmdhelp.Enabled = True
                                txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                                txtCustCode.Focus()
                                ctlCustItem.Revert()
                                DTPEndDt.Value = New Date(1990, 1, 1)
                                For i As Integer = 0 To LineItems.Items.Count - 1
                                    LineItems.SetItemChecked(i, False)
                                Next
                                Exit Sub
                            Else
                                Call RefreshForm(True)
                            End If
                            txtCustCode.Text = strPrevAccode
                            Call DispRecordfromProdPrice()
                        End If
                    Else
                        ctlCustItem.Focus()
                        Exit Sub
                    End If
                End If
                ctlCustItem.Revert()
                txtCustCode.Text = strPrevAccode
                If Len(strPrevAccode) <= 0 Then
                    ctlCustItem.Enabled(2) = False
                    ctlCustItem.Enabled(1) = False
                End If
                Call EnableControls(False, Me)
                Me.txtCustCode.Enabled = True : Me.txtCustCode.BackColor = System.Drawing.Color.White
                cmdhelp.Enabled = True
                'Call RefreshForm(False)
                txtCustCode.Focus()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
        End Select
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub aerobinvaluechanged()
        Dim strsql As String
        '-----------------------------------------------------------------------------------
        'Added BY         : SAURAV KUMAR
        'Added ON         : 19 JULY 2011
        'ISSUE ID         : 10117300
        'DESCRIPTION      : Addition of GEM and EAN functionality
        '-----------------------------------------------------------------------------------
        If chkGEM.Checked = True Then
            mP_Connection.Close()
            mP_Connection.Open()
            mP_Connection.BeginTrans()
            strsql = "UPDATE CUSTOMER_MST SET IS_AEROBIN_CUST='1' WHERE  unit_Code='" & gstrUNITID & "' and CUSTOMER_CODE='" & Trim(txtCustCode.Text) & "'"
            mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            mP_Connection.CommitTrans()
        End If
    End Sub
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        On Error GoTo errHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm") '("pddrep_listofitems.htm")
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0003_SOUTH_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo Err_Handler
        'Checking the form name in the Windows list
        mdifrmMain.CheckFormName = mintFormIndex
        txtCustCode.Focus()
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0003_SOUTH_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo Err_Handler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0003_SOUTH_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Err_Handler
        If ctlCustItem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or ctlCustItem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Or ctlCustItem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If KeyCode = System.Windows.Forms.Keys.Escape Then
                If ctlCustItem.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    Call ctlCustItem_ButtonClick(ctlCustItem, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL))
                End If
                '        ElseIf KeyCode = vbKeyReturn Then   'Enter key pressed
                '            SendKeys "{TAB}", True
            End If
        End If
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlFormHeader1_Click(ctlFormHeader1, New System.EventArgs())
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0003_SOUTH_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo Err_Handler
        'resource
        'Get the index of form in the Windows list
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FillLabelFromResFile(Me) 'To Fill label description from Resource file
        Call FitToClient(Me, fraCustItem, ctlFormHeader1, ctlCustItem) 'To fit the form in the MDI
        Call EnableControls(False, Me, True) 'To Disable controls
        txtCustCode.Enabled = True : txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdhelp.Enabled = True
        LineItems.Enabled = False

        'Disabling Edit, Delete and Print buttons
        ctlCustItem.Enabled(1) = False
        ctlCustItem.Enabled(2) = False
        ctlCustItem.Enabled(5) = False
        ctlCustItem.Enabled(4) = False
        lblDate.Text = VB6.Format(GetServerDate(), gstrDateFormat)
        txtPrice.Text = "0.00"
        txtCustSupp.Text = "0.00"
        txtTool.Text = "0.00"
        DTPEndDt.Value = New Date(1990, 1, 1)
        For i As Integer = 0 To LineItems.Items.Count - 1
            LineItems.SetItemChecked(i, False)
        Next
        Dim rsAllowBudget_flag As ClsResultSetDB
        rsAllowBudget_flag = New ClsResultSetDB
        rsAllowBudget_flag.GetResult("SELECT AllowBudget_flag from Sales_parameter where  unit_Code='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsAllowBudget_flag.RowCount > 0 Then
            blnAllowBudget = IIf(rsAllowBudget_flag.GetValue("AllowBudget_flag") = True, True, False)
        End If
        If blnAllowBudget = True Then
            rsAllowBudget_flag = Nothing
            frabudget.Visible = True
        Else
            frabudget.Visible = False
        End If
        If blnAllowBudget = True Then
            mP_Connection.Execute("delete from tmp_budgetitem_mst where  unit_Code='" & gstrUNITID & "' and  ip_address = '" & gstrIpaddressWinSck & "' And Account_Code='" & txtCustCode.Text.Trim & "' And Item_Code='" & txtItemCode.Text.Trim & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        End If
        cmbDeliveryPattrn.SelectedIndex = 0
        DTPPOFTime.Value = GetServerDateTime()
        txtShopCode.Visible = False
        lblShopcode.Visible = False
        FillItemType()
        FillLineList()
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub RefreshForm(ByRef prmRefreshFg As Boolean)
        '----------------------------------------------------------------------------
        'Argument       :
        'Return Value   :   Boolean
        'Function       :   RefreshForm
        'Comments       :   Used to display or clear Details
        '----------------------------------------------------------------------------
        On Error GoTo Err_Handler
        If prmRefreshFg = False Then
            Call EnableControls(False, Me, True)
            LineItems.Enabled = False

            Call RefreshLablelofPordPrice()
            gblnCancelUnload = False
            gblnFormAddEdit = False
        Else
            'Changed for Issue ID eMpro-20090512-31252 Starts--Show Commodity details
            mcustitemmst2 = "SELECT C.Account_Code,C.Cust_Drgno,C.Drg_Desc,I.Item_code,I.AUTO_INVOICE_PART,C.Item_Desc,C.shop_name,C.gate_no,C.container,C.container_qty,C.Active,isnull(C.VARMODEL,'') VARMODEL, isnull(C.CUST_MTRL,0.00) as CUST_MTRL, isnull(C.TOOL_COST,0.00) as TOOL_COST "
            'mcustitemmst2 = mcustitemmst2 & "FROM CustItem_mst Where Account_Code = '"
            mcustitemmst2 = mcustitemmst2 & ",C.Schupldreqd,C.commodity,C.Delivery_Pattern ,isnull(C.POF,'12:00') POF, C.GEM_CODE,ISNULL(C.shop_code,'') shop_code,C.BINQUANTITY,C.PARTTYPE , C.dockcode ,C.Fixed_Code,C.Length,C.Line,C.IsFullScanning,C.IsNissanLabel, C.TORISHI_CODE  FROM CustItem_mst C,ITEM_MST I Where  C.unit_Code='" & gstrUNITID & "' and C.ITEM_CODE=I.ITEM_CODE AND C.ITEM_CODE=I.ITEM_CODE AND C.Account_Code = '"
            mcustitemmst2 = mcustitemmst2 & txtCustCode.Text & "' and C.Item_Code = '" & Me.txtItemCode.Text & "'" 'primarykeychanged and cust_drgno =   "
            strSQL = mcustitemmst2 '& "'" & Me.txtdrgno & "'"
            'Me.txtItemCode.ExistRecQry = mcustitemmst2
            clsADOrs.GetResult(strSQL)
            Me.txtcommodity.Text = clsADOrs.GetValue("commodity")
            'Changed for Issue ID eMpro-20090512-31252 Ends
            Me.txtItemCode.Text = clsADOrs.GetValue("Item_Code")
            Me.lbldescription.Text = clsADOrs.GetValue("Item_Desc")
            Me.txtdockcode.Text = clsADOrs.GetValue("dockcode")

            '-----------------------------------------------------------------------------------
            'Added BY         : SAzURAV KUMAR
            'Added ON         : 19 JULY 2011
            'ISSUE ID         : 10117300
            'DESCRIPTION      : Addition of GEM and EAN functionality
            '-----------------------------------------------------------------------------------
            Me.txtEanNo.Text = clsADOrs.GetValue("GEM_CODE")
            chkGEM.Checked = True
            txtEanNo.Enabled = False
            txtEanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Me.txtTorishiCode.Text = clsADOrs.GetValue("TORISHI_CODE")
            Me.txtdrgno.Text = clsADOrs.GetValue("Cust_Drgno")
            Me.txtdrgdes.Text = clsADOrs.GetValue("Drg_Desc")
            Me.txtShop.Text = IIf(clsADOrs.GetValue("shop_name") = "", "", clsADOrs.GetValue("shop_name"))
            Me.txtGate.Text = IIf(clsADOrs.GetValue("gate_no") = "", "", clsADOrs.GetValue("gate_no"))
            Me.txtcontainer.Text = IIf(clsADOrs.GetValue("container") = "", "", clsADOrs.GetValue("container"))
            Me.txtcontainerqty.Text = IIf(clsADOrs.GetValue("container_qty") = "", 0, clsADOrs.GetValue("container_qty"))
            Me.cmbDeliveryPattrn.SelectedIndex = 0
            Me.cmbDeliveryPattrn.Text = clsADOrs.GetValue("Delivery_Pattern") + ""
            Me.DTPPOFTime.Text = clsADOrs.GetValue("POF")
               mboolfg = clsADOrs.GetValue("AUTO_INVOICE_PART")
            If mboolfg = True Then
                Me.lblAutoInvPart.CheckState = System.Windows.Forms.CheckState.Checked
            Else
                Me.lblAutoInvPart.CheckState = System.Windows.Forms.CheckState.Unchecked
            End If
            'issue id 10164345 
            Me.txtBinqty.Text = IIf(clsADOrs.GetValue("BinQuantity") = "", 0, clsADOrs.GetValue("BinQuantity"))
            'issue id 10164345  end
            If clsADOrs.GetValue("schupldreqd") = True Then
                chkSchUpldReqd.Checked = 1
            Else
                chkSchUpldReqd.Checked = 0
             End If
            'Code Added By Arul on 05-02-2005
            Me.TxtModel.Text = clsADOrs.GetValue("VARMODEL")
            'Addition ends here
            mboolfg = clsADOrs.GetValue("Active")
            If mboolfg = True Then
                Me.chkflag.CheckState = System.Windows.Forms.CheckState.Checked
            Else
                Me.chkflag.CheckState = System.Windows.Forms.CheckState.Unchecked
            End If
            Me.txtCustSupp.Text = clsADOrs.GetValue("CUST_MTRL")
            Me.txtTool.Text = clsADOrs.GetValue("TOOL_COST")
            If AllowShopCodeflag(Me.txtCustCode.Text) = True Then
                txtShopCode.Text = clsADOrs.GetValue("SHOP_CODE")
            End If
            If Trim(clsADOrs.GetValue("PARTTYPE")) = "" Then
                cboItmtype.SelectedIndex = -1
            Else
                cboItmtype.Text = Trim(clsADOrs.GetValue("PARTTYPE"))
            End If
            Me.txtBxFixedCd.Text = Trim(clsADOrs.GetValue("Fixed_Code"))
            Me.txtBxLength.Text = clsADOrs.GetValue("Length")
            Me.chkBxNissanLbl.Checked = clsADOrs.GetValue("IsNissanLabel")
            Me.chkBxFullScan.Checked = clsADOrs.GetValue("IsFullScanning")
            Dim chkListVal As String = clsADOrs.GetValue("Line")
            Dim lineitms() As String = chkListVal.Split(",")
            Dim strLineItems As String
            Dim iCnt As Integer
            Dim iCnt1 As Integer
            'Dim ChlLst As System.Windows.Forms.CheckedListBox.CheckedItemCollection
            'ChlLst = Me.LineItems.CheckedItems
            Dim ChlLst As System.Windows.Forms.CheckedListBox.ObjectCollection
            ChlLst = Me.LineItems.Items
            'Dim arr As String
            'Dim itm As 
            'For Each arr In lineitms
            '    For Each itm In LineItems.Items
            '        If String.Compare(arr, itm, True = 0) Then
            '            itm()
            '        End If
            '    Next
            'Next
            For iCnt1 = 0 To lineitms.Count - 1
                For iCnt = 0 To ChlLst.Count - 1
                    If String.Compare(Trim(ChlLst(iCnt).ToString).ToUpper, Trim(lineitms(iCnt1).ToString.ToUpper)) = 0 Then
                        LineItems.SetItemChecked(iCnt, True)
                        Exit For
                    End If


                Next
            Next

            'ChlLst = Me.LineItems.CheckedItems
            'For iCnt = 0 To ChlLst.Count - 1
            '    strLineItems = strLineItems + "" & Trim(ChlLst(iCnt).ToString) & ","
            'Next
            If CheckForItemMainGroup(txtItemCode.Text) = True Then
                If blnAllowBudget = True Then
                    cmdbutton4.Enabled = True
                    clsADOrs.ResultSetClose()
                    clsADOrs = New ClsResultSetDB
                    mcustitemmst2 = "select isnull(colour_code,'')as colour_code ,isnull(category_code,'')as category_code,isnull(commodity_code,'')as commodity_code,EndDate from budgetitem_mst where  unit_Code='" & gstrUNITID & "' and  account_code='" & Me.txtCustCode.Text.Trim & "'"
                    mcustitemmst2 = mcustitemmst2 & " and item_code ='" & Me.txtItemCode.Text.Trim & "' and cust_drgno='" & Me.txtdrgno.Text.Trim & "'"
                    clsADOrs.GetResult(mcustitemmst2)
                    txtcolour.Text = clsADOrs.GetValue("Colour_code")
                    txtcategory.Text = clsADOrs.GetValue("category_code")
                    txtcommod.Text = clsADOrs.GetValue("commodity_code")
                    If Not String.IsNullOrEmpty(clsADOrs.GetValue("EndDate")) Then
                        If (clsADOrs.GetValue("EndDate").ToString().ToLower().Equals("unknown")) Then
                            DTPEndDt.Value = New Date(1990, 1, 1)
                        Else
                            DTPEndDt.Value = clsADOrs.GetValue("EndDate")
                        End If
                    Else
                        DTPEndDt.Value = New Date(1990, 1, 1)
                    End If

                    DTPEndDt.Enabled = False
                    'If IsRecordExists("select * from budgetitem_mst  where  unit_Code='" & gstrUNITID & "' and  account_code ='" & txtCustCode.Text & "' and item_code='" & txtItemCode.Text & "' and cust_drgno='" & txtdrgno.Text.Trim & "'") Then
                    '    DTPEndDt.Enabled = False
                    'Else
                    '    DTPEndDt.Enabled = True
                    '    'Dim a As String = "select * from budgetitem_mst  where  unit_Code='" & gstrUNITID & "' and  account_code ='" & txtCustCode.Text & "' and item_code='" & txtItemCode.Text & "' and cust_drgno='" & txtdrgno.Text.Trim & "'"
                    'End If

                    mP_Connection.Execute("delete from tmp_budgetitem_mst where  unit_Code='" & gstrUNITID & "' and  ip_address='" & gstrIpaddressWinSck & "' and account_code='" & txtCustCode.Text.Trim & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    'query comented and added by shubhra on 21 dec 2010
                    'Issue ID   :   1054312
                    'mP_Connection.Execute("insert into tmp_budgetitem_mst(account_code,cust_drgno,item_code,colour_code,category_code,commodity_code,MODEL_CODE,USAGE_QTY,ENT_DT,ENT_USERID,UPD_DT,UPD_USERID,ip_address,VARIANT_CODE) select account_code,cust_drgno,item_code,colour_code,category_code,commodity_code,MODEL_CODE,USAGE_QTY,ENT_DT,ENT_USERID,UPD_DT,UPD_USERID,'" & gstrIpaddressWinSck & "',VARIANT_CODE from budgetitem_mst Where account_code ='" & txtCustCode.Text.Trim & "' And Item_Code='" & txtItemCode.Text.Trim & "' ")
                    mP_Connection.Execute("insert into tmp_budgetitem_mst(account_code,cust_drgno,item_code,colour_code,category_code,commodity_code,MODEL_CODE,USAGE_QTY,ENT_DT,ENT_USERID,UPD_DT,UPD_USERID,ip_address,VARIANT_CODE,Unit_Code,DefaultModel,EndDate) select account_code,cust_drgno,item_code,colour_code,category_code,commodity_code,MODEL_CODE,USAGE_QTY,ENT_DT,ENT_USERID,UPD_DT,UPD_USERID,'" & gstrIpaddressWinSck & "',VARIANT_CODE,'" & gstrUNITID & "',DefaultModel,EndDate from budgetitem_mst Where  unit_Code='" & gstrUNITID & "' and  account_code ='" & txtCustCode.Text.Trim & "' And Item_Code='" & txtItemCode.Text.Trim & "' and cust_drgno = '" & txtdrgno.Text & "' ")

                End If
            End If
            Call DispRecordfromProdPrice()

        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0003_SOUTH_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        On Error GoTo Err_Handler
        'Declarations
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        gblnCancelUnload = False
        If UnloadMode >= 0 And UnloadMode <= 5 Then
            If ctlCustItem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or ctlCustItem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                enmValue = ConfirmWindow(10055, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Or enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        'Save the data before closing
                        Call ctlCustItem_ButtonClick(ctlCustItem, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE))
                    End If
                Else
                    'Set the global variable
                    gblnCancelUnload = True
                    gblnFormAddEdit = True
                End If
            End If
        End If
        'Checking the status
        If gblnCancelUnload = True Then
            eventArgs.Cancel = 1
            If ctlCustItem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then txtdrgdes.Focus() Else txtCustCode.Focus()
        End If
        Exit Sub
Err_Handler:
        gblnCancelUnload = True
        gblnFormAddEdit = True
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTMST0003_SOUTH_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo Err_Handler
        'REFRESH
        'Removing the form name from list
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        'Setting the corresponding node's tag
        'Closing the recordset
        'Releasing the form reference
        Me.Dispose()
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function checkgem_code() As Integer
        '-----------------------------------------------------------------------------------
        'Added BY         : SAURAV KUMAR
        'Added ON         : 19 JULY 2011
        'ISSUE ID         : 10117300
        'DESCRIPTION      : Addition of GEM and EAN functionality
        '-----------------------------------------------------------------------------------
        Dim clsADOrs As ClsResultSetDB = New ClsResultSetDB
        Dim rows As Integer
        clsADOrs.GetResult("SELECT GEM_CODE FROM CUSTITEM_MST WHERE  unit_Code='" & gstrUNITID & "' and  GEM_CODE='" & Trim(txtEanNo.Text) & "'")
        rows = clsADOrs.RowCount
        If rows <> 0 Then
            MsgBox("EAN No already exists for part code " & txtItemCode.Text)
            clsADOrs.ResultSetClose()
            clsADOrs = Nothing
            Return rows
        End If
    End Function
    Private Function SaveData(ByVal prmAccode As String, ByVal prmitcode As String) As Boolean
        '----------------------------------------------------------------------------
        'Argument       :   Buyer Code
        'Return Value   :   Boolean
        'Function       :   Save data to table
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo Err_Handler
        SaveData = False
        strSQL = ""
        strBUDGETSQL = ""
        If txtcontainerqty.Text = "" Then
            txtcontainerqty.Text = CStr(0)
        End If
        If Len(Trim(txtCustSupp.Text)) = 0 Then txtCustSupp.Text = 0.0#
        If Len(Trim(txtTool.Text)) = 0 Then txtTool.Text = 0.0#
        '-----------------------------------------------------------------------------------
        'Modified BY         : SAURAV KUMAR
        'Modified ON         : 19 JULY 2011
        'ISSUE ID            : 10117300
        'DESCRIPTION         : Addition of GEM and EAN functionality
        '-----------------------------------------------------------------------------------
        If chkGEM.Checked = True Then
            If Len(Trim(txtEanNo.Text)) > 13 Then
                MsgBox("EAN No Not greater than 13-digit for Item Code " & txtItemCode.Text)
                Exit Function
            End If
        End If
        If ctlCustItem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            If chkGEM.Checked = True Then
                If checkgem_code() <> 0 Then
                    Exit Function
                End If
            End If
            If Me.txtBinqty.Text.Trim = "" Then
                Me.txtBinqty.Text = 0
            End If
            'Changed for Issue ID eMpro-20090512-31252 Starts--Add Commodity
            strSQL = "INSERT INTO CUSTITEM_MST(Unit_Code,dockcode,Account_Code,Cust_Drgno,Drg_Desc,Item_code,Item_Desc,Shop_Name,Gate_No,Container,Container_qty,Active,"
            'Code changed By Arul on 05-02-2005
            strSQL = strSQL & "Ent_dt,Ent_UserID,Upd_dt,Upd_UserId,VARMODEL,CUST_MTRL,TOOL_COST,schupldreqd,commodity,Delivery_Pattern,POF,BinQuantity,GEM_CODE,SHOP_CODE,PartType,Fixed_Code,Length,Line,IsFullScanning,IsNissanLabel)"
            'Addition Ends Here
            strSQL = strSQL & " VALUES('" & gstrUNITID & "','"
            strSQL = strSQL & Trim(txtdockcode.Text) & "','"
            strSQL = strSQL & prmAccode & "','"
            strSQL = strSQL & Trim(txtdrgno.Text) & "', '" & Trim(txtdrgdes.Text) & "','"
            strSQL = strSQL & prmitcode & "','" & Trim(lbldescription.Text) & "','" & Trim(txtShop.Text) & "','" & Trim(txtGate.Text) & "','" & Trim(txtcontainer.Text) & "'," & txtcontainerqty.Text & ", '" & mactfg & "'"
            strSQL = strSQL & ", Getdate(), '"
            'strSQL = strSQL & mP_User & "', Getdate(), '" & mP_User & "','" & Trim(TxtModel.Text) & "', " & Trim(txtCustSupp.Text) & "," & Trim(txtTool.Text) & ")"
            strSQL = strSQL & mP_User & "', Getdate(), '" & mP_User & "','" & Replace(Trim(TxtModel.Text), "'", "") & _
            "', " & Trim(txtCustSupp.Text) & "," & Trim(txtTool.Text) & ",'" & chkSchUpldReqd.Checked & _
            "','" & Trim(txtcommodity.Text) & "','" + Me.cmbDeliveryPattrn.Text.Trim + "','" + Me.DTPPOFTime.Text.Trim + "'," + Me.txtBinqty.Text + ",'"
            '"','" & Trim(txtBxFixedCd.Text) & "','" & Trim(txtBxLength.Text) & "','" + Me.cmbLine.Text.Trim + "'"
            If chkGEM.Checked = True Then
                strSQL = strSQL & Trim(txtEanNo.Text) & "','"
            Else
                strSQL = strSQL & "0" & "','"
            End If
            If AllowShopCodeflag(txtCustCode.Text) = True Then
                strSQL = strSQL & Trim(txtShopCode.Text) & "','" & cboItmtype.Text.Trim & "'"
            Else
                strSQL = strSQL & "0" & "','" & cboItmtype.Text.Trim & "'"
            End If
            Dim iCnt As Integer
            Dim intFound As Integer = 0

            Dim strLineItems As String
            Dim ChlLst As System.Windows.Forms.CheckedListBox.CheckedItemCollection
            ChlLst = Me.LineItems.CheckedItems
            For iCnt = 0 To ChlLst.Count - 1
                strLineItems = strLineItems + "" & Trim(ChlLst(iCnt).ToString) & ","
            Next

            If ChlLst.Count > 0 AndAlso strLineItems IsNot Nothing Then '24052024
                strLineItems = strLineItems.Remove(strLineItems.LastIndexOf(","))
            End If



            strSQL = strSQL & ",'" & Trim(txtBxFixedCd.Text) & "','" & txtBxLength.Text & "','" + Trim(strLineItems) + "','" & Me.chkBxFullScan.Checked & "','" & Me.chkBxNissanLbl.Checked & "')"
            If mactfg1 = 1 Then 'auto_invoice_part update
                strSQL = strSQL & "Update Item_mst Set AUTO_INVOICE_PART = 1 where unit_Code='" & gstrUNITID & "' and  Item_code = '" & prmitcode & "'  "
            End If

            'Changed for Issue ID eMpro-20090512-31252 Ends
            Call EditUpdateProdPrice("Add")
            If blnAllowBudget = True Then
                strBUDGETSQL = "Insert into budgetitem_mst(Unit_Code,account_code,cust_drgno,item_code,colour_code,category_code,commodity_code,model_code,usage_qty,"
                strBUDGETSQL = strBUDGETSQL & " ent_dt,ent_userid,upd_dt,upd_userid,Variant_Code,DefaultModel,EndDate) select unit_Code, account_code,cust_drgno,item_code,colour_code,category_code,"
                strBUDGETSQL = strBUDGETSQL & " commodity_code,model_code,usage_qty,ent_dt,ent_userid,ent_dt,upd_userid,Variant_Code,DefaultModel,'" & DTPEndDt.Value & "' from tmp_budgetitem_mst where  unit_Code='" & gstrUNITID & "' and  ip_address='" & gstrIpaddressWinSck & "' And Account_Code='" & txtCustCode.Text.Trim & "' And Item_Code='" & txtItemCode.Text.Trim & "' and cust_drgno='" & txtdrgno.Text.Trim & "' "
            End If
        ElseIf ctlCustItem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            'If chkGEM.Checked = True Then
            'If checkgem_code() <> 0 Then
            'Exit Function
            'End If
            'End If
            'updation of DrgNo and DrgDesc is removed from the query by shubhra
            '1054312
            If Me.txtBinqty.Text.Trim = "" Then
                Me.txtBinqty.Text = 0
            End If

            Dim strLineItems As String
            Dim iCnt As Integer
            Dim ChlLst As System.Windows.Forms.CheckedListBox.CheckedItemCollection
            ChlLst = Me.LineItems.CheckedItems
            For iCnt = 0 To ChlLst.Count - 1
                strLineItems = strLineItems + "" & Trim(ChlLst(iCnt).ToString) & ","
            Next
            If ChlLst.Count > 0 AndAlso strLineItems IsNot Nothing Then
                strLineItems = strLineItems.Remove(strLineItems.LastIndexOf(","))
            End If


            strSQL = "UPDATE CustItem_mst set Shop_Name='" & Trim(txtShop.Text) & "'," & _
                    " Gate_No='" & Trim(txtGate.Text) & "',container='" & (txtcontainer.Text) & "',Container_qty=" & txtcontainerqty.Text & ", Active = '" & mactfg & "'"
            strSQL = strSQL & ", Upd_dt=Getdate(),upd_userid='"
            strSQL = strSQL & mP_User & "',Commodity='" & Trim(txtcommodity.Text) & "',VARMODEL = '"
            strSQL = strSQL & Replace(Trim(TxtModel.Text), "'", "") & "',CUST_MTRL = " & Trim(txtCustSupp.Text) & ", TOOL_COST = " & Trim(txtTool.Text) & " ,SCHUPLDREQD = '" & chkSchUpldReqd.Checked & "' ,TORISHI_CODE = '" & txtTorishiCode.Text & "'"

            'strSQL = "SELECT ISNULL(PACO_FUNCTIONALITY,0) from Sales_parameter WHERE UNIT_CODE ='" & gstrUNITID & "'"
            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("SELECT ISNULL(PACO_FUNCTIONALITY,0) from Sales_parameter WHERE UNIT_CODE ='" & gstrUNITID & "'")) = True Then
                strSQL = strSQL & ", Fixed_Code='" & Trim(txtBxFixedCd.Text) & "',Length = " + Me.txtBxLength.Text + ", Line = '" & Trim(strLineItems) & "' , IsFullScanning = '" & chkBxFullScan.Checked & "', IsNissanLabel = '" & chkBxNissanLbl.Checked & "'"
            End If


            strSQL = strSQL & ", Delivery_Pattern='" + Me.cmbDeliveryPattrn.Text.Trim + "',POF='" + Me.DTPPOFTime.Text + "',BinQuantity=" + Me.txtBinqty.Text + ",GEM_Code='"
            If chkGEM.Checked = True Then
                strSQL = strSQL & Trim(txtEanNo.Text) & "',"
            Else
                strSQL = strSQL & "0'" & ","
            End If
            strSQL = strSQL & "DOCKCODE ='" & Me.txtdockcode.Text.Trim & "',"
            If AllowShopCodeflag(txtCustCode.Text) = True Then
                strSQL = strSQL & "SHOP_CODE='" & Me.txtShopCode.Text.Trim & "'"
            Else
                strSQL = strSQL & "SHOP_CODE=''"
            End If
            strSQL = strSQL & ",PartType = '" & cboItmtype.Text.Trim & "' "
            strSQL = strSQL & " Where  unit_Code='" & gstrUNITID & "' and  Account_Code = '" & prmAccode & "' and Item_Code ='"
            strSQL = strSQL & prmitcode & "'"
            Call EditUpdateProdPrice("Edit")
            If mactfg1 = 1 Then  'auto_invoice_part update
                strSQL = strSQL & "Update Item_mst Set AUTO_INVOICE_PART = 1 where unit_Code='" & gstrUNITID & "' and  Item_code = '" & prmitcode & "'  "
            End If
            If CheckForItemMainGroup(prmitcode) = True Then
                If blnAllowBudget = True Then
                    '10808160
                    'If IsRecordExists("select * from budgetitem_mst  where  unit_Code='" & gstrUNITID & "' and  account_code ='" & prmAccode & "' and item_code='" & prmitcode & "' and cust_drgno='" & txtdrgno.Text.Trim & "'") Then
                    If DTPEndDt.Enabled = False Then


                        strBUDGETSQL = "Insert into budgetitem_mst(Unit_Code,account_code,cust_drgno,item_code,colour_code,category_code,commodity_code,model_code,usage_qty,"
                        strBUDGETSQL = strBUDGETSQL & " ent_dt,ent_userid,upd_dt,upd_userid,Variant_Code,DefaultModel,EndDate) select Unit_Code,account_code,cust_drgno,item_code,colour_code,category_code,"
                        strBUDGETSQL = strBUDGETSQL & " commodity_code,model_code,usage_qty,ent_dt,ent_userid,ent_dt,upd_userid,Variant_Code,DefaultModel,EndDate from tmp_budgetitem_mst where  unit_Code='" & gstrUNITID & "' and ip_address='" & gstrIpaddressWinSck & "' And Account_Code='" & txtCustCode.Text.Trim & "' And Item_Code='" & txtItemCode.Text.Trim & "' "
                    Else
                        strBUDGETSQL = "Insert into budgetitem_mst(Unit_Code,account_code,cust_drgno,item_code,colour_code,category_code,commodity_code,model_code,usage_qty,"
                        strBUDGETSQL = strBUDGETSQL & " ent_dt,ent_userid,upd_dt,upd_userid,Variant_Code,DefaultModel,EndDate) select Unit_Code,account_code,cust_drgno,item_code,colour_code,category_code,"
                        strBUDGETSQL = strBUDGETSQL & " commodity_code,model_code,usage_qty,ent_dt,ent_userid,ent_dt,upd_userid,Variant_Code,DefaultModel,'" & DTPEndDt.Value & "' from tmp_budgetitem_mst where  unit_Code='" & gstrUNITID & "' and ip_address='" & gstrIpaddressWinSck & "' And Account_Code='" & txtCustCode.Text.Trim & "' And Item_Code='" & txtItemCode.Text.Trim & "' "

                    End If

                    'Else
                    'If (DTPEndDt.Enabled = True) Then
                    '    strBUDGETSQL = "update budgetitem_mst  "
                    '    strBUDGETSQL &= "set  EndDate ='" & DTPEndDt.Value & "' where unit_Code='" & gstrUNITID & "' and  account_code ='" & prmAccode & "' and item_code='" & prmitcode & "' and cust_drgno='" & txtdrgno.Text.Trim & "' "
                    'End If
                    'End If
                End If
            End If
        End If
        If Len(strSQL) > 0 Then
            'mP_Connection.Close()
            'mP_Connection.Open()
            ResetDatabaseConnection()
            mP_Connection.BeginTrans()
            mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            If blnAllowBudget = True And Len(strBUDGETSQL) > 0 Then
                mP_Connection.Execute("Delete from budgetitem_mst where  unit_Code='" & gstrUNITID & "' and  account_code ='" & prmAccode & "' and item_code='" & prmitcode & "' and cust_drgno='" & txtdrgno.Text.Trim & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                mP_Connection.Execute(strBUDGETSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            mP_Connection.CommitTrans()
            clsADOrs.ResultSetClose()
            clsADOrs = New ClsResultSetDB
            clsADOrs.GetResult("Select * from CustItem_Mst where  unit_Code='" & gstrUNITID & "' and Account_code ='" & txtCustCode.Text & "' and Item_Code ='" & txtItemCode.Text & "' and  cust_drgno = '" & Me.txtdrgno.Text & "'")
            If blnAllowBudget = True Then
                mP_Connection.Execute("delete from tmp_budgetitem_mst where unit_Code='" & gstrUNITID & "' and  ip_address='" & gstrIpaddressWinSck & "' And Account_Code='" & txtCustCode.Text.Trim & "' And Item_Code='" & txtItemCode.Text.Trim & "' ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            SaveData = True
        End If
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mP_Connection.RollbackTrans()
    End Function
    Private Function ValidRecord() As Boolean
        '----------------------------------------------------------------------------
        'Argument       :
        'Return Value   :   Boolean
        'Function       :   Valid Check
        'Comments       :   used to check validity of data before Save
        '----------------------------------------------------------------------------
        On Error GoTo Err_Handler
        Dim blnInvalidData As Boolean
        Dim strErrMsg As String
        Dim ctlBlank As System.Windows.Forms.Control = Nothing
        Dim lNo As Integer
        Dim intMaxCount, intCount As Integer
        Dim rstmp_budgetitem_mst As ClsResultSetDB
        ValidRecord = False
        lNo = 1
        strErrMsg = ResolveResString(10059) & vbCrLf & vbCrLf
        If Len(Trim(txtCustCode.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & ". Customer Code"
            If ctlBlank Is Nothing Then ctlBlank = txtCustCode
            lNo = lNo + 1
        End If
        If Len(Trim(txtItemCode.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & ". Item Code"
            If ctlBlank Is Nothing Then ctlBlank = txtItemCode
            lNo = lNo + 1
        End If
        '-----------------------------------------------------------------------------------
        'Added BY         : SAURAV KUMAR
        'Added ON         : 19 JULY 2011
        'ISSUE ID         : 10117300
        'DESCRIPTION      : Addition of GEM and EAN functionality
        '-----------------------------------------------------------------------------------
        If chkGEM.Checked = True Then
            If Len(Trim(txtEanNo.Text)) = 0 Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & ". EAN No"
                If ctlBlank Is Nothing Then ctlBlank = txtEanNo
                lNo = lNo + 1
            End If
        End If
        If Len(Trim(txtdrgno.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & ". CustDrgNo. "
            If ctlBlank Is Nothing Then ctlBlank = txtdrgno
            lNo = lNo + 1
        End If
        If Len(Trim(txtdrgdes.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & ". Drg Description "
            If ctlBlank Is Nothing Then ctlBlank = Me.txtdrgdes
            lNo = lNo + 1
        End If
        strSQL = "SELECT ISNULL(PACO_FUNCTIONALITY,0) from Sales_parameter WHERE UNIT_CODE ='" & gstrUNITID & "'"
        If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
            If Len(Trim(txtBxFixedCd.Text)) = 0 Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & ".Fixed Code "
                If ctlBlank Is Nothing Then ctlBlank = Me.txtBxFixedCd
                lNo = lNo + 1
                'Set visible property true if Invoice Address frame is not visible
            End If
            If Len(Trim(txtBxLength.Text)) = 0 Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & ".Length "
                If ctlBlank Is Nothing Then ctlBlank = Me.txtBxLength
                lNo = lNo + 1
                'Set visible property true if Invoice Address frame is not visible
            End If
            If LineItems.CheckedItems.Count < 1 Then
                blnInvalidData = True
                'strErrMsg = "Please select Line Item."
                strErrMsg = strErrMsg & vbCrLf & lNo & ".Line "
                If ctlBlank Is Nothing Then ctlBlank = Me.LineItems
                lNo = lNo + 1
            End If
        End If

        If CheckForItemMainGroup(txtItemCode.Text) = True Then
            If blnAllowBudget = True Then
                If Len(Trim(txtcolour.Text)) = 0 Then
                    blnInvalidData = True
                    strErrMsg = strErrMsg & vbCrLf & lNo & ". Colour Code"
                    If ctlBlank Is Nothing Then ctlBlank = Me.txtcolour
                    lNo = lNo + 1
                    'Set visible property true if Invoice Address frame is not visible
                End If
                If Len(Trim(txtcategory.Text)) = 0 Then
                    blnInvalidData = True
                    strErrMsg = strErrMsg & vbCrLf & lNo & ". Category  Code"
                    If ctlBlank Is Nothing Then ctlBlank = Me.txtcategory
                    lNo = lNo + 1
                    'Set visible property true if Invoice Address frame is not visible
                End If
                If Len(Trim(txtcommod.Text)) = 0 Then
                    blnInvalidData = True
                    strErrMsg = strErrMsg & vbCrLf & lNo & ".Commodity Code"
                    If ctlBlank Is Nothing Then ctlBlank = Me.txtcommodity
                    lNo = lNo + 1
                    'Set visible property true if Invoice Address frame is not visible
                End If


                rstmp_budgetitem_mst = New ClsResultSetDB
                rstmp_budgetitem_mst.GetResult("select * from tmp_budgetitem_mst where unit_Code='" & gstrUNITID & "' and ip_address = '" & gstrIpaddressWinSck & "' And Account_Code='" & txtCustCode.Text.Trim & "' And Item_Code='" & txtItemCode.Text.Trim & "'")
                intMaxCount = rstmp_budgetitem_mst.RowCount
                rstmp_budgetitem_mst.MoveFirst()
                If intMaxCount = 0 Then
                    blnInvalidData = True
                    strErrMsg = strErrMsg & vbCrLf & lNo & ". Model Code"
                    lNo = lNo + 1
                End If
                rstmp_budgetitem_mst.ResultSetClose()
                rstmp_budgetitem_mst = Nothing
            End If
        End If
        Dim rs As New ClsResultSetDB
        If Trim(Len(TxtModel.Text)) > 0 And Not (ctlCustItem.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT Or ctlCustItem.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) Then
            rs.GetResult("SELECT DISTINCT A.VARMODEL,A.ACCOUNT_CODE,B.CUST_NAME FROM CUSTITEM_MST A,CUSTOMER_MST B WHERE A.unit_Code = B.unit_Code and  A.unit_Code='" & gstrUNITID & "' and  A.ACCOUNT_CODE = B.CUSTOMER_CODE AND A.VARMODEL = '" & Replace(Trim(TxtModel.Text), "'", "") & "'")
            If Not rs.GetNoRows > 0 Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & ". Model Code "
                If ctlBlank Is Nothing Then ctlBlank = Me.txtcommod
                lNo = lNo + 1
            End If
            rs.ResultSetClose()
            rs = Nothing
        End If
        'changes by Prashant Rajpal
        If AllowShopCodeflag(Me.txtCustCode.Text) = True Then
            If Len(Trim(txtShopCode.Text)) = 0 Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & ".Shop Code"
                If ctlBlank Is Nothing Then ctlBlank = Me.txtShopCode
                lNo = lNo + 1
            End If
            If Len(Trim(txtShopCode.Text)) > 0 And IsNumeric(txtShopCode.Text) = True Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & ".Alphanumeric Shop Code"
                If ctlBlank Is Nothing Then ctlBlank = Me.txtShopCode
                lNo = lNo + 1
            End If
        End If
        'changes done by prashant Rajpal
        ''10856126'
        strSQL = "select dbo.UDF_ISDOCKCODE_ASN( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "')"
        If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
            If Len(Trim(txtdockcode.Text)) = 0 Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & ". .DOCK Code "
                If ctlBlank Is Nothing Then ctlBlank = Me.txtdockcode
                lNo = lNo + 1
            End If
        End If
        ''10856126
        '27 jan 2017
        strSQL = "SELECT AllowMultiplePartCodes from Sales_parameter WHERE UNIT_CODE ='" & gstrUNITID & "'"
        If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = False Then
            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("SELECT TOP 1 1 From CustItem_Mst WHERE Account_Code='" & Trim(txtCustCode.Text) & "' AND Item_code='" & Trim(txtItemCode.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'AND Active=1")) = True Then

                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & ".Already Linked with different Internal Code ."
                If ctlBlank Is Nothing Then ctlBlank = Me.txtItemCode
                lNo = lNo + 1
            End If
        End If


        strSQL = "SELECT AllowMultipleDrawingNo from Sales_parameter WHERE UNIT_CODE ='" & gstrUNITID & "'"
        If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = False Then
            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("SELECT TOP 1 1 From CustItem_Mst WHERE Account_Code='" & Trim(txtCustCode.Text) & "' AND cust_drgno='" & Trim(txtdrgno.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'AND Active=1")) = True Then

                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & ".Already Linked with different Cust Part Code ."
                If ctlBlank Is Nothing Then ctlBlank = Me.txtdrgno
                lNo = lNo + 1
            End If

        End If

        '27 jan 2017

        strErrMsg = VB.Left(strErrMsg, Len(strErrMsg) - 1)
        strErrMsg = strErrMsg & "."
        If blnInvalidData = True Then
            gblnCancelUnload = True : gblnFormAddEdit = True
            Call MsgBox(strErrMsg, MsgBoxStyle.Information, "Error")
            Exit Function
        End If
        ValidRecord = True
        gblnCancelUnload = True
        gblnFormAddEdit = True
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function PrimaryKeycheck() As Boolean
        '----------------------------------------------------------------------------
        'Argument       :
        'Return Value   :   Boolean
        'Function       :   PrimaryKeycheck
        'Comments       :   checks the PrimaryKeyConstraint
        '----------------------------------------------------------------------------
        Dim strSQL As String
        Dim rsCustItemMst As ClsResultSetDB
        On Error GoTo Err_Handler
        PrimaryKeycheck = False
        strSQL = "select * from CustItem_mst where  unit_Code='" & gstrUNITID & "' and  Account_Code = '" & txtCustCode.Text.Trim & "' and "
        strSQL = strSQL & "Item_code = '" & txtItemCode.Text.Trim & "' And cust_drgno = '" & txtdrgno.Text.Trim & "'"
        rsCustItemMst = New ClsResultSetDB
        rsCustItemMst.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If rsCustItemMst.GetNoRows > 0 Then
            PrimaryKeycheck = False
            MessageBox.Show("Record already exists.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.txtItemCode.Focus()
            Exit Function
        End If
        PrimaryKeycheck = True
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub txtconainerqty_Change()
        txtcontainerqty.Text = Number_Chk(txtcontainerqty.Text)
    End Sub
    Private Sub txtcontainer_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtcontainer.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Me.txtcontainerqty.Focus()
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCustCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustCode.TextChanged
        If Len(Trim(txtCustCode.Text)) = 0 Then
            Select Case ctlCustItem.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    Call EnableControls(False, Me, True)
                    LineItems.Enabled = False
                    txtCustCode.Enabled = True : txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdhelp.Enabled = True
                    lblcustcode.Text = "" : lbldescription.Text = ""
                    lblDate.Text = "" : lblEff_dt1.Text = ""
                    lblEff_dt2.Text = "" : lblPrice1.Text = "" : lblPrice2.Text = ""
                    Me.ctlCustItem.Enabled(1) = False
                    Me.ctlCustItem.Enabled(2) = False
                    Me.ctlCustItem.Enabled(5) = False
                    txtCustCode.Focus()
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    lblcustcode.Text = ""
            End Select
        Else
            txtItemCode.Enabled = True : txtItemCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdhelp1.Enabled = True
        End If
    End Sub
    Private Sub txtCustCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustCode.Enter
        txtCustCode.SelectionStart = 0 : txtCustCode.SelectionLength = Len(txtCustCode.Text)
    End Sub
    Private Sub txtCustCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.ctlCustItem.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                        If Len(txtCustCode.Text) > 0 Then
                            Call txtCustCode_Validating(txtCustCode, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            Me.ctlCustItem.Focus()
                        End If
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Len(txtCustCode.Text) > 0 Then
                            Call txtCustCode_Validating(txtCustCode, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            txtItemCode.Focus()
                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCustCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            If cmdhelp.Enabled Then Call cmdHelp_Click(cmdhelp, New System.EventArgs())
        End If
    End Sub
    Private Sub txtCustCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        If Len(Trim(txtCustCode.Text)) = 0 Then
            GoTo EventExitSub
        Else
            Select Case ctlCustItem.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    If SelectQuery("distinct(cust_name)", "Customer_Mst", " where  unit_Code='" & gstrUNITID & "' and  customer_Code ='" & Trim(txtCustCode.Text) & "'") = False Then
                        Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_EXCLAMATION)
                        txtCustCode.Text = ""
                        Cancel = True
                        GoTo EventExitSub
                    Else
                        txtItemCode.Focus()
                    End If
                    If AllowShopCodeflag(Me.txtCustCode.Text) = True Then
                        lblShopcode.Visible = True
                        txtShopCode.Visible = True
                    Else
                        lblShopcode.Visible = False
                        txtShopCode.Visible = False
                    End If
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    If SelectQuery("distinct(cust_name)", "Customer_Mst", " where  unit_Code='" & gstrUNITID & "' and  customer_code ='" & Trim(txtCustCode.Text) & "'") = False Then
                        Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_EXCLAMATION)
                        txtCustCode.Text = "" : Cancel = True : GoTo EventExitSub
                    Else
                        If txtItemCode.Enabled = True Then
                            txtItemCode.Focus()
                        Else
                            ctlCustItem.Focus()
                        End If
                    End If
                    If AllowShopCodeflag(Me.txtCustCode.Text) = True Then
                        lblShopcode.Visible = True
                        txtShopCode.Visible = True
                    Else
                        lblShopcode.Visible = False
                        txtShopCode.Visible = False
                    End If
            End Select
            mvalid = True
        End If
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtdrgdes_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtdrgdes.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case ctlCustItem.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        chkflag.Focus()
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtDrgNo_KeyPress(ByVal Sender As System.Object, ByVal e As CtlGeneral.KeyPressEventArgs) Handles txtdrgno.KeyPress
        Dim KeyAscii As Short = e.KeyAscii
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                If Len(Trim(txtdrgno.Text)) > 0 Then
                    Select Case ctlCustItem.Mode
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                            txtdrgdes.Focus()
                    End Select
                Else
                    txtdrgdes.Focus()
                End If
            Case 39, 34, 96
                KeyAscii = 0
        End Select
    End Sub
    Public Sub DispRecordfromProdPrice()
        '----------------------------------------------------------------------------
        'Argument       :
        'Return Value   :   Nil
        'Function       :   DispRecordFromProdPrice
        'Comments       :   Used to display details from ProdPrice_Mst
        '----------------------------------------------------------------------------
        Dim mRdoCls As New ClsResultSetDB
        Dim strProdPrice As String
        strProdPrice = "SELECT * FROM prod_price_mst " & "WHERE  unit_Code='" & gstrUNITID & "' and  UPPER(cust_C) = '" & Trim(txtCustCode.Text) & "' AND Product_no='" & Trim(txtItemCode.Text) & "'"
        mRdoCls = New ClsResultSetDB
        If Not mRdoCls.GetResult(strProdPrice, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) Then Exit Sub
        If mRdoCls.GetNoRows = 0 Then
            MsgBox("Item Price not defined in Product price master", MsgBoxStyle.OkOnly, "empower")
            Exit Sub
        End If
        If mRdoCls.GetValue("eff_dt1") = "" Then
            lblEff_dt1.Text = ""
        Else
            lblEff_dt1.Text = mRdoCls.GetValue("eff_dt1")
        End If
        If mRdoCls.GetValue("eff_dttill") = "" Then
            lblEff_dt2.Text = ""
        Else
            lblEff_dt2.Text = mRdoCls.GetValue("eff_dttill")
        End If
        lblPrice1.Text = mRdoCls.GetValue("price1")
        txtPrice.Text = ""
    End Sub
    Public Sub EditUpdateProdPrice(ByRef mode As String)
        '----------------------------------------------------------------------------
        'Argument       :
        'Return Value   :   Nil
        'Function       :   DispRecordFromProdPrice
        'Comments       :   Used to Insert/Update details to ProdPrice_Mst
        '----------------------------------------------------------------------------
        Dim effdate2 As String
        effdate2 = VB6.Format(lblDate.Text, "mm/dd/yyyy")
        Select Case mode
            Case "Add"
                strProdPrice = "INSERT INTO prod_price_mst(Unit_Code,product_no, Cust_C,"
                strProdPrice = strProdPrice & " eff_dt1, eff_dt2,"
                strProdPrice = strProdPrice & " price1, price2,"
                strProdPrice = strProdPrice & " ent_Userid,ent_dt,Upd_UserID,upd_dt) VALUES ('" & gstrUNITID & "','"
                strProdPrice = strProdPrice & Trim(txtItemCode.Text) & "','" & Trim(txtCustCode.Text)
                strProdPrice = strProdPrice & "',Null, getdate(),"
                strProdPrice = strProdPrice & " Null," & IIf(Len(Trim(txtPrice.Text)) = 0, 0, Trim(txtPrice.Text))
                strProdPrice = strProdPrice & ",'"
                strProdPrice = strProdPrice & Trim(mP_User) & "',GETDATE(),'" & Trim(mP_User)
                strProdPrice = strProdPrice & "',GETDATE()" & ")"
            Case "Edit"
                strProdPrice = "UPDATE prod_price_mst SET eff_dt1='" & Trim(lblEff_dt2.Text) & "',eff_dt2 ='" & Trim(lblDate.Text) & "',price1 = " & IIf(Len(Trim(lblPrice2.Text)) = 0, 0, Trim(lblPrice2.Text)) & ",price2=" & IIf(Len(Trim(txtPrice.Text)) = 0, 0, Trim(txtPrice.Text)) & ",Upd_userid='" & Trim(mP_User) & "',upd_dt=GETDATE()" & " WHERE  unit_Code='" & gstrUNITID & "' and  product_no ='" & Trim(txtItemCode.Text) & "' AND Cust_C='" & Trim(txtCustCode.Text) & "'"
        End Select
    End Sub
    Public Sub RefreshLablelofPordPrice()
        '----------------------------------------------------------------------------
        'Argument       :
        'Return Value   :   Nil
        'Function       :   DispRecordFromProdPrice
        'Comments       :   Used to Clear Label ProdPrice_Mst
        '----------------------------------------------------------------------------
        lblDate.Text = ""
        lblEff_dt1.Text = ""
        lblEff_dt2.Text = ""
        lblPrice1.Text = ""
        lblPrice2.Text = ""
        'Code Added By Arul on 05-02-2005
        lblcustcode.Text = ""
        lbldescription.Text = ""
        'Addition ends here
    End Sub
    Public Function SelectQuery(ByRef pstrFName As String, ByRef pstrTableName As String, Optional ByRef pstrCondition As String = "") As Boolean
        '----------------------------------------------------------------------------
        'Argument       :
        'Return Value   :   Nil
        'Function       :   DispRecordFromProdPrice
        'Comments       :   Used to Clear Label ProdPrice_Mst
        '----------------------------------------------------------------------------
        Dim strSQL As String
        Dim rsRecordset As ClsResultSetDB
        strSQL = "Select " & pstrFName & " from " & pstrTableName & " " & pstrCondition
        rsRecordset = New ClsResultSetDB
        rsRecordset.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsRecordset.GetNoRows = 0 Then
            SelectQuery = False
        Else
            SelectQuery = True
        End If
        rsRecordset.ResultSetClose()
        rsRecordset = Nothing
    End Function
    Private Sub txtGate_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtGate.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                txtcontainer.Focus()
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtItemCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtItemCode.TextChanged
        If Len(Trim(txtItemCode.Text)) = 0 Then
            Select Case ctlCustItem.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    txtdrgno.Text = "" : txtdrgdes.Text = ""
                    lblcustcode.Text = "" : lbldescription.Text = ""
                    lblDate.Text = "" : lblEff_dt1.Text = ""
                    lblEff_dt2.Text = "" : lblPrice1.Text = "" : lblPrice2.Text = ""
                    If blnAllowBudget = True Then
                        txtcolour.Text = ""
                        txtcategory.Text = ""
                        txtcommod.Text = ""
                    End If
                    Me.ctlCustItem.Enabled(1) = False
                    Me.ctlCustItem.Enabled(2) = False
                    Me.ctlCustItem.Enabled(5) = False
                    chkflag.CheckState = False
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    lbldescription.Text = ""
            End Select
        End If
    End Sub
    Private Sub txtItemCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtItemCode.Enter
        txtItemCode.SelectionStart = 0 : txtItemCode.SelectionLength = Len(txtItemCode.Text)
    End Sub
    Private Sub txtItemCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtItemCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case ctlCustItem.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                        If Len(txtItemCode.Text) > 0 Then
                            Call txtItemCode_Validating(txtItemCode, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            Me.ctlCustItem.Focus()
                        End If
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Len(txtItemCode.Text) > 0 Then
                            Call txtItemCode_Validating(txtItemCode, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            txtdrgno.Focus()
                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtItemCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtItemCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            If cmdhelp1.Enabled Then Call cmdHelp1_Click(cmdhelp1, New System.EventArgs())
        End If
    End Sub
    Private Sub txtItemCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtItemCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        If Len(Trim(txtItemCode.Text)) = 0 Then
            GoTo EventExitSub
        Else
            Select Case ctlCustItem.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    If SelectQuery("distinct(description)", "Item_Mst", " where  unit_Code='" & gstrUNITID & "' and  item_code = '" & Trim(txtItemCode.Text) & "'") = False Then
                        Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_EXCLAMATION)
                        txtItemCode.Text = "" : Cancel = True : GoTo EventExitSub
                    Else
                        txtdrgno.Focus()
                    End If
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    If SelectQuery("distinct(item_desc)", "custItem_Mst ", " where  unit_Code='" & gstrUNITID & "' and  item_code = '" & Trim(txtItemCode.Text) & "'") = False Then
                        'Commented and added by Shubhra Verma
                        'Issue ID : eMpro-20090122-26322
                        'if no record found system gives message "Transaction Completed Successfully"
                        'Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_EXCLAMATION)
                        Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_EXCLAMATION)
                        txtItemCode.Text = "" : Cancel = True : GoTo EventExitSub
                    Else
                        Call RefreshForm(True)

                        ctlCustItem.Enabled(1) = True
                        ctlCustItem.Enabled(2) = True
                        Me.ctlCustItem.Focus()
                    End If
            End Select
            mvalid = True
        End If
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtShop_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtShop.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                '        Me.ctlCustItem.SetFocus
                txtGate.Focus()
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtmodel_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TxtModel.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            If CmdModHlp.Enabled Then Call CmdModHlp_Click(CmdModHlp, New System.EventArgs())
        End If
    End Sub
    Private Sub txtModel_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtModel.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Me.txtCustSupp.Focus()
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub CmdModHlp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdModHlp.Click
        '----------------------------------------------------------------------------
        'Argument       :   NIL
        'Return Value   :   NIL
        'Function       :   To show help on Existing Model Lists
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        Dim varRetVal As Object
        Dim strHelp() As String
        Dim strQuery As String
        On Error GoTo Err_Handler
        Select Case ctlCustItem.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                With Me.TxtModel
                    strQuery = "SELECT DISTINCT A.VARMODEL,A.ACCOUNT_CODE,B.CUST_NAME FROM CUSTITEM_MST A,CUSTOMER_MST B WHERE A.unit_code=B.unit_Code and A.unit_Code='" & gstrUNITID & "' and  A.ACCOUNT_CODE = B.CUSTOMER_CODE AND A.VARMODEL IS NOT NULL AND A.VARMODEL <> '' ORDER BY A.VARMODEL,A.ACCOUNT_CODE"
                    strHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Model Lists")
                    .Focus()
                End With
                If UBound(strHelp) < 0 Then
                    Me.TxtModel.Text = "" : Exit Sub
                ElseIf strHelp(0) = "0" Then
                    Call ConfirmWindow(10070)
                    Me.TxtModel.Focus()
                    Exit Sub
                ElseIf UBound(strHelp) > 0 Then
                    Me.TxtModel.Text = Trim(strHelp(0))
                    Me.ctlCustItem.Focus()
                    Exit Sub
                End If
        End Select
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustSupp_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles txtCustSupp.KeyPress
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Me.txtTool.Focus()
        End Select
    End Sub
    Private Sub txtTool_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles txtTool.KeyPress
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Me.txtcommodity.Focus()
        End Select
    End Sub
    Private Sub txtPrice_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles txtPrice.KeyPress
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                '        Me.ctlCustItem.SetFocus
                txtShop.Focus()
        End Select
    End Sub
    Private Sub TxtModel_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TxtModel.Validating
        Dim rs As New ClsResultSetDB
        'Added by shubhra
        'Issue ID : eMpro-20090122-26322
        'Begin
        If ctlCustItem.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE Or ctlCustItem.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT Or ctlCustItem.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD Then
            rs.ResultSetClose()
            rs = Nothing
            Exit Sub
        End If
        If Trim(Len(TxtModel.Text)) > 0 Then
            rs.GetResult("SELECT DISTINCT A.VARMODEL,A.ACCOUNT_CODE,B.CUST_NAME FROM CUSTITEM_MST A,CUSTOMER_MST B WHERE A.unit_Code= B.unit_Code and A.unit_Code='" & gstrUNITID & "' and  A.ACCOUNT_CODE = B.CUSTOMER_CODE AND A.VARMODEL = '" & Replace(Trim(TxtModel.Text), "'", "") & "'")
            If Not rs.GetNoRows > 0 Then
                MsgBox("Invalid Model Code", , ResolveResString(100))
                TxtModel.Text = ""
                TxtModel.Focus()
                'commented by Shubhra Verma. Issue ID : eMpro-20090122-26322
                'rs.ResultSetClose()
                'rs = Nothing
                Exit Sub
            End If
            rs.ResultSetClose()
            rs = Nothing
        End If
    End Sub
    Private Sub txtcommodity_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtcommodity.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Me.ctlCustItem.Focus()
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim varHelpColour As Object
        varHelpColour = ShowList(1, 6, "", "colour_code", "colour_desc", "colour_mst", "and active=1 ")
        If varHelpColour = "-1" Then
            Call ConfirmWindow(10013, ConfirmWindowButtonsEnum.BUTTON_OK, ConfirmWindowImagesEnum.IMG_INFO)
        ElseIf varHelpColour = "" Or varHelpColour = String.Empty Then
            txtcolour.Text = ""
            Exit Sub
        Else
            txtcolour.Text = varHelpColour
        End If
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim varHelpCategory As Object
        varHelpCategory = ShowList(1, 6, "", "Category", "colour_desc", "colour_mst", " and colour_code ='" & txtcolour.Text.Trim & "' and active=1 ")
        If varHelpCategory = "-1" Then
            Call ConfirmWindow(10013, ConfirmWindowButtonsEnum.BUTTON_OK, ConfirmWindowImagesEnum.IMG_INFO)
        ElseIf varHelpCategory = "" Or varHelpCategory = String.Empty Then
            txtcategory.Text = ""
            Exit Sub
        Else
            txtcategory.Text = varHelpCategory
        End If
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim varHelpCommoity As Object
        varHelpCommoity = ShowList(1, 6, "", "commodity_code", "commodity_desc", "commodity_mst", "and active=1 ")
        If varHelpCommoity = "-1" Then
            Call ConfirmWindow(10013, ConfirmWindowButtonsEnum.BUTTON_OK, ConfirmWindowImagesEnum.IMG_INFO)
        ElseIf varHelpCommoity = "" Or varHelpCommoity = String.Empty Then
            txtcommod.Text = ""
            Exit Sub
        Else
            txtcommod.Text = varHelpCommoity
        End If
    End Sub
    Private Sub cmdbutton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdbutton4.Click
        Dim frmForm_modeldetails As FrmModelDetails
        If txtCustCode.Text.Trim.Length = 0 Then
            MessageBox.Show("Enter Customer code !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
            If txtCustCode.Enabled = True Then txtCustCode.Focus()
            Exit Sub
        End If
        If txtItemCode.Text.Trim.Length = 0 Then
            MessageBox.Show("Enter Item code !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
            If txtItemCode.Enabled = True Then txtItemCode.Focus()
            Exit Sub
        End If
        frmForm_modeldetails = New FrmModelDetails
        frmForm_modeldetails.Customercode = txtCustCode.Text
        frmForm_modeldetails.Itemcode = txtItemCode.Text
        frmForm_modeldetails.Custdrgno = txtdrgno.Text
        frmForm_modeldetails.Colourcode = txtcolour.Text
        frmForm_modeldetails.Categorycode = txtcategory.Text
        frmForm_modeldetails.Commoditycode = txtcommod.Text
        frmForm_modeldetails.Mode = ctlCustItem.Mode
        frmForm_modeldetails.ShowDialog()
    End Sub
    Private Sub txtcolour_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtcolour.Validating
        Dim rs As New ClsResultSetDB
        If Trim(Len(txtcolour.Text)) > 0 Then
            rs.GetResult("SELECT * FROM COLOUR_MST WHERE  unit_Code='" & gstrUNITID & "' and  ACTIVE =1 AND  COLOUR_CODE = '" & Replace(Trim(txtcolour.Text), "'", "") & "'")
            If Not rs.GetNoRows > 0 Then
                MsgBox("Invalid Colour Code", , ResolveResString(100))
                txtcolour.Text = ""
                txtcolour.Focus()
                Exit Sub
            End If
            rs.ResultSetClose()
            rs = Nothing
        End If
    End Sub
    Private Sub txtcategory_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtcategory.Validating
        Dim rs As New ClsResultSetDB
        If Trim(Len(txtcategory.Text)) > 0 Then
            rs.GetResult("SELECT * FROM COLOUR_MST WHERE unit_Code='" & gstrUNITID & "' and  ACTIVE =1 AND Category = '" & Replace(Trim(txtcategory.Text), "'", "") & "'")
            If Not rs.GetNoRows > 0 Then
                MsgBox("Invalid Category Code", , ResolveResString(100))
                txtcategory.Text = ""
                txtcategory.Focus()
                Exit Sub
            End If
            rs.ResultSetClose()
            rs = Nothing
        End If
    End Sub
    Private Sub txtcommod_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtcommod.Validating
        Dim rs As New ClsResultSetDB
        If Trim(Len(txtcommod.Text)) > 0 Then
            rs.GetResult("SELECT * FROM COMMODITY_MST  WHERE unit_Code='" & gstrUNITID & "' and  ACTIVE =1 AND COMMODITY_CODE = '" & Replace(Trim(txtcommod.Text), "'", "") & "'")
            If Not rs.GetNoRows > 0 Then
                MsgBox("Invalid Commodity Code", , ResolveResString(100))
                txtcommod.Text = ""
                txtcommod.Focus()
                Exit Sub
            End If
            rs.ResultSetClose()
            rs = Nothing
        End If
    End Sub
    Private Function CheckForItemMainGroup(ByVal Item_code As String) As Boolean
        Dim rstHelpDb As ClsResultSetDB
        CheckForItemMainGroup = False
        Try
            rstHelpDb = New ClsResultSetDB
            Call rstHelpDb.GetResult("Select * from item_mst where unit_Code='" & gstrUNITID & "' and  item_main_grp in ('F','S') and item_code = '" & Item_code & "' ", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rstHelpDb.GetNoRows >= 1 Then
                CheckForItemMainGroup = True
            End If
            rstHelpDb.ResultSetClose()
            rstHelpDb = Nothing
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    Private Sub ctlCustItem_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ctlCustItem.Load
    End Sub
    '-----------------------------------------------------------------------------------
    'Added BY         : SAURAV KUMAR
    'Added ON         : 19 JULY 2011
    'ISSUE ID         : 10117300
    'DESCRIPTION      : Addition of GEM and EAN functionality
    '-----------------------------------------------------------------------------------    
    Private Sub chkGEM_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkGEM.CheckedChanged
        If chkGEM.Checked = True Then
            txtEanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            txtEanNo.Enabled = True
        Else
            txtEanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtEanNo.Enabled = False
        End If
    End Sub
    '-----------------------------------------------------------------------------------
    'Added BY         : SAURAV KUMAR
    'Added ON         : 19 JULY 2011
    'ISSUE ID         : 10117300
    'DESCRIPTION      : Addition of GEM and EAN functionality
    '-----------------------------------------------------------------------------------
    Private Sub txtEanNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEanNo.KeyPress
        If (Microsoft.VisualBasic.Asc(e.KeyChar) < 48) Or (Microsoft.VisualBasic.Asc(e.KeyChar) > 57) Then
            e.Handled = True
        End If
        If (Microsoft.VisualBasic.Asc(e.KeyChar) = 8) Then
            e.Handled = False
        End If
    End Sub
    Private Function AllowShopCodeflag(ByVal pstraccoutncode As String) As Boolean
        On Error GoTo ErrHandler
        Dim strQry As String
        Dim Rs As ClsResultSetDB
        AllowShopCodeflag = False
        strQry = "Select isnull(Allow_ShopCode,0) as Allow_ShopCode from customer_mst where  unit_Code='" & gstrUNITID & "' and  Customer_Code='" & Trim(pstraccoutncode) & "'"
        Rs = New ClsResultSetDB
        If Rs.GetResult(strQry) = False Then GoTo ErrHandler
        If Rs.GetValue("Allow_ShopCode") = "True" Then
            AllowShopCodeflag = True
        Else
            AllowShopCodeflag = False
        End If
        Rs.ResultSetClose()
        Rs = Nothing
        Exit Function
ErrHandler:
        Rs = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub FillItemType()
        Dim strSQL As String
        Dim dt As DataTable
        Try
            cboItmtype.Items.Clear()
            strSQL = "SELECT KEY2 FROM LISTS WHERE UNIT_CODE='" & gstrUNITID & "' AND KEY1='ITEMTYPE' ORDER BY KEY2"
            dt = SqlConnectionclass.GetDataTable(strSQL)
            If dt.Rows.Count > 0 Then
                For Each row As DataRow In dt.Rows
                    cboItmtype.Items.Add(row("KEY2").ToString)
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    'Added Against 10737738
    Private Function ValidateVEHBOM() As Boolean
        Try
            Using sqlCmd As SqlCommand = New SqlCommand("USP_VCHBOM_MODEL_DTL_VALIDATE")
                sqlCmd.CommandType = CommandType.StoredProcedure
                sqlCmd.Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                sqlCmd.Parameters.Add("@ITEM_CODE", SqlDbType.VarChar, 16).Value = txtItemCode.Text.Trim
                sqlCmd.Parameters.Add("@CUST_CODE", SqlDbType.VarChar, 8).Value = Trim(txtCustCode.Text)
                sqlCmd.Parameters.Add("@CUST_DRGNO", SqlDbType.VarChar, 30).Value = Me.txtdrgno.Text.Trim
                sqlCmd.Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 50).Value = gstrIpaddressWinSck
                If (ctlCustItem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD) Then
                    sqlCmd.Parameters.Add("@MODE", SqlDbType.Bit).Value = 1
                Else
                    sqlCmd.Parameters.Add("@MODE", SqlDbType.Bit).Value = 0
                End If
                sqlCmd.Parameters.Add("@MSG", SqlDbType.VarChar, 5000).Direction = ParameterDirection.Output
                SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                If sqlCmd.Parameters("@MSG").Value.ToString.Trim.Length > 0 Then
                    MsgBox(sqlCmd.Parameters("@MSG").Value.ToString.Trim, MsgBoxStyle.Information, ResolveResString(100))
                    Return False
                End If
                Return True
            End Using

        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Private Sub FillLineList()
        Dim strSql As String
        Dim SqlRd As SqlDataReader
        Try
            strSql = "select Key2 from Lists WHERE Key1='QA_LINE' and UNIT_CODE='" & gstrUNITID & "'"
            SqlRd = SqlConnectionclass.ExecuteReader(strSql)
            If SqlRd.HasRows = True Then
                While SqlRd.Read
                    Me.LineItems.Items.Add(SqlRd.GetValue(SqlRd.GetOrdinal("Key2")))
                End While
            End If
            If SqlRd.IsClosed = False Then SqlRd.Close()
        Catch Ex As Exception
            RaiseException(Ex)
        Finally
            If Not SqlRd Is Nothing AndAlso SqlRd.IsClosed = False Then
                SqlRd.Close()
            End If
            SqlRd = Nothing

        End Try

    End Sub

End Class