Option Strict Off
Option Explicit On
Friend Class frmMKTMST0009
	Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
	' File Name         :   FRMMKTMST0008.frm
	' Function          :   Used to define customer supplied material details (Used only for mate chennai)
	' Created By        :   Arul Mozhi
	' Created On        :   20 April, 2001
	'===================================================================================
    'Revised By        : Manoj Kr. Vaish
    'Issue ID          : eMpro-20090625-32893
    'Revision Date     : 25 Jun 2009
    'History           : Record was not saving in Customer Supplier Item Master due to amount mismatch problem.
    '===================================================================================
    'Revised By        : Siddharth Ranjan
    'Issue ID          : eMpro-20090709-33428
    'Revision Date     : 09 Jul 2009
    'History           : CSI functionality
    'Modified By Nitin Mehta on 26 april 2011
    'Modified to support MultiUnit functionality
    '****************************************************************************************
    'Revised By        : Prashant Rajpal
    'Issue ID          : 10249069
    'Revision Date     : 11 Jul 2012
    'History           : CSI FUNCTIONALITY CHANGE- BOMEXPLOSION  PROCEDURE PARTAMETER CHANGE
    '****************************************************************************************

    Private Enum EnumCustItemGrid
        CustItemCode = 1
        CustItemDesc = 2
        CustDrgNo = 3
        Qty = 4
        Rate = 5
        Active = 6
        Valid_from = 7
        Valid_to = 8
        Data_type = 9
    End Enum
    Dim dblcustmtrl As Double
    Dim Msgdes As String
    Dim intCounter As Short
    Dim blnEdit_allowed As Boolean

    Dim Rs As New ADODB.Recordset
    Dim RsQty As New ADODB.Recordset
    Dim qry As String
    Dim mintFormIndex As Short
    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdhelp.Click
        On Error GoTo Errorhandler
        Dim strHelp As String
        Dim strHelp1() As String
        Dim rsResult As ClsResultSetDB
        If Me.ctlCustItem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            strHelp = ShowList(0, (Me.txtCustCode.MaxLength), , "Customer_Code", "Cust_Name", "Customer_mst")
            If strHelp = "-1" Then Exit Sub
            If strHelp = " " Then
                Call MsgBox("Customer Not Exist.Please Define New Customer In Customer Master", MsgBoxStyle.Information, "Help")
            Else
                Me.txtCustCode.Text = strHelp
                rsResult = New ClsResultSetDB
                Call rsResult.GetResult("Select Cust_name from customer_mst where Customer_Code = '" & Trim(strHelp) & "' and UNIT_CODE = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                Me.lblcustcode.Text = rsResult.GetValue("Cust_Name")
                rsResult = Nothing
                cmdhelp1.Enabled = True
                Me.txtItemCode.Enabled = True
                Me.txtItemCode.Focus()
            End If
        Else
            strHelp1 = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT DISTINCT A.CUSTOMER_CODE, B.CUST_NAME FROM CUSTSUPPLIEDITEM_MST A INNER JOIN CUSTOMER_MST B ON A.CUSTOMER_CODE = B.CUSTOMER_CODE and A.UNIT_CODE = B.UNIT_CODE WHERE A.UNIT_CODE = '" & gstrUNITID & "'", "Customer Help")
            If Not (UBound(strHelp1) = -1) Then
                If (Len(strHelp1(0)) >= 1) And strHelp1(0) = "0" Then
                    Call MsgBox("Customer Not Exist In Customer Supplier Master", MsgBoxStyle.Information, ResolveResString(100))
                Else
                    Me.txtCustCode.Text = strHelp1(0)
                    Me.lblcustcode.Text = strHelp1(1)
                    cmdhelp1.Enabled = True
                    Me.txtItemCode.Enabled = True
                    Me.txtItemCode.Focus()
                End If
            End If
        End If
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdHelp1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdhelp1.Click
        On Error GoTo Errorhandler
        Dim strhelp1 As String
        Dim strhelp() As String
        Dim rsResult As ClsResultSetDB
        If Me.ctlCustItem.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            strhelp1 = ShowList(0, (Me.txtItemCode.MaxLength), , " Finish_Item_Code", "description", "VW_CUSTOMERSUPPLIED_ITEMMASTER ", " and customer_code = '" & Trim(txtCustCode.Text) & "'")
            If strhelp1 = "" Then
                Call MsgBox("Item Not Exist.Please Define New Item In Item Master", MsgBoxStyle.Information, "Help")
            Else
                Call Showdetails(Trim(txtCustCode.Text), strhelp1)
            End If
        Else
            strhelp1 = "SELECT A.ITEM_CODE, A.CUST_DRGNO, A.DRG_DESC " & _
                        " FROM VW_CUST_SUPPLIED_ITEM_MST_FG_HELP A " & _
                        " INNER JOIN " & _
                        " (" & _
                                " SELECT FINISH_ITEM_CODE, CUST_DRGNO,UNIT_CODE" & _
                                " FROM CUSTSUPPLIEDITEM_MST" & _
                                " WHERE CUSTOMER_CODE = '" & txtCustCode.Text.Trim & "' and UNIT_CODE = '" & gstrUNITID & "'" & _
                                " AND ACTIVE_FLAG = 1 " & _
                                " GROUP BY FINISH_ITEM_CODE, CUST_DRGNO,UNIT_CODE" & _
                        " )B" & _
                        " ON (A.ITEM_CODE <> B.FINISH_ITEM_CODE" & _
                        " AND A.CUST_DRGNO <> B.CUST_DRGNO" & _
                        " AND A.UNIT_CODE = B.UNIT_CODE)" & _
                        " WHERE A.ACCOUNT_CODE = '" & txtCustCode.Text.Trim & "' and A.UNIT_CODE = '" & gstrUNITID & "'" & _
                        " GROUP BY A.ITEM_CODE, A.CUST_DRGNO, A.DRG_DESC "
            strhelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strhelp1)
            If Not (UBound(strhelp) = -1) Then
                If (Len(strhelp(0)) >= 1) And strhelp(0) = "0" Then
                    MsgBox("Item Does Not Exist.Please Define Customer Item Linking For Selected Customer", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Me.lbldescription.Text = ""
                    Me.TxtItemdrgNO.Text = ""
                    With Me.fpDTLSpread
                        .MaxRows = 1
                        .Row = 1 : .Row2 = 1 : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
                    End With
                    Exit Sub
                Else
                    Me.txtItemCode.Text = strhelp(0)
                    Me.TxtItemdrgNO.Text = strhelp(1)
                    Me.lbldescription.Text = strhelp(2)
                    Me.TxtItemdrgNO.Enabled = False
                    Me.txtItemCode.Enabled = False
                    Me.txtCustCode.Enabled = False
                    '10249069
                    mP_Connection.Execute("BOMExplosion '" & strhelp(0) & "','" & strhelp(0) & "',1, 0,0,0,'" & gstrIpaddressWinSck & "','" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    Call AddBlankRow()
                    With fpDTLSpread
                        .Row = 1
                        .Col = EnumCustItemGrid.Valid_from
                        .Col2 = EnumCustItemGrid.Data_type
                        .BlockMode = True
                        .Lock = True
                        .BlockMode = False
                        Call .SetText(EnumCustItemGrid.Data_type, .Row, "N")
                        .Row = 1
                        .Col = EnumCustItemGrid.CustItemCode
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        .Enabled = True
                        .Focus()
                    End With
                End If
            End If
        End If
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ctlCustItem_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles ctlCustItem.ButtonClick
        On Error GoTo Errorhandler
        Dim Finish_Item_Code As String
        Dim Item_Code As String
        Dim Item_drgno As String
        Dim Description As String
        Dim dblqty As Double
        Dim dblRate As Double
        Dim Active As String
        Dim SqlQry As String
        Dim strData_type As String
        Dim varActive As Object, intKounter As Integer
        Select Case ctlCustItem.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                Call EnableControls(True, Me, True)
                lblcustcode.Text = ""
                lbldescription.Text = ""
                fpDTLSpread.MaxRows = 0
                Me.txtCustCode.Focus()
                ctlCustItem.Enabled(2) = False
                blnEdit_allowed = False
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                If Trim(txtCustCode.Text) <> "" And Trim(txtItemCode.Text) <> "" Then
                    Call EnableControls(True, Me, False)
                    txtCustCode.Enabled = False
                    txtItemCode.Enabled = False
                    cmdhelp.Enabled = False
                    cmdhelp1.Enabled = False
                    TxtItemdrgNO.Enabled = False
                    blnEdit_allowed = True
                    With fpDTLSpread
                        .Col = EnumCustItemGrid.Valid_from : .Col2 = EnumCustItemGrid.Data_type : .Row = 1 : .Row2 = .MaxRows : .BlockMode = True : .Lock = True : .BlockMode = False
                        For intKounter = 1 To .MaxRows
                            varActive = Nothing
                            Call .GetText(EnumCustItemGrid.Active, intKounter, varActive)
                            If varActive Then
                                .Col = EnumCustItemGrid.Active : .Col2 = EnumCustItemGrid.Active : .Row = intKounter : .Row2 = intKounter : .BlockMode = True : .Lock = False : .BlockMode = False
                            End If
                        Next intKounter
                        .Row = 1
                        .Col = EnumCustItemGrid.CustItemCode
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        .Enabled = True
                        .Focus()
                    End With
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                If validate_save() = True Then
                    Select Case ctlCustItem.Mode
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                            With fpDTLSpread
                                For intCounter = 1 To .MaxRows
                                    .Row = intCounter : .Col = EnumCustItemGrid.CustItemCode
                                    Item_Code = Trim(.Text)
                                    .Col = EnumCustItemGrid.CustDrgNo
                                    Item_drgno = Trim(.Text)
                                    .Col = EnumCustItemGrid.CustItemDesc
                                    Description = Trim(.Text)
                                    .Col = EnumCustItemGrid.Qty
                                    dblqty = Val(.Text)
                                    .Col = EnumCustItemGrid.Rate
                                    dblRate = Val(.Text)
                                    .Col = EnumCustItemGrid.Active
                                    Active = CStr(Val(.Value))
                                    If Item_Code <> "" And dblRate > 0 Then
                                        SqlQry = "insert into CustSuppliedItem_Mst (Customer_code,Finish_Item_Code,Item_Code,Cust_Drgno,Item_drgno,Description,Qty,Rate,Active_Flag,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,VALID_FROM,VALID_TO,GRIN_AUTH_DATE,GRIN_NO,UNIT_CODE)" & _
                                                 "values('" & Trim(txtCustCode.Text) & "','" & Trim(txtItemCode.Text) & "','" & Trim(Item_Code) & "','" & Trim(TxtItemdrgNO.Text) & "','" & Trim(Item_drgno) & "','" & Trim(Description) & "'," & dblqty & "," & dblRate & " ," & Active & ",getdate(),'" & mP_User & "',getdate(),'" & mP_User & "', getdate(),'01 JAN 2056 00:00:00',NULL,NULL,'" & gstrUNITID & "')"
                                        mP_Connection.Execute(SqlQry, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    End If
                                Next
                            End With
                            MsgBox("Record Saved successfully", MsgBoxStyle.Information, ResolveResString(100))
                            Call EnableControls(True, Me)
                            ctlCustItem.Revert()
                            fpDTLSpread.MaxRows = 0
                            Me.txtCustCode.Text = ""
                            Me.txtItemCode.Text = ""
                            Me.TxtItemdrgNO.Text = ""
                            Me.lbldescription.Text = ""
                            Me.lblcustcode.Text = ""
                            blnEdit_allowed = False
                            ctlCustItem.Enabled(2) = False
                            ctlCustItem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                            ctlCustItem.Enabled(1) = False
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                            With fpDTLSpread
                                For intCounter = 1 To .MaxRows
                                    .Row = intCounter : .Col = EnumCustItemGrid.CustItemCode
                                    Item_Code = Trim(.Text)
                                    .Col = EnumCustItemGrid.CustDrgNo
                                    Item_drgno = Trim(.Text)
                                    .Col = EnumCustItemGrid.CustItemDesc
                                    Description = Trim(.Text)
                                    .Col = EnumCustItemGrid.Qty
                                    dblqty = Val(.Text)
                                    .Col = EnumCustItemGrid.Rate
                                    dblRate = Val(.Text)
                                    .Col = EnumCustItemGrid.Active
                                    Active = CStr(Val(.Value))
                                    .Col = EnumCustItemGrid.Data_type
                                    strData_type = .Text
                                    .Col = EnumCustItemGrid.Valid_from
                                    Dim strFromDate = .Text
                                    .Col = EnumCustItemGrid.Valid_to
                                    Dim strToDate = .Text
                                    If Item_Code <> "" And dblRate > 0 Then
                                        If (strData_type = "E") And (Active = 0) Then
                                            SqlQry = "UPDATE CustSuppliedItem_Mst SET ACTIVE_FLAG = " & Active & " ,VALID_TO = GETDATE() WHERE CUSTOMER_CODE = '" & txtCustCode.Text & "' AND FINISH_ITEM_CODE = '" & txtItemCode.Text & "' AND ITEM_CODE = '" & Item_Code & "' AND ACTIVE_FLAG = 1 and UNIT_CODE = '" & gstrUNITID & "'  and " & _
                                            "convert(date,valid_from,103)=convert(date,'" & strFromDate & "',103)  and convert(date,valid_to,103)=convert(date,'" & strToDate & "',103)"
                                            mP_Connection.Execute(SqlQry, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        ElseIf strData_type = "N" And (Active = 1) Then
                                            SqlQry = "insert into CustSuppliedItem_Mst (Customer_code,Finish_Item_Code,Item_Code,Cust_Drgno,Item_drgno,Description,Qty,Rate,Active_Flag,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,VALID_FROM,VALID_TO,GRIN_AUTH_DATE,GRIN_NO,UNIT_CODE)" & _
                                                     "values('" & Trim(txtCustCode.Text) & "','" & Trim(txtItemCode.Text) & "','" & Trim(Item_Code) & "','" & Trim(TxtItemdrgNO.Text) & "','" & Trim(Item_drgno) & "','" & Trim(Description) & "'," & dblqty & "," & dblRate & " ," & Active & ",getdate(),'" & mP_User & "',getdate(),'" & mP_User & "', getdate(), '01 JAN 2056 00:00:00',NULL,NULL,'" & gstrUNITID & "')"
                                            mP_Connection.Execute(SqlQry, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        End If
                                    End If
                                Next
                            End With
                            MsgBox("Record Updated successfully", MsgBoxStyle.Information, ResolveResString(100))
                            fpDTLSpread.MaxRows = 0
                            Me.txtCustCode.Text = ""
                            Me.txtItemCode.Text = ""
                            Me.TxtItemdrgNO.Text = ""
                            Me.lbldescription.Text = ""
                            Me.lblcustcode.Text = ""
                            blnEdit_allowed = False
                            Call EnableControls(True, Me)
                            ctlCustItem.Revert()
                            ctlCustItem.Enabled(2) = False
                            ctlCustItem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                            ctlCustItem.Enabled(1) = False
                    End Select
                Else
                    MsgBox(Msgdes, MsgBoxStyle.Information, "Empower")
                    Exit Sub
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL 'for cancel
                If (Me.ctlCustItem.Mode > 0) Or (Me.ctlCustItem.Button > 0) Then
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION, 60095) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        gblnCancelUnload = False
                        gblnFormAddEdit = False
                        ctlCustItem.Focus()
                        fpDTLSpread.MaxRows = 0
                        If Len(txtCustCode.Text) = 0 Then
                            Call RefreshForm(False)
                            txtCustCode.Enabled = True
                            cmdhelp.Enabled = True
                            txtCustCode.Focus()
                        Else
                            If (Me.ctlCustItem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD) Or (Me.ctlCustItem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT) Then
                                Call RefreshForm(False)
                                txtCustCode.Enabled = True
                                cmdhelp.Enabled = True
                                txtCustCode.Focus()
                            Else
                                Call RefreshForm(True)
                            End If
                        End If
                    Else
                        ctlCustItem.Focus()
                        Exit Sub
                    End If
                End If
                blnEdit_allowed = False
                ctlCustItem.Revert()
                Call EnableControls(False, Me)
                Me.txtCustCode.Enabled = True
                cmdhelp.Enabled = True
                ctlCustItem.Enabled(2) = False
                ctlCustItem.Enabled(1) = False
                ctlCustItem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                txtCustCode.Focus()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
        End Select
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function validate_save() As Boolean
        Dim Msgctr As Integer
        Msgctr = 1
        Msgdes = ""
        If Me.txtCustCode.Text = "" Then
            validate_save = False
            Msgdes = Msgdes & Msgctr & ". Customer Code can not be blank " & vbCrLf
            Exit Function
        End If
        If Me.txtItemCode.Text = "" Then
            validate_save = False
            Msgdes = Msgdes & Msgctr & ". Item Code can not be blank " & vbCrLf
            Exit Function
        End If
        validate_save = True
    End Function
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        On Error GoTo errHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm") '("HLPCSTMS0001.htm")
        Exit Sub
errHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTMST0009_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo Err_Handler
        'Checking the form name in the Windows list
        mdifrmMain.CheckFormName = mintFormIndex
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0009_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo Err_Handler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0009_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Err_Handler
        If ctlCustItem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or ctlCustItem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Or ctlCustItem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If KeyCode = System.Windows.Forms.Keys.Escape Then
                If ctlCustItem.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    Call ctlCustItem_ButtonClick(ctlCustItem, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL))
                End If
            End If
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0009_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo Err_Handler
        'Get the index of form in the Windows list
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FitToClient(Me, fraCustItem, ctlFormHeader1, ctlCustItem) 'To fit the form in the MDI
        Call EnableControls(False, Me, True) 'To Disable controls
        Me.txtCustCode.Enabled = True : Me.cmdhelp.Enabled = True : SetGridProperty()
        ctlCustItem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
        ctlCustItem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        ctlCustItem.Enabled(2) = False
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0009_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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
            Cancel = 1
        End If
        Exit Sub
Err_Handler:
        gblnCancelUnload = True
        gblnFormAddEdit = True
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTMST0009_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo Err_Handler
        'REFRESH
        'Removing the form name from list
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        'Setting the corresponding node's tag
        'Closing the recordset
        'Releasing the form reference
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub fpDTLSpread_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles fpDTLSpread.ButtonClicked
        Dim vardata As Object, VarFlag As Object
        If blnEdit_allowed Then
            With fpDTLSpread
                Select Case .ActiveCol
                    Case EnumCustItemGrid.Active
                        .Row = .ActiveRow
                        vardata = Nothing
                        Call .GetText(EnumCustItemGrid.Data_type, .Row, vardata)
                        VarFlag = Nothing
                        Call .GetText(EnumCustItemGrid.Active, .Row, VarFlag)
                        If VarFlag = "0" And vardata = "A" Then
                            Call .SetText(EnumCustItemGrid.Data_type, .Row, "E")
                        End If
                End Select
            End With
        End If
    End Sub
    Private Sub fpDTLSpread_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles fpDTLSpread.KeyDownEvent
        On Error GoTo Errorhandler
        Dim strhelp1 As String
        Dim strhelp() As String
        Dim strItem_condition As String
        Dim kounter As Integer
        With fpDTLSpread
            .Row = .ActiveRow
            .Col = .ActiveCol
            If eventArgs.keyCode = 13 Then
                Select Case .ActiveCol
                    Case EnumCustItemGrid.CustItemCode And Trim(.Text) <> ""
                        Call GetItem_Details()
                    Case EnumCustItemGrid.CustItemDesc And Trim(.Text) <> ""
                        .Row = .ActiveRow
                        .Col = EnumCustItemGrid.CustDrgNo
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    Case EnumCustItemGrid.CustDrgNo And Trim(.Text) <> ""
                        .Row = .ActiveRow
                        .Col = EnumCustItemGrid.Rate
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    Case EnumCustItemGrid.Qty And Val(.Text) > 0
                        .Row = .ActiveRow
                        .Col = EnumCustItemGrid.Rate
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = 2
                        .TypeFloatMin = CDbl("0.00")
                        .TypeFloatMax = CDbl("9999999.99")
                    Case EnumCustItemGrid.Rate And Val(.Text) > 0
                        If Val(.Text) > 0 And .Row = .MaxRows Then
                            .Col = EnumCustItemGrid.Active
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        ElseIf Val(.Text) > 0 Then
                            .Col = EnumCustItemGrid.Active
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        End If
                    Case EnumCustItemGrid.Active
                        If .Row = .MaxRows Then
                            .MaxRows = .MaxRows + 1
                            .Row = .MaxRows
                            .Col = EnumCustItemGrid.Valid_from : .Col2 = EnumCustItemGrid.Valid_to : .Row = 1 : .Row2 = .MaxRows : .BlockMode = True : .Lock = True : .BlockMode = False
                            .Row = .MaxRows
                            .Col = EnumCustItemGrid.CustItemCode
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                        Else
                            .Row = .Row + 1
                            .Col = EnumCustItemGrid.CustItemCode
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                        End If
                End Select
            ElseIf eventArgs.keyCode = System.Windows.Forms.Keys.F1 And .Col = EnumCustItemGrid.CustItemCode Then
                For kounter = 1 To .MaxRows
                    .Row = kounter
                    .Col = EnumCustItemGrid.Active
                    strhelp1 = .Text
                    .Col = EnumCustItemGrid.CustItemCode
                    If Len(.Text) > 0 And strhelp1 = "1" Then
                        strItem_condition = strItem_condition & .Text.Trim & ","
                    End If
                Next kounter
                If Not IsNothing(strItem_condition) Then
                    strItem_condition = "'" & Mid(strItem_condition, 1, (Len(strItem_condition) - 1)) & "'"
                End If
                strhelp1 = "SELECT DISTINCT ITEM_CODE, DESCRIPTION, ISNULL(DRAWINGNO,'') DRAWINGNO" & _
                         " FROM ITEM_MST"
                If Not IsNothing(strItem_condition) Then
                    '10249069
                    strhelp1 = strhelp1 & " WHERE ITEM_CODE NOT IN (SELECT * FROM DBO.UDF_SPLIT_STRING(" & strItem_condition & ",','))" & _
                          " AND ITEM_CODE IN (SELECT DISTINCT ITEM_CODE FROM TMPBOM (NOLOCK) WHERE FINISHEDITEM = '" & Trim(txtItemCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "' and ip_address='" & gstrIpaddressWinSck & "')"
                Else
                    strhelp1 = strhelp1 & " WHERE ITEM_CODE IN (SELECT DISTINCT ITEM_CODE FROM TMPBOM (NOLOCK) WHERE FINISHEDITEM = '" & Trim(txtItemCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "' and ip_address='" & gstrIpaddressWinSck & "')"
                End If
                strhelp1 = strhelp1 & " AND STATUS = 'A' AND CSI_Flag = 1 and UNIT_CODE = '" & gstrUNITID & "'"
                '10249069
                strhelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strhelp1)
                If Not (UBound(strhelp) = -1) Then
                    If (Len(strhelp(0)) >= 1) And strhelp(0) = "0" Then
                        MsgBox("Item Does Not Exist.Please Define BOM For Selected Finished Good", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        .Row = .ActiveRow
                        .Col = EnumCustItemGrid.CustItemCode : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                        eventArgs.keyCode = 0
                        Exit Sub
                    Else
                        .Row = .ActiveRow
                        .Col = EnumCustItemGrid.CustItemCode
                        .Text = strhelp(0)
                        .Col = EnumCustItemGrid.CustItemDesc : .Lock = True : .BlockMode = False : .Text = strhelp(1)
                        .Col = EnumCustItemGrid.CustDrgNo : .Lock = True : .BlockMode = False : .Text = strhelp(2)
                        .Col = EnumCustItemGrid.CustItemCode
                        qry = "Select ISNULL(Required_Qty,0) Required_Qty from bom_mst where Item_Code = '" & Trim(txtItemCode.Text) & "' and RawMaterial_Code = '" & Trim(.Text) & "' and UNIT_CODE = '" & gstrUNITID & "' "
                        RsQty = mP_Connection.Execute(qry)
                        If RsQty.EOF <> True Then
                            .Col = EnumCustItemGrid.Valid_from : .Col2 = EnumCustItemGrid.Valid_to : .Row = 1 : .Row2 = .MaxRows : .BlockMode = True : .Lock = True : .BlockMode = False
                            .Col = EnumCustItemGrid.Qty : .Row = .ActiveRow : .Lock = True : .BlockMode = False : .Text = Str(RsQty.Fields("Required_Qty").Value) : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .BlockMode = True
                            .Col = EnumCustItemGrid.Rate : .BlockMode = False : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2
                            .Row = .ActiveRow
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                        End If
                    End If
                End If
            ElseIf .Col = EnumCustItemGrid.CustItemCode And Trim(.Text) <> "" And (eventArgs.keyCode = 37 Or eventArgs.keyCode = 38 Or eventArgs.keyCode = 39 Or eventArgs.keyCode = 40) Then
                Call GetItem_Details()
            ElseIf eventArgs.keyCode = System.Windows.Forms.Keys.N And eventArgs.shift = VB6.ShiftConstants.CtrlMask Then
                Call AddBlankRow()
            End If
        End With
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub GetItem_Details()
        On Error GoTo Errorhandler
        Dim kounter As Integer
        Dim strItem_condition As String
        Dim strhelp1 As String
        With fpDTLSpread
            .Row = .ActiveRow
            .Col = EnumCustItemGrid.CustItemCode
            'qry = "select distinct  a.Item_Code,b.description,ISNULL(b.Drawingno,'') Drawingno from grn_dtl a (NOLOCK), item_mst b (NOLOCK) where Doc_No in (select doc_no from grn_hdr where Doc_Category = 'Z'and GRN_Cancelled = 0 and vendor_code = '" & Trim(txtCustCode.Text) & "') and a.item_code = b.item_code and a.Item_code = '" & Trim(.Text) & "' and  a.Item_Code in (select RawMaterial_Code from bom_mst (NOLOCK) where Item_Code = '" & Trim(txtItemCode.Text) & "')"
            For kounter = 1 To .MaxRows
                .Row = kounter
                .Col = EnumCustItemGrid.Active
                strhelp1 = .Text
                .Col = EnumCustItemGrid.CustItemCode
                If Len(.Text) > 0 And strhelp1 = "1" Then
                    strItem_condition = strItem_condition & .Text.Trim & ","
                End If
            Next kounter
            If Not IsNothing(strItem_condition) Then
                strItem_condition = "'" & Mid(strItem_condition, 1, (Len(strItem_condition) - 1)) & "'"
            End If
            qry = "SELECT DISTINCT ITEM_CODE, DESCRIPTION, ISNULL(DRAWINGNO,'') DRAWINGNO" & _
                     " FROM ITEM_MST" & _
                     " WHERE ITEM_CODE IN (SELECT DISTINCT RAWMATERIAL_CODE FROM BOM_MST WHERE FINISHED_PRODUCT_CODE = '" & Trim(txtItemCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "')"
            If Len(.Text.Trim) > 0 Then
                qry = qry & "AND ITEM_CODE = '" & .Text.Trim & "'"
            End If
            If Not IsNothing(strItem_condition) Then
                qry = qry & " AND ITEM_CODE NOT IN (SELECT * FROM DBO.UDF_SPLIT_STRING(" & strItem_condition & ",','))"
            End If
            qry = qry & " AND STATUS = 'A' AND CSI_FLAG = 1 and UNIT_CODE = '" & gstrUNITID & "'"
            Rs = mP_Connection.Execute(qry)
            If Rs.EOF <> True Then
                .Row = .ActiveRow
                .Col = EnumCustItemGrid.CustItemDesc : .Lock = True : .BlockMode = False : .Text = Rs.Fields("Description").Value
                .Col = EnumCustItemGrid.CustDrgNo : .Lock = True : .BlockMode = False : .Text = Rs.Fields("Drawingno").Value
                .Col = EnumCustItemGrid.CustItemCode
                qry = "Select ISNULL(Required_Qty,0) Required_Qty from bom_mst where Item_Code = '" & Trim(txtItemCode.Text) & "' and RawMaterial_Code = '" & Trim(.Text) & "' and UNIT_CODE = '" & gstrUNITID & "' "
                RsQty = mP_Connection.Execute(qry)
                If RsQty.EOF <> True Then
                    .Col = EnumCustItemGrid.Qty : .Lock = True : .BlockMode = False : .Text = Str(RsQty.Fields("Required_Qty").Value) : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .BlockMode = True
                    .Col = EnumCustItemGrid.Rate : .BlockMode = False : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                End If
            Else
                Call MsgBox("Item Not Exist In The Grin Against This Customer.Please Check", MsgBoxStyle.Information, "Help")
                .Row = .ActiveRow
                .Col = EnumCustItemGrid.CustItemCode : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                Exit Sub
            End If
        End With
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtItemCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtItemCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Errorhandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            If cmdhelp1.Enabled Then Call cmdHelp1_Click(cmdhelp1, New System.EventArgs())
        End If
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
                        Call txtItemCode_Validating(txtItemCode, New System.ComponentModel.CancelEventArgs(False))
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtItemCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtItemCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo errHandler
        If Len(Trim(txtItemCode.Text)) = 0 Then
            GoTo EventExitSub
        Else
            Select Case ctlCustItem.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    Rs = mP_Connection.Execute("SELECT ITEM_CODE, CUST_DRGNO, DRG_DESC FROM VW_CUST_SUPPLIED_ITEM_MST_FG_HELP WHERE ACCOUNT_CODE = '" & txtCustCode.Text.Trim & "' AND ITEM_CODE = '" & txtItemCode.Text.Trim & "' and UNIT_CODE = '" & gstrUNITID & "'")
                    If Rs.EOF <> True Then
                        Me.lbldescription.Text = Rs.Fields("DRG_DESC").Value
                        Me.TxtItemdrgNO.Text = Rs.Fields("CUST_DRGNO").Value
                        'dblcustmtrl = Rs.Fields("Cust_Mtrl").Value
                        Me.TxtItemdrgNO.Enabled = False
                        Me.txtItemCode.Enabled = False
                        Me.txtCustCode.Enabled = False
                        With fpDTLSpread
                            .Enabled = True
                            .MaxRows = 1
                            .Row = 1
                            .Col = 1
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                        End With
                    Else
                        Call MsgBox("Item Does Not Exist.Please Define Customer Item Linking For This Customer", MsgBoxStyle.Information, "Help")
                        Me.txtItemCode.Text = ""
                        Me.lbldescription.Text = ""
                        Me.TxtItemdrgNO.Text = ""
                        Me.txtItemCode.Focus()
                        'With Me.fpDTLSpread
                        '    .MaxRows = 1
                        '    .Row = 1 : .Row2 = 1 : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
                        'End With
                    End If
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    Call Showdetails(Trim(txtCustCode.Text), Trim(txtItemCode.Text))
            End Select
        End If
        GoTo EventExitSub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub AddBlankRow()
        With fpDTLSpread
            .Enabled = True
            .MaxRows = .MaxRows + 1
            .Row = 1
            .Col = EnumCustItemGrid.Valid_from
            .Col2 = EnumCustItemGrid.Data_type
            .BlockMode = True
            .Lock = True
            .BlockMode = False
            Call .SetText(EnumCustItemGrid.Data_type, .MaxRows, "N")
            .Col = 1
            .Row = .MaxRows
            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
            .Focus()
        End With
    End Sub
    Private Sub txtCustCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustCode.TextChanged
        If Trim(txtCustCode.Text) <> "" Then
            Call txtCustCode_Validating(txtCustCode, New System.ComponentModel.CancelEventArgs(False))
        Else
            fpDTLSpread.MaxRows = 0
            Me.txtCustCode.Text = ""
            Me.txtItemCode.Text = ""
            Me.TxtItemdrgNO.Text = ""
            Me.lbldescription.Text = ""
            Me.lblcustcode.Text = ""
        End If
    End Sub
    Private Sub txtCustCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo Errorhandler
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
                            txtCustCode.Focus()
                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
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
                    Rs = mP_Connection.Execute("select cust_name from customer_mst where Customer_code ='" & Trim(txtCustCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'")
                    If Rs.EOF <> True Then
                        lblcustcode.Text = Rs.Fields("Cust_Name").Value
                        Me.txtItemCode.Enabled = True
                        Me.txtItemCode.Focus() : Me.cmdhelp1.Enabled = True
                    Else
                        Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_EXCLAMATION)
                        txtCustCode.Text = ""
                        lblcustcode.Text = ""
                        txtCustCode.Focus()
                        Cancel = True
                        GoTo EventExitSub
                    End If
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    qry = "select cust_name from customer_mst A,custsupplieditem_mst b where a.Customer_code = b.CUSTOMER_CODE AND a.UNIT_CODE = b.UNIT_CODE and a.CUSTomer_CODE ='" & Trim(txtCustCode.Text) & "' and a.UNIT_CODE = '" & gstrUNITID & "'"
                    Rs = mP_Connection.Execute(qry)
                    If Rs.EOF <> True Then
                        lblcustcode.Text = Rs.Fields("Cust_Name").Value
                        txtItemCode.Enabled = True
                        txtItemCode.Focus()
                        cmdhelp1.Enabled = True
                    Else
                        Cancel = True
                        Call RefreshForm(False)
                        txtCustCode.Enabled = True
                        cmdhelp.Enabled = True
                        txtCustCode.Focus()
                        GoTo EventExitSub
                    End If
            End Select
        End If
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub TxtItemdrgNO_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TxtItemdrgNO.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Errorhandler
        If KeyCode = 13 And TxtItemdrgNO.Text <> "" Then
            If fpDTLSpread.MaxRows = 0 Then
                fpDTLSpread.MaxRows = fpDTLSpread.MaxRows + 1
            End If
            If ctlCustItem.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                fpDTLSpread.Enabled = True
                fpDTLSpread.Focus()
                fpDTLSpread.Col = 1
                fpDTLSpread.Row = 1
                fpDTLSpread.Action = FPSpreadADO.ActionConstants.ActionActiveCell
            End If
        Else
            TxtItemdrgNO.Focus()
        End If
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub TxtItemdrgNO_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtItemdrgNO.Leave
        On Error GoTo Errorhandler
        If fpDTLSpread.MaxRows = 0 Then
            fpDTLSpread.MaxRows = fpDTLSpread.MaxRows + 1
            fpDTLSpread.Col = 1
            fpDTLSpread.Action = FPSpreadADO.ActionConstants.ActionActiveCell
        End If
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function SelectQuery(ByRef pstrFName As String, ByRef pstrTableName As String, Optional ByRef pstrCondition As String = "") As Boolean
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
    Public Function Showdetails(ByVal Strcust_code As String, ByVal StrItemCode As String) As Object
        On Error GoTo Errorhandler
        Dim RSDETAILS As New ClsResultSetDB
        Dim strString As String
        Dim intCounter As Short
        Dim rsResult As ClsResultSetDB
        If Len(Strcust_code) > 0 And Len(StrItemCode) > 0 Then
            rsResult = New ClsResultSetDB
            Call rsResult.GetResult("select description from item_mst where item_code = '" & StrItemCode & "' and UNIT_CODE = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            txtItemCode.Text = Trim(StrItemCode)
            txtItemCode.Enabled = False
            Me.lbldescription.Text = rsResult.GetValue("description")
            rsResult = Nothing
            strString = "Select Item_Code,Item_drgno,Description,Qty,Rate,Cust_Drgno,Active_Flag,valid_from,valid_to from CustSuppliedItem_Mst where Customer_Code = '" & Trim(Strcust_code) & "' and Finish_item_Code = '" & Trim(Me.txtItemCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "' "
            RSDETAILS.GetResult(strString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockPessimistic)
            If RSDETAILS.RowCount > 0 Then
                TxtItemdrgNO.Text = RSDETAILS.GetValue("Cust_Drgno")
                Me.fpDTLSpread.MaxRows = 0
                intCounter = 1
                While RSDETAILS.EOFRecord <> True
                    With Me.fpDTLSpread
                        .MaxRows = .MaxRows + 1
                        .Row = intCounter
                        .Col = EnumCustItemGrid.CustItemCode
                        .Text = RSDETAILS.GetValue("Item_Code")
                        .Col = EnumCustItemGrid.CustItemDesc
                        .Text = RSDETAILS.GetValue("Description")
                        .Col = EnumCustItemGrid.CustDrgNo
                        .Text = RSDETAILS.GetValue("Item_drgno")
                        .Col = EnumCustItemGrid.Qty
                        .Text = RSDETAILS.GetValue("Qty")
                        .Col = EnumCustItemGrid.Rate : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2
                        .Text = RSDETAILS.GetValue("Rate")
                        .Col = EnumCustItemGrid.Active ': .CellType = CellTypeCheckBox
                        If RSDETAILS.GetValue("Active_Flag") = "True" Then
                            .Value = CStr(1)
                        ElseIf RSDETAILS.GetValue("Active_Flag") = "False" Then
                            .Value = CStr(0)
                            .Lock = True
                        End If
                        Call .SetText(EnumCustItemGrid.Valid_from, .Row, VB6.Format(RSDETAILS.GetValue("valid_from"), "DD MMM YYYY HH:MM:SS"))
                        Call .SetText(EnumCustItemGrid.Valid_to, .Row, VB6.Format(RSDETAILS.GetValue("valid_To"), "DD MMM YYYY HH:MM:SS"))
                        Call .SetText(EnumCustItemGrid.Data_type, .Row, "A")
                        intCounter = intCounter + 1
                        RSDETAILS.MoveNext()
                    End With
                End While
                With Me.fpDTLSpread
                    .Col = EnumCustItemGrid.CustItemCode
                    .Col2 = EnumCustItemGrid.Valid_to
                    .Row = 1
                    .Row2 = .MaxRows
                    .BlockMode = True
                    .ForeColor = Color.SlateGray
                    .Lock = True
                    .BlockMode = False
                End With
                ctlCustItem.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
            Else
                Call MsgBox(" Data Not Exist To Display", MsgBoxStyle.Information)
                Me.txtItemCode.Text = ""
            End If
        End If
        Exit Function
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
	Public Sub SetGridProperty()
		On Error GoTo Errorhandler
        If Me.ctlCustItem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
        ElseIf Me.ctlCustItem.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
        Else
            With Me.fpDTLSpread
                .MaxRows = 0 : .MaxCols = 9
                .Row = 0
                .Col = EnumCustItemGrid.CustItemCode : .Text = "Customer Supplied Item Code" : .BlockMode = True : .ColHidden = True : .BlockMode = False
                .Col = EnumCustItemGrid.CustItemDesc : .Text = " Description" : .BlockMode = True : .ColHidden = True : .BlockMode = False
                .Col = EnumCustItemGrid.CustDrgNo : .Text = "Drawing No" : .BlockMode = True : .ColHidden = True : .BlockMode = False
                .Col = EnumCustItemGrid.Qty : .Text = "Quantity" : .set_ColWidth(EnumCustItemGrid.Qty, 7) : .BlockMode = True : .ColHidden = True : .BlockMode = False
                .Col = EnumCustItemGrid.Rate : .Text = "Rate" : .set_ColWidth(EnumCustItemGrid.Rate, 6) : .BlockMode = True : .ColHidden = True : .BlockMode = False
                .Col = EnumCustItemGrid.Active : .Text = "Active Flag" : .set_ColWidth(EnumCustItemGrid.Active, 9) : .BlockMode = True : .ColHidden = True : .BlockMode = False
                .Col = EnumCustItemGrid.Valid_from : .Text = "Valid From" : .set_ColWidth(EnumCustItemGrid.Valid_from, 14) : .BlockMode = True : .ColHidden = True : .BlockMode = False
                .Col = EnumCustItemGrid.Valid_to : .Text = "Valid To" : .set_ColWidth(EnumCustItemGrid.Valid_to, 14) : .BlockMode = True : .ColHidden = True : .BlockMode = False
                .Col = EnumCustItemGrid.Data_type : .Text = "Data Type" : .ColHidden = True
                .Enabled = False
            End With
        End If
		Exit Sub
Errorhandler: 
		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection)
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
			lblDescription.Text = ""
			lblCustCode.Text = ""
			gblnCancelUnload = False
			gblnFormAddEdit = False
			With Me.fpDTLSpread
				.maxRows = 1
				.Row = 1 : .Row2 = 1 : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
            End With
        Else
            Exit Sub
		End If
		Exit Sub
Err_Handler: 
		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection)
    End Sub
End Class