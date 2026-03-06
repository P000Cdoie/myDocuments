'********************************************************************************
'COPYRIGHT (C)  : MIND  
'FORM NAME      : AGREEMENT MASTER
'CREATED BY     : VINOD SINGH KEMWAL
'CREATED DATE   : 11/02/2010
'********************************************************************************
'MODIFIED BY     : SHABBIR HUSSAIN
'MODIFIED DATE   : CUST PART NO GETS TRUNCATED IN THE GRID 
'                  DUE TO MAX LEN SET TO 15
'********************************************************************************
' Revised By                 -   Roshan Singh
' Revision Date              -   15 JULY 2011
' Description                -   FOR MULTIUNIT FUNCTIONALITY
'-----------------------------------------------------------------------
'MODIFIED BY     : PRASHANT RAJPAL
'MODIFIED DATE   : MAKE CHANGES FOR KEEPING AGREEMENT HISTORY
'                  ISSUE ID: 10116365
'MODIFIED BY AVANISH PATHAK ON 08 NOV 2011 FOR MULTIUNIT CHANGE MANAGEMENT
Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class frmMKTMST0019
    Private Enum EnmGridCol
        PO_No = 1
        CustPartNo
        HelpCustPart
        InternalPartNo
        ProcessType
        BasicValue
        Usage
        TaxDtl
    End Enum
    Private Enum EnmTaxDtl
        TaxId = 1
        TaxHelp
        TaxValue
    End Enum
    Private Enum EnmLbrCost
        Cost_Desc = 1
        Cost_Value
    End Enum
    Dim mintFormIndex As Integer
    Dim mConnString As String
    Private Sub frmMKTMST0019_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
    End Sub

    Private Sub frmMKTMST0019_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Tag) = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
    End Sub

    Private Sub frmMKTMST0019_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            Me.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub frmMKTMST0019_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        Dim keyascii As Short = Asc(e.KeyChar)
        Try
            If keyascii = 39 Then
                keyascii = 0
            End If
            e.KeyChar = Chr(keyascii)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Private Sub frmMKTMST0019_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Call FitToClient(Me, GrpMain, ctlHeader, (Me.cmdGrp), 500)
            Me.MdiParent = mdifrmMain
            mintFormIndex = mdifrmMain.AddFormNameToWindowList(Me.ctlHeader.Tag)
            SetGridHeading()
            Me.chkActive.Checked = True
            dtAgreementDate.Format = DateTimePickerFormat.Custom
            dtAgreementDate.CustomFormat = gstrDateFormat
            dtAgreementDate.Value = GetServerDate()
            Me.cmdGrp.ShowButtons(True, True, False, False)
            cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub cmdGrp_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles cmdGrp.ButtonClick
        Try

            Select Case e.Button
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                    RefreshForm()
                    Me.TxtAgreementNo.Text = ""
                    Me.TxtAgreementNo.Enabled = False
                    Me.CmdAgreementNoHelp.Enabled = False
                    Me.sprItems.MaxRows = 0
                    AddBlankRow()
                    Me.TxtCustCode.Focus()
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                    If Me.GrpTax.Visible Then
                        MsgBox("Please confirm tax details first.", MsgBoxStyle.Information, ResolveResString(100))
                        Exit Sub
                    End If
                    If Me.GrpLabourCost.Visible Then
                        MsgBox("Please confirm Labour Cost Details first.", MsgBoxStyle.Information, ResolveResString(100))
                        Exit Sub
                    End If
                    If Me.cmdGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                        If ValidateForNewRecord() = False Then Exit Sub
                        If ValidateChildParts() = False Then Exit Sub

                        ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
                        If ValidateData() = False Then
                            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
                            Exit Sub
                        End If
                        If SaveRecord() = False Then
                            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
                            Exit Sub
                        End If
                        'RefreshForm()
                        cmdGrp.Revert()
                    ElseIf Me.cmdGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                        ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
                        If ValidateData() = False Then
                            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
                            Exit Sub
                        End If
                        If UpdateRecord() = False Then
                            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
                            Exit Sub
                        End If
                        RefreshForm()
                        cmdGrp.Revert()
                    End If

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                    PopulateRecord(Val(TxtAgreementNo.Text), "EDIT")
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                    Me.Close()
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                    cmdGrp.Revert()
                    RefreshForm()
            End Select

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Private Sub AddTaxDtlRow()
        With Me.SprTaxDtl
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .set_RowHeight(.Row, 12)
            .Col = EnmTaxDtl.TaxId : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = EnmTaxDtl.TaxHelp : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = "..."
            .Col = EnmTaxDtl.TaxValue : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = EnmTaxDtl.TaxId : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
        End With
    End Sub
    Private Sub AddLabourCostRow()
        With Me.SprLbrCost
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .set_RowHeight(.Row, 12)
            .Col = EnmLbrCost.Cost_Desc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .Col = EnmLbrCost.Cost_Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMin = 0
        End With
    End Sub
    Private Sub AddBlankRow()
        ''PURPOSE       : ADD A NEW BLANK ROW IN THE GRID
        On Error GoTo ErrHandler
        With Me.sprItems
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .BackColorStyle = FPSpreadADO.BackColorStyleConstants.BackColorStyleUnderGrid
            .set_RowHeight(.Row, 15)
            .Col = EnmGridCol.PO_No : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .Col = EnmGridCol.CustPartNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True
            .Col = EnmGridCol.HelpCustPart : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = "..."
            .Col = EnmGridCol.InternalPartNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True
            .Col = EnmGridCol.ProcessType : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True
            .Col = EnmGridCol.BasicValue : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Col = EnmGridCol.Usage : .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
            .Col = EnmGridCol.TaxDtl : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = "View"

            .Row = .MaxRows : .Col = EnmGridCol.PO_No
            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(CInt(Err.Description), Err.Source, Err.Description, mP_Connection)

    End Sub
    Private Sub SetGridHeading()
        On Error GoTo ErrHandler
        With Me.sprItems
            .MaxRows = 0
            .MaxCols = EnmGridCol.TaxDtl
            .Row = 0
            .set_RowHeight(0, 20)
            .Col = 0 : .set_ColWidth(0, 3)
            .Col = EnmGridCol.PO_No : .Text = "PO No." : .set_ColWidth(EnmGridCol.PO_No, 14)
            .Col = EnmGridCol.CustPartNo : .Text = "Customer Part No." : .set_ColWidth(EnmGridCol.CustPartNo, 22)
            .Col = EnmGridCol.HelpCustPart : .Text = " " : .set_ColWidth(EnmGridCol.HelpCustPart, 3)
            .Col = EnmGridCol.InternalPartNo : .Text = "Internal Part No." : .set_ColWidth(EnmGridCol.InternalPartNo, 22)
            .Col = EnmGridCol.ProcessType : .Text = "Process Type" : .set_ColWidth(EnmGridCol.ProcessType, 8)
            .Col = EnmGridCol.BasicValue : .Text = "Basic Value" : .set_ColWidth(EnmGridCol.BasicValue, 14)
            .Col = EnmGridCol.Usage : .Text = "Usage" : .set_ColWidth(EnmGridCol.Usage, 8)
            .Col = EnmGridCol.TaxDtl : .Text = "Tax Details" : .set_ColWidth(EnmGridCol.TaxDtl, 10) : .ColHidden = True

        End With
        With Me.SprTaxDtl
            .MaxRows = 0
            .MaxCols = EnmTaxDtl.TaxValue
            .Row = 0
            .set_RowHeight(0, 15)
            .Col = EnmTaxDtl.TaxId : .Text = "Tax Type" : .set_ColWidth(EnmTaxDtl.TaxId, 10)
            .Col = EnmTaxDtl.TaxHelp : .Text = " " : .set_ColWidth(EnmTaxDtl.TaxHelp, 3)
            .Col = EnmTaxDtl.TaxValue : .Text = "Value (%)" : .set_ColWidth(EnmTaxDtl.TaxValue, 8)
            GrpTax.Visible = False
        End With
        With Me.SprLbrCost
            .MaxRows = 0
            .MaxCols = EnmLbrCost.Cost_Value
            .Row = 0
            .set_RowHeight(0, 15)
            .Col = EnmLbrCost.Cost_Desc : .Text = "Labour Cost" : .set_ColWidth(EnmLbrCost.Cost_Desc, 12)
            .Col = EnmLbrCost.Cost_Value : .Text = "Value" : .set_ColWidth(EnmLbrCost.Cost_Value, 10)
            GrpLabourCost.Visible = False
        End With

        Exit Sub
ErrHandler:
        gobjError.RaiseError(CInt(Err.Description), Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub sprItems_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles sprItems.ButtonClicked
        Dim strCustPart() As String
        Dim strQry As String
        Dim intRow As Integer
        If cmdGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            If e.col = EnmGridCol.HelpCustPart Then

                Dim strItem As String = ""
                If Me.TxtCustCode.Text.Trim = "" Then
                    MsgBox("Please First Select Customer", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If
                With Me.sprItems
                    For intRow = 1 To .MaxRows
                        .Row = intRow : .Col = EnmGridCol.CustPartNo
                        If .Text.Trim <> "" Then
                            strItem = strItem & ",'" & .Text.Trim & "'"
                        End If
                    Next
                    If strItem.Trim.Length > 0 Then
                        strItem = "(" & strItem.Substring(1) & ")"
                    End If

                End With

                strQry = " SELECT A.CUST_DRGNO,DRG_DESC,A.ITEM_CODE,B.DESCRIPTION,B.TYPE_CODE "
                strQry += " FROM CUSTITEM_MST A INNER JOIN ITEM_MST B"
                strQry += " ON A.ITEM_CODE=B.ITEM_CODE and A.unit_code = B.unit_code"
                strQry += " WHERE B.ITEM_MAIN_GRP <> 'F' AND A.ACTIVE=1"
                strQry += " AND A.ACCOUNT_CODE='" & Me.TxtCustCode.Text.Trim & "' and A.unit_code= '" & gstrUNITID & "'"
                If strItem.Trim.Length > 0 Then
                    strQry += " AND A.CUST_DRGNO NOT IN " & strItem & ""
                End If
                strQry += " GROUP BY A.CUST_DRGNO,DRG_DESC,A.ITEM_CODE, B.DESCRIPTION, B.TYPE_CODE"
                Application.DoEvents()
                strCustPart = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
                If Not (UBound(strCustPart) = -1) Then
                    If (Len(strCustPart(0)) >= 1) And strCustPart(0) = "0" Then
                        MsgBox("No Customer Assy Part Found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                        Exit Sub
                    Else
                        With sprItems
                            .Row = e.row
                            .Col = EnmGridCol.CustPartNo : .Text = strCustPart(0)
                            .Col = EnmGridCol.InternalPartNo : .Text = strCustPart(2)
                            .Col = EnmGridCol.ProcessType : .Text = strCustPart(4)
                        End With
                    End If
                End If

            End If
        End If
    End Sub
    Private Sub sprItems_KeyDownEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles sprItems.KeyDownEvent
        Try
            If e.keyCode = Keys.Delete Then
                With Me.sprItems
                    .Row = .ActiveRow
                    .Col = .ActiveCol
                    If .Text.Trim = "" Or Val(.Text.Trim) >= 0 Then
                        If MsgBox("Do you want to delete this row?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                            .Action = FPSpreadADO.ActionConstants.ActionDeleteRow
                            .MaxRows = .MaxRows - 1
                            If .MaxRows = 0 Then
                                AddBlankRow()
                            End If
                        End If
                    End If
                End With
            End If
            If e.keyCode = Keys.Enter Then
                With Me.sprItems
                    If .ActiveCol = EnmGridCol.Usage Then
                        AddBlankRow()
                        .Row = .MaxRows : .Col = EnmGridCol.PO_No - 1 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    End If
                End With
            End If
            If e.keyCode = Keys.F1 Then
                With sprItems
                    If .ActiveCol = EnmGridCol.CustPartNo Then
                        sprItems_ButtonClicked(Me.sprItems, New AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent(EnmGridCol.HelpCustPart, .ActiveRow, 1))
                    End If
                End With
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Private Sub cmdCustHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCustHelp.Click
        Dim strQry As String
        Dim strCust() As String
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            If cmdGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                strQry = "SELECT A.CUSTOMER_CODE, A.CUST_NAME AS 'CUSTOMER_NAME'"
                strQry += " FROM CUSTOMER_MST A INNER JOIN CUSTITEM_MST B"
                strQry += " ON A.CUSTOMER_CODE=B.ACCOUNT_CODE and A.unit_code = B.unit_code where A.unit_code = '" & gstrUNITID & "'"
                strQry += " GROUP BY A.CUSTOMER_CODE,A.CUST_NAME"
            Else
                strQry = "SELECT B.CUSTOMER_CODE,B.CUST_NAME"
                strQry += " FROM AGREEMENT_HDR A INNER JOIN CUSTOMER_MST B"
                strQry += " ON A.CUSTOMER_CODE=B.CUSTOMER_CODE and A.unit_code = B.unit_code "
                strQry += " WHERE A.ACTIVE=1 and  A.unit_code = '" & gstrUNITID & "'"
                strQry += " GROUP BY B.CUSTOMER_CODE,B.CUST_NAME"
            End If
            If strQry.Trim.Length = 0 Then
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                Exit Sub
            End If
            strCust = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
            If Not (UBound(strCust) = -1) Then
                If (Len(strCust(0)) >= 1) And strCust(0) = "0" Then
                    MsgBox("No Customer Found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                    Exit Sub
                Else
                    Me.TxtCustCode.Text = strCust(0)
                    Me.lblCustName.Text = strCust(1)
                End If
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)

        Catch EX As Exception
            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Private Sub cmdInternalPartHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInternalPartHelp.Click
        Dim strQry As String
        Dim strItem() As String
        If Me.TxtCustCode.Text.Trim = "" Then
            MsgBox("Please First Select Customer", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If
        Try

            If cmdGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                strQry = " SELECT B.ITEM_CODE,B.DESCRIPTION "
                strQry += " FROM CUSTITEM_MST A INNER JOIN ITEM_MST B"
                strQry += " ON A.ITEM_CODE=B.ITEM_CODE and A.unit_code = B.unit_code "
                strQry += " WHERE B.ITEM_MAIN_GRP='F' AND A.ACCOUNT_CODE='" & Me.TxtCustCode.Text.Trim & "' and A.UNIT_CODE = '" & gstrUNITID & "'"
                strQry += " GROUP BY B.ITEM_CODE,B.DESCRIPTION"
            Else
                strQry = " SELECT INTERNALASSYPART,DESCRIPTION "
                strQry += " FROM AGREEMENT_HDR A INNER JOIN ITEM_MST B "
                strQry += " ON A.INTERNALASSYPART=B.ITEM_CODE AND A.UNIT_CODE = B.UNIT_CODE "
                strQry += " WHERE(A.ACTIVE = 1) AND A.CUSTOMER_CODE='" & Me.TxtCustCode.Text.Trim & "' and A.UNIT_CODE = '" & gstrUNITID & "'"
                strQry += " GROUP BY INTERNALASSYPART,DESCRIPTION "
            End If
            If strQry.Trim.Length = 0 Then Exit Sub
            strItem = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
            If Not (UBound(strItem) = -1) Then
                If (Len(strItem(0)) >= 1) And strItem(0) = "0" Then
                    MsgBox("No Internal Assy Part Found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                    Exit Sub
                Else
                    Me.TxtInternalAssyPart.Text = strItem(0)
                    Me.lblInternalAssyDesc.Text = strItem(1)
                End If
            End If

        Catch EX As Exception
            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Private Sub cmdCustPartHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCustPartHelp.Click
        Dim strQry As String
        Dim strCustPart() As String
        If Me.TxtCustCode.Text.Trim = "" Then
            MsgBox("Please First Select Customer", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        ElseIf Me.TxtInternalAssyPart.Text.Trim = "" Then
            MsgBox("Please First Select Internal Assy. Part", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If
        Try
            If cmdGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                strQry = " SELECT A.CUST_DRGNO,DRG_DESC"
                strQry += " FROM CUSTITEM_MST A INNER JOIN ITEM_MST B"
                strQry += " ON A.ITEM_CODE=B.ITEM_CODE and A.unit_code = B.unit_code "
                strQry += " WHERE B.ITEM_MAIN_GRP='F' AND A.ACTIVE=1 and A.unit_code = '" & gstrUNITID & "' "
                strQry += " AND A.ACCOUNT_CODE='" & Me.TxtCustCode.Text.Trim & "' AND A.ITEM_CODE='" & Me.TxtInternalAssyPart.Text.Trim & "'"
            Else
                strQry = "SELECT CUSTASSYPART,DRG_DESC"
                strQry += " FROM AGREEMENT_HDR A INNER JOIN CUSTITEM_MST B"
                strQry += " ON A.CUSTASSYPART=B.CUST_DRGNO and A.unit_code = B.unit_code "
                strQry += " WHERE(A.ACTIVE = 1 And B.ACTIVE = 1) and A.unit_code = '" & gstrUNITID & "'  "
                strQry += " AND B.ACCOUNT_CODE='" & Me.TxtCustCode.Text.Trim & "' AND B.ITEM_CODE='" & Me.TxtInternalAssyPart.Text.Trim & "'"
                strQry += " GROUP BY CUSTASSYPART,DRG_DESC"
            End If
            If strQry.Trim.Length = 0 Then Exit Sub
            strCustPart = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
            If Not (UBound(strCustPart) = -1) Then
                If (Len(strCustPart(0)) >= 1) And strCustPart(0) = "0" Then
                    MsgBox("No Internal Assy Part Found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                    Exit Sub
                Else
                    Me.TxtCustAssyPart.Text = strCustPart(0)
                    Me.lblCustAssyDesc.Text = strCustPart(1)
                End If
            End If

        Catch EX As Exception
            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Private Sub cmdTaxDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTaxDetails.Click

        With Me.GrpTax
            .Visible = True
            .Top = cmdTaxDetails.Top
            .Left = cmdTaxDetails.Left - (.Width - cmdTaxDetails.Width)
        End With
        If cmdGrp.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If Me.SprTaxDtl.MaxRows = 0 Then
                AddTaxDtlRow()
            End If
        End If
    End Sub

    Private Sub cmdTaxOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTaxOk.Click

        GrpTax.Visible = False
    End Sub

    Private Sub SprTaxDtl_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles SprTaxDtl.ButtonClicked
        Dim strQry As String
        Dim strTax() As String
        Dim strAddedTax As String = ""
        Dim intRow As Integer
        Try
            If cmdGrp.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                With Me.SprTaxDtl
                    For intRow = 1 To .MaxRows
                        .Row = intRow : .Col = EnmTaxDtl.TaxId
                        If .Text.Trim <> "" Then
                            strAddedTax = strAddedTax & "'" & .Text.Trim & "',"
                        End If
                    Next
                    If strAddedTax.Trim.Length > 0 Then
                        strAddedTax = "(" & strAddedTax.Trim.Substring(0, strAddedTax.Trim.Trim.Length - 1) & ")"
                    End If
                End With
                strQry = "SELECT TX_TAXEID,TXRT_RATE_NO FROM GEN_TAXRATE WHERE unit_code='" & gstrUNITID & "' "
                If strAddedTax.Trim.Length > 0 Then
                    strQry = strQry & "  and TX_TAXEID NOT IN " & strAddedTax
                End If
                strQry = strQry & " ORDER BY TX_TAXEID"
                strTax = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
                If Not (UBound(strTax) = -1) Then
                    If (Len(strTax(0)) >= 1) And strTax(0) = "0" Then
                        MsgBox("No Tax defined.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                        Exit Sub
                    Else
                        With Me.SprTaxDtl
                            .Row = e.row
                            .Col = EnmTaxDtl.TaxId
                            .Text = strTax(0)
                            .Col = EnmTaxDtl.TaxValue
                            .Text = strTax(1)
                        End With
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub
    Private Sub SprTaxDtl_KeyDownEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles SprTaxDtl.KeyDownEvent
        If cmdGrp.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If e.keyCode = Keys.Enter Then
                With Me.SprTaxDtl
                    .Row = .ActiveRow
                    .Col = EnmTaxDtl.TaxId
                    If .Text.Trim <> "" Then
                        AddTaxDtlRow()
                    End If
                End With
            ElseIf e.keyCode = Keys.Delete Then
                With Me.SprTaxDtl
                    If MsgBox("Do you want to delete this row", MsgBoxStyle.Question + MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                        .Row = .ActiveRow
                        .Action = FPSpreadADO.ActionConstants.ActionDeleteRow
                        .MaxRows = .MaxRows - 1
                    End If
                End With
            End If
        End If
    End Sub
    Private Function SaveRecord() As Boolean
        SaveRecord = False
        Dim strQry As String
        Dim sqlConn As New SqlConnection
        Dim sqlCmd As SqlCommand
        Dim sqlTrans As SqlTransaction
        Dim sqlRdr As SqlDataReader
        Dim intRow As Integer

        Dim strPOno As String
        Dim strIntrPart As String
        Dim strCustPart As String
        Dim strProcessType As String
        Dim dblBasicValue As Double
        Dim intUsage As Integer
        Dim AgreementNo As Integer
        Dim strTaxId As String
        Dim strTaxValue As String
        Dim strCostDesc As String
        Dim dblCostValue As Double
        Dim blnTran As Boolean
        Try
            sqlConn = SqlConnectionclass.GetConnection
            sqlTrans = sqlConn.BeginTransaction
            blnTran = True

            sqlCmd = New SqlCommand()
            sqlCmd.Connection = sqlConn
            sqlCmd.Transaction = sqlTrans

            strQry = "SELECT ISNULL(MAX(DOC_NO),0)+1 FROM AGREEMENT_HDR where unit_code = '" & gstrUNITID & "'"
            sqlCmd.CommandType = CommandType.Text
            sqlCmd.CommandText = strQry

            sqlRdr = sqlCmd.ExecuteReader
            While sqlRdr.Read
                AgreementNo = sqlRdr.GetInt32(0)
            End While

            sqlRdr.Close()

            strQry = "INSERT INTO AGREEMENT_HDR(DOC_NO,CUSTOMER_CODE,INTERNALASSYPART,CUSTASSYPART,LABOURCOST,TOTALBASICVALUE,ACTIVE,ENT_DT,ENT_USERID,UPD_DT,UPD_USERID,unit_code)"
            strQry += " VALUES(" & AgreementNo & ",'" & Me.TxtCustCode.Text & "','" & Me.TxtInternalAssyPart.Text & "','" & Me.TxtCustAssyPart.Text & "'," & Val(Me.TxtLabourCost.Text) & "," & Val(Me.lblTotalBasicValue.Text) & " ," & IIf(chkActive.Checked, 1, 0) & ",GETDATE(),'" & mP_User & "',GETDATE(),'" & mP_User & "' ,'" & gstrUNITID & "') "
            sqlCmd.CommandText = strQry
            sqlCmd.ExecuteNonQuery()

            With Me.sprItems
                For intRow = 1 To .MaxRows
                    .Row = intRow
                    .Col = EnmGridCol.PO_No : strPOno = .Text
                    .Col = EnmGridCol.InternalPartNo : strIntrPart = .Text.Trim
                    .Col = EnmGridCol.CustPartNo : strCustPart = .Text.Trim
                    .Col = EnmGridCol.ProcessType : strProcessType = .Text.Trim()
                    .Col = EnmGridCol.BasicValue : dblBasicValue = Val(.Value)
                    .Col = EnmGridCol.Usage : intUsage = Val(.Value)
                    strQry = "INSERT INTO AGREEMENT_DTL(DOC_NO,PO_NO,CUSTPARTNO,INTERNALPARTNO,PROCESSTYPE,BASIC_VALUE,USAGE,unit_code)"
                    strQry += " VALUES('" & AgreementNo & "','" & strPOno & "','" & strCustPart & "','" & strIntrPart & "','" & strProcessType & "'," & dblBasicValue & "," & intUsage & ",'" & gstrUNITID & "')"
                    sqlCmd.CommandText = strQry
                    sqlCmd.ExecuteNonQuery()
                Next
            End With
            With Me.SprTaxDtl
                For intRow = 1 To .MaxRows
                    .Row = intRow
                    .Col = EnmTaxDtl.TaxId : strTaxId = .Text
                    .Col = EnmTaxDtl.TaxValue : strTaxValue = .Text
                    If strTaxId = "" Or strTaxValue = "" Then Continue For
                    strQry = "INSERT INTO AGREEMENTTAXDTL (DOC_NO,TAXID,TAXVALUE,unit_code) "
                    strQry += " VALUES ('" & AgreementNo & "','" & strTaxId & "','" & strTaxValue & "','" & gstrUNITID & "')"
                    sqlCmd.CommandText = strQry
                    sqlCmd.ExecuteNonQuery()
                Next
            End With
            With Me.SprLbrCost
                For intRow = 1 To .MaxRows
                    .Row = intRow
                    .Col = EnmLbrCost.Cost_Desc : strCostDesc = .Text.Trim
                    .Col = EnmLbrCost.Cost_Value : dblCostValue = Val(.Value)
                    If strCostDesc = "" Or dblCostValue = 0 Then Continue For
                    strQry = "INSERT INTO AGREEMENTLABOURCOSTDTL (DOC_NO,COST_DESC,COST_VALUE,unit_code) "
                    strQry += " VALUES ('" & AgreementNo & "','" & strCostDesc & "'," & dblCostValue & ",'" & gstrUNITID & "')"
                    sqlCmd.CommandText = strQry
                    sqlCmd.ExecuteNonQuery()
                Next
            End With
            sqlTrans.Commit()
            blnTran = False
            sqlConn.Close()
            SaveRecord = True
            Me.TxtAgreementNo.Text = AgreementNo
            MsgBox("Record successfully saved with new Agreement No [" & AgreementNo & "].", MsgBoxStyle.Information, ResolveResString(100))
        Catch ex As Exception
            If blnTran Then sqlTrans.Rollback()
            If sqlConn.State = ConnectionState.Open Then sqlConn.Close()
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Function
    Private Function UpdateRecord() As Boolean
        Dim strQry As String
        Dim sqlConn As New SqlConnection
        Dim sqlCmd As SqlCommand
        Dim sqlTrans As SqlTransaction = Nothing
        Dim intRow As Integer
        Dim strTaxId As String
        Dim strTaxValue As String
        Dim blnTran As Boolean
        Dim strPOno As String
        Dim strIntrPart As String
        Dim strCustPart As String
        Dim strProcessType As String
        Dim dblBasicValue As Double
        Dim intUsage As Integer
        Dim dblCostValue As Double
        UpdateRecord = False
        Try
            sqlConn = SqlConnectionclass.GetConnection
            sqlTrans = sqlConn.BeginTransaction
            blnTran = True

            sqlCmd = New SqlCommand()
            sqlCmd.Connection = sqlConn
            sqlCmd.Transaction = sqlTrans

            With sqlCmd
                .Parameters.Clear()
                .CommandText = "usp_LogAgreement_History"
                .CommandType = CommandType.StoredProcedure
                .Parameters.Add("@AGREEMENTNO", SqlDbType.Int).Value = Val(TxtAgreementNo.Text)
                .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar).Value = gstrUNITID
                .ExecuteNonQuery()
            End With

            strQry = "DELETE FROM AGREEMENT_DTL WHERE DOC_NO=" & Me.TxtAgreementNo.Text & " and UNIT_CODE='" & gstrUNITID & "' "
            sqlCmd.CommandType = CommandType.Text
            sqlCmd.CommandText = strQry
            sqlCmd.ExecuteNonQuery()

            With Me.sprItems
                For intRow = 1 To .MaxRows
                    .Row = intRow
                    .Col = EnmGridCol.PO_No : strPOno = .Text
                    .Col = EnmGridCol.InternalPartNo : strIntrPart = .Text.Trim
                    .Col = EnmGridCol.CustPartNo : strCustPart = .Text.Trim
                    .Col = EnmGridCol.ProcessType : strProcessType = .Text.Trim()
                    .Col = EnmGridCol.BasicValue : dblBasicValue = Val(.Value)
                    .Col = EnmGridCol.Usage : intUsage = Val(.Value)
                    strQry = "INSERT INTO AGREEMENT_DTL(DOC_NO,PO_NO,CUSTPARTNO,INTERNALPARTNO,PROCESSTYPE,BASIC_VALUE,USAGE,UNIT_CODE)"
                    strQry += " VALUES('" & Me.TxtAgreementNo.Text.Trim & "','" & strPOno & "','" & strCustPart & "','" & strIntrPart & "','" & strProcessType & "'," & dblBasicValue & "," & intUsage & ",'" & gstrUNITID & "')"
                    sqlCmd.CommandType = CommandType.Text
                    sqlCmd.CommandText = strQry
                    sqlCmd.ExecuteNonQuery()
                    dblCostValue += dblBasicValue
                Next
            End With
            lblTotalBasicValue.Text = dblCostValue
            strQry = "UPDATE AGREEMENT_HDR SET TOTALBASICVALUE=" & dblCostValue & " , ACTIVE=" & IIf(chkActive.Checked, 1, 0) & ",UPD_DT=GETDATE(), UPD_USERID='" & mP_User & "' WHERE DOC_NO=" & Me.TxtAgreementNo.Text & " and UNIT_CODE='" & gstrUNITID & "' "
            sqlCmd.CommandType = CommandType.Text
            sqlCmd.CommandText = strQry
            sqlCmd.ExecuteNonQuery()
            strQry = "DELETE FROM AGREEMENTTAXDTL WHERE DOC_NO=" & Me.TxtAgreementNo.Text & " and UNIT_CODE='" & gstrUNITID & "' "
            sqlCmd.CommandType = CommandType.Text
            sqlCmd.CommandText = strQry
            sqlCmd.ExecuteNonQuery()
            With Me.SprTaxDtl
                For intRow = 1 To .MaxRows
                    .Row = intRow
                    .Col = EnmTaxDtl.TaxId : strTaxId = .Text
                    .Col = EnmTaxDtl.TaxValue : strTaxValue = .Text
                    strQry = "INSERT INTO AGREEMENTTAXDTL (DOC_NO,TAXID,TAXVALUE,unit_code) "
                    strQry += " VALUES ('" & Me.TxtAgreementNo.Text & "','" & strTaxId & "','" & strTaxValue & "','" & gstrUNITID & "')"
                    sqlCmd.CommandType = CommandType.Text
                    sqlCmd.CommandText = strQry
                    sqlCmd.ExecuteNonQuery()
                Next
            End With
            sqlTrans.Commit()
            blnTran = False
            sqlConn.Close()
            UpdateRecord = True
            MsgBox("Record successfully Updated.", MsgBoxStyle.Information, ResolveResString(100))
        Catch ex As Exception
            If blnTran Then sqlTrans.Rollback()
            If sqlConn.State = ConnectionState.Open Then sqlConn.Close()
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try

    End Function
    Private Function ValidateData() As Boolean
        ValidateData = False
        Dim intRow As Integer
        Try
            If Me.TxtCustCode.Text = "" Then
                MsgBox("Please Select Customer.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            ElseIf Me.TxtInternalAssyPart.Text = "" Then
                MsgBox("Please Select Internal Assy. Part", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            ElseIf Me.TxtCustAssyPart.Text.Trim = "" Then
                MsgBox("Please Select Customer Assy. Part", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            End If
            With Me.sprItems
                For intRow = 1 To .MaxRows
                    .Row = intRow
                    .Col = EnmGridCol.PO_No
                    If .Text.Trim = "" Then
                        MsgBox("Please enter PO No at row [" & intRow & " ]")
                        Exit Function
                    End If
                    .Col = EnmGridCol.InternalPartNo
                    If .Text.Trim = "" Then
                        MsgBox("Please enter Internal Part No at row [" & intRow & " ]")
                        Exit Function
                    End If
                    .Col = EnmGridCol.CustPartNo
                    If .Text.Trim = "" Then
                        MsgBox("Please enter Customer Part NO at row [" & intRow & " ]")
                        Exit Function
                    End If
                    .Col = EnmGridCol.BasicValue
                    If .Text.Trim = "" Then
                        MsgBox("Please enter Basic Value at row [" & intRow & " ]")
                        Exit Function
                    End If
                    .Col = EnmGridCol.Usage
                    If .Text.Trim = "" Then
                        MsgBox("Please enter Usage at row [" & intRow & " ]")
                        Exit Function
                    End If
                Next
                ValidateData = True
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Function

    Private Sub RefreshForm()
        Me.SprTaxDtl.MaxRows = 0
        Me.sprItems.MaxRows = 0
        Me.SprLbrCost.MaxRows = 0
        Me.TxtCustCode.Text = ""
        Me.TxtInternalAssyPart.Text = ""
        Me.TxtCustAssyPart.Text = ""
        Me.lblCustAssyDesc.Text = ""
        Me.lblCustName.Text = ""
        Me.lblInternalAssyDesc.Text = ""
        dtAgreementDate.Value = GetServerDate()
        Me.TxtAgreementNo.Text = ""
        Me.TxtAgreementNo.Enabled = True
        Me.TxtLabourCost.Text = ""
        Me.CmdAgreementNoHelp.Enabled = True
        Me.lblTotalBasicValue.Text = ""
        cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
    End Sub

    Private Sub TxtCustCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCustCode.KeyDown
        If e.KeyCode = Keys.F1 Then
            Me.cmdCustHelp.PerformClick()
        End If
    End Sub

    Private Sub TxtCustCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustCode.TextChanged
        Me.TxtInternalAssyPart.Text = ""
        Me.lblInternalAssyDesc.Text = ""
        Me.TxtCustAssyPart.Text = ""
        Me.lblCustAssyDesc.Text = ""
        Me.lblCustName.Text = ""
        Me.lblCustName.Text = ""
        Me.sprItems.MaxRows = 0
        If Me.cmdGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            AddBlankRow()
        End If
        Me.SprTaxDtl.MaxRows = 0
    End Sub


    Private Sub TxtCustCode_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TxtCustCode.Validating
        Dim strQry As String
        Dim dtCustomer As New DataTable
        If Me.TxtCustCode.Text.Trim = "" Then Exit Sub
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            If cmdGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                strQry = "SELECT A.CUSTOMER_CODE, A.CUST_NAME "
                strQry += " FROM CUSTOMER_MST A INNER JOIN CUSTITEM_MST B"
                strQry += " ON A.CUSTOMER_CODE=B.ACCOUNT_CODE and A.unit_code = B.unit_code  WHERE A.CUSTOMER_CODE= '" & Me.TxtCustCode.Text.Trim & "' and A.unit_code= '" & gstrUNITID & "' "
                strQry += " GROUP BY A.CUSTOMER_CODE,A.CUST_NAME"
            Else
                strQry = "SELECT B.CUSTOMER_CODE,B.CUST_NAME"
                strQry += " FROM AGREEMENT_HDR A INNER JOIN CUSTOMER_MST B"
                strQry += " ON A.CUSTOMER_CODE=B.CUSTOMER_CODE and A.unit_code = B.unit_code"
                strQry += " WHERE A.ACTIVE=1 AND A.CUSTOMER_CODE= '" & Me.TxtCustCode.Text.Trim & "' and A.unit_code = '" & gstrUNITID & "'"
                strQry += " GROUP BY B.CUSTOMER_CODE,B.CUST_NAME"
            End If
            Call GetData(strQry, dtCustomer)
            If dtCustomer.Rows.Count = 0 Then
                MsgBox("Invalid Customer Code", MsgBoxStyle.Information, ResolveResString(100))
                Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
                e.Cancel = True
                Exit Sub
            End If
            Me.lblCustName.Text = dtCustomer.Rows(0).Item("cust_name")
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)

        Catch EX As Exception
            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Private Sub GetData(ByVal strQry As String, ByRef dt As DataTable)
        Dim ConnString As String = "Server=" & gstrCONNECTIONSERVER & ";Database=" & gstrDatabaseName & ";user=" & gstrCONNECTIONUSER & ";password=" & gstrCONNECTIONPASSWORD & " "
        Dim oDataAdapter As SqlDataAdapter
        Try
            oDataAdapter = New SqlDataAdapter(strQry, ConnString)
            oDataAdapter.Fill(dt)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub TxtInternalAssyPart_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtInternalAssyPart.KeyDown
        If e.KeyCode = Keys.F1 Then
            Me.cmdInternalPartHelp.PerformClick()
        End If
    End Sub

    Private Sub TxtInternalAssyPart_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtInternalAssyPart.TextChanged
        Me.lblInternalAssyDesc.Text = ""
        Me.TxtCustAssyPart.Text = ""
        Me.lblCustAssyDesc.Text = ""
        Me.sprItems.MaxRows = 0
        If Me.cmdGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            AddBlankRow()
        End If
        Me.SprTaxDtl.MaxRows = 0
    End Sub

    Private Sub TxtInternalAssyPart_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TxtInternalAssyPart.Validating
        Dim strQry As String
        Dim dt As New DataTable
        If TxtInternalAssyPart.Text.Trim = "" Then Exit Sub
        If Me.TxtCustCode.Text.Trim = "" Then
            MsgBox("Please First Select Customer", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If
        Try
            If cmdGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                strQry = " SELECT B.ITEM_CODE,B.DESCRIPTION "
                strQry += " FROM CUSTITEM_MST A INNER JOIN ITEM_MST B"
                strQry += " ON A.ITEM_CODE=B.ITEM_CODE and A.unit_code = B.unit_code "
                strQry += " WHERE B.ITEM_MAIN_GRP='F' AND A.ACCOUNT_CODE='" & Me.TxtCustCode.Text.Trim & "' AND B.ITEM_CODE='" & Me.TxtInternalAssyPart.Text.Trim & "' and A.unit_code = '" & gstrUNITID & "'"
                strQry += " GROUP BY B.ITEM_CODE,B.DESCRIPTION"
            Else
                strQry = " SELECT INTERNALASSYPART,DESCRIPTION "
                strQry += " FROM AGREEMENT_HDR A INNER JOIN ITEM_MST B"
                strQry += " ON A.INTERNALASSYPART=B.ITEM_CODE and A.unit_code = B.unit_code"
                strQry += " WHERE(A.ACTIVE = 1) AND A.CUSTOMER_CODE='" & Me.TxtCustCode.Text.Trim & "' AND A.INTERNALASSYPART='" & Me.TxtInternalAssyPart.Text.Trim & "' and A.unit_code = '" & gstrUNITID & "' "
                strQry += " GROUP BY INTERNALASSYPART,DESCRIPTION "
            End If
            GetData(strQry, dt)
            If dt.Rows.Count = 0 Then
                MsgBox("Invalid Internal Assy. Part Code", MsgBoxStyle.Information, ResolveResString(100))
                Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
                e.Cancel = True
                Exit Sub
            End If
            Me.lblInternalAssyDesc.Text = dt.Rows(0).Item("DESCRIPTION")
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)


        Catch EX As Exception
            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Private Sub TxtCustAssyPart_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TxtCustAssyPart.KeyDown
        If e.KeyCode = Keys.F1 Then
            Me.cmdCustPartHelp.PerformClick()
        End If
    End Sub

    Private Sub TxtCustAssyPart_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtCustAssyPart.TextChanged
        Me.lblCustAssyDesc.Text = ""
        Me.sprItems.MaxRows = 0
        If Me.cmdGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            AddBlankRow()
        End If
        Me.SprTaxDtl.MaxRows = 0
    End Sub

    Private Sub TxtCustAssyPart_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TxtCustAssyPart.Validating
        Dim strQry As String
        Dim dt As New DataTable
        If TxtCustAssyPart.Text.Trim = "" Then Exit Sub
        If Me.TxtCustCode.Text.Trim = "" Then
            MsgBox("Please First Select Customer", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        ElseIf Me.TxtInternalAssyPart.Text.Trim = "" Then
            MsgBox("Please First Select Internal Assy. Part", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If
        Try
            If cmdGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                strQry = " SELECT A.CUST_DRGNO,DRG_DESC"
                strQry += " FROM CUSTITEM_MST A INNER JOIN ITEM_MST B"
                strQry += " ON A.ITEM_CODE=B.ITEM_CODE  and A.unit_code = B.unit_code "
                strQry += " WHERE B.ITEM_MAIN_GRP='F' AND A.ACTIVE=1"
                strQry += " AND A.ACCOUNT_CODE='" & Me.TxtCustCode.Text.Trim & "' AND A.ITEM_CODE='" & Me.TxtInternalAssyPart.Text.Trim & "' AND A.CUST_DRGNO='" & Me.TxtCustAssyPart.Text.Trim & "' and  A.unit_code = '" & gstrUNITID & "'"
            Else
                strQry = "SELECT CUSTASSYPART,DRG_DESC"
                strQry += " FROM AGREEMENT_HDR A INNER JOIN CUSTITEM_MST B"
                strQry += " ON A.CUSTASSYPART=B.CUST_DRGNO and A.unit_code = B.unit_code "
                strQry += " WHERE(A.ACTIVE = 1 And B.ACTIVE = 1)"
                strQry += " AND B.ACCOUNT_CODE='" & Me.TxtCustCode.Text.Trim & "' AND B.ITEM_CODE='" & Me.TxtInternalAssyPart.Text.Trim & "' AND A.CUSTASSYPART='" & Me.TxtCustAssyPart.Text.Trim & "' and A.unit_code = '" & gstrUNITID & "'"
                strQry += " GROUP BY CUSTASSYPART,DRG_DESC"
            End If
            GetData(strQry, dt)
            If dt.Rows.Count = 0 Then
                MsgBox("Invalid Customer Assy. Part Code", MsgBoxStyle.Information, ResolveResString(100))
                Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
                e.Cancel = True
                Exit Sub
            End If
            Me.lblCustAssyDesc.Text = dt.Rows(0).Item("DRG_DESC")
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)

        Catch EX As Exception
            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Private Sub sprItems_LeaveCell(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles sprItems.LeaveCell

        Dim strQry As String
        Dim dt As New DataTable
        If e.newRow = -1 Or e.newCol = -1 Then Exit Sub
        Try
            If cmdGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If e.col = EnmGridCol.CustPartNo Then
                    With Me.sprItems
                        .Row = e.row : .Col = EnmGridCol.CustPartNo
                        If .Text.Trim = "" Then Exit Sub

                        Dim strItem As String
                        If Me.TxtCustCode.Text.Trim = "" Then
                            MsgBox("Please First Select Customer", MsgBoxStyle.Information, ResolveResString(100))
                            Exit Sub
                        End If

                        .Row = e.row : .Col = EnmGridCol.CustPartNo
                        strQry = " SELECT A.ITEM_CODE,B.DESCRIPTION,B.TYPE_CODE "
                        strQry += " FROM CUSTITEM_MST A INNER JOIN ITEM_MST B"
                        strQry += " ON A.ITEM_CODE=B.ITEM_CODE and A.unit_code = B.unit_code "
                        strQry += " WHERE B.ITEM_MAIN_GRP <> 'F' AND A.ACTIVE=1"
                        strQry += " AND A.ACCOUNT_CODE='" & Me.TxtCustCode.Text.Trim & "' AND A.CUST_DRGNO='" & .Text & "' and A.unit_code = '" & gstrUNITID & "'"
                        GetData(strQry, dt)
                        If dt.Rows.Count = 0 Then
                            MsgBox("Invalid Customer Part No.", MsgBoxStyle.Information, ResolveResString(100))
                            e.cancel = True
                            Exit Sub
                        Else
                            .Row = e.row
                            .Col = EnmGridCol.InternalPartNo : .Text = dt.Rows(0).Item("ITEM_CODE")
                            .Col = EnmGridCol.ProcessType : .Text = dt.Rows(0).Item("TYPE_CODE")
                            .Col = EnmGridCol.BasicValue : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        End If
                    End With

                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Private Sub CmdAgreementNoHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdAgreementNoHelp.Click
        Dim strQry As String
        Dim strDocNo() As String
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            If cmdGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                strQry = "SELECT DOC_NO AS AGREEMENT_NO,CUSTOMER_CODE FROM AGREEMENT_HDR WHERE ACTIVE=1 and unit_code = '" & gstrUNITID & "'"
                If Me.TxtCustCode.Text.Trim <> "" Then
                    strQry += " AND CUSTOMER_CODE='" & Me.TxtCustCode.Text.Trim & "'"
                End If
                If Me.TxtInternalAssyPart.Text.Trim <> "" Then
                    strQry += " AND INTERNALASSYPART='" & Me.TxtInternalAssyPart.Text.Trim & "'"
                End If
                If Me.TxtCustAssyPart.Text.Trim <> "" Then
                    strQry += " AND CUSTASSYPART='" & Me.TxtCustAssyPart.Text.Trim & "'"
                End If

                strDocNo = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
                If Not (UBound(strDocNo) = -1) Then
                    If (Len(strDocNo(0)) >= 1) And strDocNo(0) = "0" Then
                        MsgBox("No Agreement No. Found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                        Exit Sub
                    Else
                        Me.TxtAgreementNo.Text = strDocNo(0)
                        PopulateRecord(Val(TxtAgreementNo.Text), "ADD")
                    End If
                End If
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Catch EX As Exception
            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub
    Private Function PopulateRecord(ByVal intDocNo As Integer, ByRef pstrMode As String) As Boolean
        Dim strQry As String
        Dim sqlConn As SqlConnection
        Dim dtAgreement As New DataTable
        Dim dtAgreementHdr As New DataTable
        Dim dtTax As New DataTable
        Dim dtCost As New DataTable
        Dim intRow As Integer
        Dim Row As DataRow
        Try
            strQry = "SELECT DOC_NO,CUSTOMER_CODE,CUST_NAME,INTERNALASSYPART,ITEM_DESC,CUSTASSYPART,DRG_DESC,LABOURCOST,TOTALBASICVALUE,ENT_DT,ACTIVE FROM VW_AGREEMENT WHERE DOC_NO=" & intDocNo & " and Unit_code = '" & gstrUNITID & "'"
            GetData(strQry, dtAgreementHdr)
            With dtAgreementHdr
                If .Rows.Count > 0 Then
                    cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                    Me.TxtCustCode.Text = .Rows(0).Item("customer_code")
                    Me.lblCustName.Text = .Rows(0).Item("cust_name")
                    Me.TxtInternalAssyPart.Text = .Rows(0).Item("InternalAssyPart")
                    Me.lblInternalAssyDesc.Text = .Rows(0).Item("item_desc")
                    Me.TxtCustAssyPart.Text = .Rows(0).Item("CustAssyPart")
                    Me.lblCustAssyDesc.Text = .Rows(0).Item("Drg_Desc")
                    Me.chkActive.Checked = .Rows(0).Item("Active")
                    Me.dtAgreementDate.Value = .Rows(0).Item("Ent_dt")
                    Me.TxtLabourCost.Text = .Rows(0).Item("LABOURCOST")
                    Me.lblTotalBasicValue.Text = .Rows(0).Item("TOTALBASICVALUE")
                Else
                    MsgBox("No record found", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Function
                End If
            End With
            strQry = " SELECT DOC_NO,PO_NO,CUSTPARTNO,INTERNALPARTNO,PROCESSTYPE,BASIC_VALUE,USAGE "
            strQry += " FROM  AGREEMENT_DTL WHERE DOC_NO=" & intDocNo & " and Unit_code = '" & gstrUNITID & "'"
            GetData(strQry, dtAgreement)
            If dtAgreement.Rows.Count > 0 Then
                With Me.sprItems
                    .MaxRows = 0
                    For Each Row In dtAgreement.Rows
                        AddBlankRow()
                        .Row = .MaxRows
                        .Col = EnmGridCol.PO_No : .Text = Row("PO_no")
                        .Col = EnmGridCol.CustPartNo : .Text = Row("CUSTPARTNO")
                        .Col = EnmGridCol.InternalPartNo : .Text = Row("INTERNALPARTNO")
                        .Col = EnmGridCol.ProcessType : .Text = Row("PROCESSTYPE")
                        .Col = EnmGridCol.BasicValue : .Text = Row("BASIC_VALUE")
                        .Col = EnmGridCol.Usage : .Text = Row("USAGE")
                    Next
                    If pstrMode = "ADD" Then
                    .Row = 1 : .Row2 = .MaxRows
                    .Col = EnmGridCol.PO_No : .Col2 = EnmGridCol.Usage
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                    End If
                    If pstrMode = "EDIT" Then
                        .Row = 1 : .Row2 = .MaxRows
                        .Col = EnmGridCol.PO_No : .Col2 = EnmGridCol.Usage
                        .BlockMode = True
                        .Lock = False
                        .BlockMode = False
                    End If
                End With
            End If

            strQry = " SELECT DOC_NO,TAXID,TAXVALUE FROM AGREEMENTTAXDTL WHERE DOC_NO= " & intDocNo & " and unit_code = '" & gstrUNITID & "'"
            GetData(strQry, dtTax)
            If dtTax.Rows.Count > 0 Then
                With Me.SprTaxDtl
                    .MaxRows = 0
                    For Each Row In dtTax.Rows
                        AddTaxDtlRow()
                        .Row = .MaxRows
                        .Col = EnmTaxDtl.TaxId : .Text = Row("TAXID")
                        .Col = EnmTaxDtl.TaxValue : .Text = Row("TAXVALUE")
                    Next
                    .Row = 1 : .Row2 = .MaxRows
                    .Col = EnmTaxDtl.TaxHelp : .Col2 = EnmTaxDtl.TaxValue
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                End With
            End If

            strQry = "SELECT DOC_NO,COST_DESC,COST_VALUE FROM AGREEMENTLABOURCOSTDTL WHERE DOC_NO=" & intDocNo & " and Unit_code = '" & gstrUNITID & "'"
            GetData(strQry, dtCost)
            If dtCost.Rows.Count > 0 Then
                With Me.SprLbrCost
                    .MaxRows = 0
                    For Each Row In dtCost.Rows
                        AddLabourCostRow()
                        .Row = .MaxRows
                        .Col = EnmLbrCost.Cost_Desc : .Text = Row("cost_desc")
                        .Col = EnmLbrCost.Cost_Value : .Text = Row("cost_value")
                    Next
                    .Row = 1 : .Row2 = .MaxRows
                    .Col = EnmLbrCost.Cost_Desc : .Col2 = EnmLbrCost.Cost_Value
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                End With
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Function

    Private Sub TxtAgreementNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtAgreementNo.TextChanged
        If Me.cmdGrp.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            Me.SprTaxDtl.MaxRows = 0
            Me.sprItems.MaxRows = 0
            Me.SprLbrCost.MaxRows = 0
            Me.TxtCustCode.Text = ""
            Me.TxtInternalAssyPart.Text = ""
            Me.TxtCustAssyPart.Text = ""
            Me.lblCustAssyDesc.Text = ""
            Me.lblCustName.Text = ""
            Me.lblInternalAssyDesc.Text = ""
            Me.lblTotalBasicValue.Text = ""
            Me.TxtLabourCost.Text = ""
            cmdGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
        End If
    End Sub

    Private Sub CmdLbrCostOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdLbrCostOK.Click
        Dim intRow As Integer
        Dim dblCost As Double
        Dim strCostDesc As String
        Try
            With Me.SprLbrCost
                For intRow = 1 To .MaxRows
                    .Row = intRow
                    .Col = EnmLbrCost.Cost_Desc
                    strCostDesc = .Text.Trim
                    .Col = EnmLbrCost.Cost_Value
                    If Val(.Value) > 0 And strCostDesc = "" Then
                        MsgBox("Please enter Cost Description at row [ " & intRow & "].", MsgBoxStyle.Information, ResolveResString(100))
                        Exit Sub
                    End If
                    If strCostDesc <> "" And Val(.Value) = 0 Then
                        MsgBox("Please Cost Value at row [ " & intRow & "].", MsgBoxStyle.Information, ResolveResString(100))
                        Exit Sub
                    End If
                    dblCost = dblCost + Val(.Value)
                Next
                Me.TxtLabourCost.Text = dblCost

            End With
            GrpLabourCost.Visible = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try


    End Sub

    Private Sub cmdLabourCost_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLabourCost.Click
        With Me.GrpLabourCost
            .Visible = True
            .Top = cmdLabourCost.Top
            .Left = Me.TxtLabourCost.Left
            .BringToFront()
        End With

        If cmdGrp.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If Me.SprLbrCost.MaxRows = 0 Then
                AddLabourCostRow()
            End If
        End If
    End Sub


    Private Sub SprLbrCost_KeyDownEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles SprLbrCost.KeyDownEvent
        If cmdGrp.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If e.keyCode = Keys.Enter Then
                With Me.SprLbrCost
                    .Row = .ActiveRow
                    .Col = EnmLbrCost.Cost_Value
                    If .Text.Trim <> "" Then
                        AddLabourCostRow()
                    End If
                End With
            ElseIf e.keyCode = Keys.Delete Then
                With Me.SprLbrCost
                    If MsgBox("Do you want to delete this row", MsgBoxStyle.Question + MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                        .Row = .ActiveRow
                        .Action = FPSpreadADO.ActionConstants.ActionDeleteRow
                        .MaxRows = .MaxRows - 1
                    End If
                End With
            End If
        End If
    End Sub


    Private Sub sprItems_EditChange(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles sprItems.EditChange
        Try
            If e.col = EnmGridCol.BasicValue Then
                Me.lblTotalBasicValue.Text = TotalBasicValue()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub
    Private Function TotalBasicValue() As Double
        Dim intRow As Integer
        Dim dblValue As Double = 0
        With Me.sprItems
            For intRow = 1 To .MaxRows
                .Row = intRow : .Col = EnmGridCol.BasicValue
                dblValue = dblValue + Val(.Value)
            Next
            dblValue = dblValue + Val(Me.TxtLabourCost.Text)
            Return dblValue
        End With
    End Function
    Private Function ValidateForNewRecord() As Boolean
        Dim dt As New DataTable
        Dim strQry As String
        Try
            ValidateForNewRecord = False
            strQry = "Select top  1 1 from agreement_hdr where customer_code='" & Me.TxtCustCode.Text.Trim & "' and InternalAssyPart='" & Me.TxtInternalAssyPart.Text.Trim & "' and custassypart='" & Me.TxtCustAssyPart.Text & "' and active=1 and Unit_code = '" & gstrUNITID & "'"
            Call GetData(strQry, dt)
            If dt.Rows.Count = 1 Then
                MsgBox("Record already exists for Customer[" & Me.TxtCustCode.Text & "], Internal Assy. Part[" & Me.TxtInternalAssyPart.Text.Trim & "] and Customer Assy. Part[" & Me.TxtCustAssyPart.Text.Trim & "]", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            End If
            ValidateForNewRecord = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Function
    Private Function ValidateChildParts() As Boolean
        Dim strIntrPart As String
        Dim strCustPart As String
        Dim intUsage As Integer
        Dim intRow As Integer
        Dim intDocNo As Object = Nothing
        Dim strQry As String
        Dim sqlCmd As New SqlCommand
        Dim sqlConn As New SqlConnection
        Try
            ValidateChildParts = False
            sqlConn = SqlConnectionclass.GetConnection
            sqlCmd.Connection = sqlConn
            strQry = "DELETE TMPAGREEMENTCHILDPART WHERE IP_ADDRESS='" & gstrIpaddressWinSck & "' and unit_code = '" & gstrUNITID & "'"
            sqlCmd.CommandText = strQry
            sqlCmd.CommandType = CommandType.Text
            sqlCmd.ExecuteNonQuery()
            With Me.sprItems
                For intRow = 1 To .MaxRows
                    .Row = intRow
                    .Col = EnmGridCol.InternalPartNo : strIntrPart = .Text.Trim
                    .Col = EnmGridCol.CustPartNo : strCustPart = .Text.Trim
                    .Col = EnmGridCol.Usage : intUsage = Val(.Value)
                    strQry = "INSERT INTO TMPAGREEMENTCHILDPART(CUSTPARTNO,INTERNALPARTNO,USAGE,IP_ADDRESS,unit_code)"
                    strQry += " VALUES('" & strCustPart & "','" & strIntrPart & "'," & intUsage & ",'" & gstrIpaddressWinSck & "','" & gstrUNITID & "')"
                    sqlCmd.CommandText = strQry
                    sqlCmd.ExecuteNonQuery()
                Next
            End With
            strQry = "SELECT DBO.FN_VALIDATE_AGREEMENTCHILDPART('" & gstrIpaddressWinSck & "','" & gstrUNITID & "') AS DOC_NO"
            sqlCmd.CommandText = strQry
            intDocNo = sqlCmd.ExecuteScalar
            sqlConn.Close()
            If Convert.ToInt32(intDocNo) > 0 Then
                MsgBox("Customer Part Details already exists with Agreement No [" & intDocNo & "]. Cannot save new record.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            End If
            ValidateChildParts = True
        Catch ex As Exception
            If sqlConn.State = ConnectionState.Open Then sqlConn.Close()
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Function
End Class
