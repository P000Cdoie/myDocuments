'---------------------------------------------------------------------------------
'COPYRIGHT(C)   : MOTHERSON SUMI INFOTECH & DESIGN LTD.
'FORM NAME      : FRMMKTTRN0088 - DAILY MARKETTING SCHEDULE - VECHICLE BOM
'CREATED BY     : VINOD SINGH 
'CREATED DATE   : 12 FEB 2015
'ISSUE ID       : 10737738 - eMPro Vehicle BOM 
'---------------------------------------------------------------------------------

Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Public Class frmMKTTRN0088

#Region "Form Level Declaration"

    Dim mintFormIndex As Integer
    Dim mdtbBreakupSch As DataTable

    Private Enum enmConsolidate
        CHK = 1
        ItemCode
        ItemDesc
        CustItemCode
        CustItemDesc
        Qty
        DailyBreakUp
    End Enum

    Private Enum enmBreakup
        SchDate = 1
        ItemCode
        ItemDesc
        CustItemCode
        CustItemDesc
        Qty
    End Enum

    Dim ListModelVariant As New List(Of ModelVariantStructure)

#End Region

#Region "Methods"

    Private Sub AddBlankRowForConsolidate()
        With Me.SprConsolidate
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .BackColorStyle = FPSpreadADO.BackColorStyleConstants.BackColorStyleUnderGrid
            .Col = enmConsolidate.CHK : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .TypeCheckCenter = True : .Value = "1"
            .Col = enmConsolidate.ItemCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = enmConsolidate.ItemDesc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = enmConsolidate.CustItemCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = enmConsolidate.CustItemDesc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = enmConsolidate.Qty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Lock = True
            .Col = enmConsolidate.DailyBreakUp : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = "Breakup"
        End With
    End Sub

    Private Sub AddBlankRowForDailyBreakup()
        With Me.SprBreakup
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .BackColorStyle = FPSpreadADO.BackColorStyleConstants.BackColorStyleUnderGrid
            .Col = enmBreakup.SchDate : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = enmBreakup.ItemCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = enmBreakup.ItemDesc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = enmBreakup.CustItemCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = enmBreakup.CustItemDesc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = enmBreakup.Qty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .Lock = True
        End With
    End Sub

    Private Sub SetGridHeading()
        Try
            With Me.SprConsolidate
                .MaxRows = 0
                .MaxCols = [Enum].GetNames(GetType(enmConsolidate)).Count
                .Row = 0
                .EditEnterAction = FPSpreadADO.EditEnterActionConstants.EditEnterActionNext
                .EditModeReplace = True
                .set_RowHeight(0, 20)
                .Col = 0 : .set_ColWidth(0, 3)
                .Col = enmConsolidate.CHK : .Text = " " : .set_ColWidth(enmConsolidate.CHK, 3) : .TypeCheckCenter = True
                .Col = enmConsolidate.ItemCode : .Text = "Item Code" : .set_ColWidth(enmConsolidate.ItemCode, 14)
                .Col = enmConsolidate.ItemDesc : .Text = "Item Description" : .set_ColWidth(enmConsolidate.ItemDesc, 18)
                .Col = enmConsolidate.CustItemCode : .Text = "Customer Item Code" : .set_ColWidth(enmConsolidate.CustItemCode, 14)
                .Col = enmConsolidate.CustItemDesc : .Text = "Customer Item Description" : .set_ColWidth(enmConsolidate.CustItemDesc, 18)
                .Col = enmConsolidate.Qty : .Text = "Consolidate Month Qty." : .set_ColWidth(enmConsolidate.Qty, 8)
                .Col = enmConsolidate.DailyBreakUp : .Text = "Daily Breakup" : .set_ColWidth(enmConsolidate.DailyBreakUp, 8)
            End With

            With Me.SprBreakup
                .MaxRows = 0
                .MaxCols = [Enum].GetNames(GetType(enmBreakup)).Count
                .Row = 0
                .EditEnterAction = FPSpreadADO.EditEnterActionConstants.EditEnterActionNext
                .EditModeReplace = True
                .set_RowHeight(0, 20)
                .Col = 0 : .set_ColWidth(0, 3)
                .Col = enmBreakup.SchDate : .Text = "Schedule Date" : .set_ColWidth(enmBreakup.SchDate, 10)
                .Col = enmBreakup.ItemCode : .Text = "Item Code" : .set_ColWidth(enmBreakup.ItemCode, 16) : .ColHidden = True
                .Col = enmBreakup.ItemDesc : .Text = "Item Description" : .set_ColWidth(enmBreakup.ItemDesc, 30)
                .Col = enmBreakup.CustItemCode : .Text = "Customer Item Code" : .set_ColWidth(enmBreakup.CustItemCode, 16) : .ColHidden = True
                .Col = enmBreakup.CustItemDesc : .Text = "Customer Item Description" : .set_ColWidth(enmBreakup.CustItemDesc, 30)
                .Col = enmBreakup.Qty : .Text = "Schedule Qty." : .set_ColWidth(enmBreakup.Qty, 15)

            End With

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub FillSearchCategory()
        Try
            With cboSearchCategory
                .Items.Clear()
                .DataSource = [Enum].GetNames(GetType(enmConsolidate))
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

#Region "Form & Controls Events"

    Private Sub frmMKTTRN0088_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub frmMKTTRN0088_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Tag) = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub frmMKTTRN0088_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            Me.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub frmMKTTRN0088_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            If MsgBox("Do you want to close the screen?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, ResolveResString(100)) = MsgBoxResult.No Then
                e.Cancel = True
                Return
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub frmMKTTRN0088_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
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

    Private Sub frmMKTTRN0088_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Call FitToClient(Me, GrpMain, CtlHeader, Me.btnGrp, 500)
            Me.MdiParent = mdifrmMain
            mintFormIndex = mdifrmMain.AddFormNameToWindowList(Me.CtlHeader.Tag)
            SetGridHeading()
            btnGrp.ShowButtons(True, False, False, False)
            cmdCustHelp.Enabled = False
            SprBreakup.Enabled = False
            SprConsolidate.Enabled = False
            btnConsolidateSchedule.Enabled = False
            btnSelectModelVariant.Enabled = False
            dtpSchMonth.Format = DateTimePickerFormat.Custom
            dtpSchMonth.CustomFormat = "MMM/yyyy"
            FillSearchCategory()
        Catch ex As Exception
            RaiseException(ex)
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub cmdCustHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCustHelp.Click
        Dim strVal() As String
        Dim strSQL As String = String.Empty
        Try
            If btnGrp.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then Return
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
            strSQL = " SELECT A.CUSTOMER_CODE,A.CUST_NAME FROM CUSTOMER_MST A INNER JOIN BUDGETITEM_MST B " & _
                    " ON A.UNIT_CODE = B.UNIT_CODE AND A.CUSTOMER_CODE = B.ACCOUNT_CODE " & _
                    " WHERE A.UNIT_CODE ='" & gstrUNITID & "' and SCH_UPLOAD_CODE='VEHBOM'  GROUP BY A.CUSTOMER_CODE,A.CUST_NAME "

            strVal = ctlHelp1.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, ResolveResString(100))

            If Not strVal Is Nothing Then
                If UBound(strVal) > 0 And strVal(0) <> "0" Then
                    txtCustCode.Text = strVal(0)
                    txtCustName.Text = strVal(1)
                Else
                    MsgBox("No record found.", MsgBoxStyle.Information, ResolveResString(100))
                End If
            Else
                MsgBox("No record found.", MsgBoxStyle.Information, ResolveResString(100))
            End If

            SprBreakup.MaxRows = 0
            SprConsolidate.MaxRows = 0
            btnConsolidateSchedule.Enabled = True
            btnSelectModelVariant.Enabled = True
            btnEditSch.Enabled = True
            btnSaveBreakupSch.Enabled = False
            txtBreakupItemCode.Text = ""
            txtConsolidateQty.Text = ""
            txtConsolidateSchSearch.Text = ""
            txtCustItemCode.Text = ""
            ListModelVariant = New List(Of ModelVariantStructure)
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub btnGrp_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles btnGrp.ButtonClick
        Dim strSQL As String
        Dim blnSelected As Boolean
        Try
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
            Select Case e.Button
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                    strSQL = "SELECT * FROM SALES_PARAMETER WHERE UNIT_CODE='" & gstrUNITID & "' AND VEHICLE_BOM =1"
                    If IsRecordExists(strSQL) = False Then
                        MsgBox("Schedule Generation using Vehicle BOM functionality is not enabled for this Unit.", MsgBoxStyle.Critical, ResolveResString(100))
                        btnGrp.Revert()
                        Return
                    End If
                    txtCustCode.Text = ""
                    txtCustName.Text = ""
                    SprBreakup.MaxRows = 0
                    SprConsolidate.MaxRows = 0
                    cmdCustHelp.Enabled = True
                    SprBreakup.Enabled = True
                    SprConsolidate.Enabled = True
                    btnConsolidateSchedule.Enabled = True
                    btnSelectModelVariant.Enabled = True
                    txtBreakupItemCode.Text = ""
                    txtConsolidateQty.Text = ""
                    txtConsolidateSchSearch.Text = ""
                    txtCustItemCode.Text = ""
                    ListModelVariant = New List(Of ModelVariantStructure)
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                    If btnEditSch.Enabled = False And btnSaveBreakupSch.Enabled = True Then
                        MsgBox("Record(s) of daily breakup is not saved yet, Either save or cancel it before saving schedule.", MsgBoxStyle.Critical, ResolveResString(100))
                        Return
                    End If

                    Using sqlCmd As SqlCommand = New SqlCommand
                        strSQL = "DELETE TMP_CONSOLIDATE_SCHEDULE WHERE UNIT_CODE = @UNIT_CODE AND IP_ADDRESS =@IP_ADDRESS AND ITEM_CODE =@ITEM_CODE AND CUST_ITEM_CODE = @CUST_ITEM_CODE"

                        sqlCmd.CommandText = strSQL
                        sqlCmd.CommandType = CommandType.Text
                        sqlCmd.CommandTimeout = 0
                        sqlCmd.Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                        sqlCmd.Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                        sqlCmd.Parameters.Add("@ITEM_CODE", SqlDbType.VarChar, 16)
                        sqlCmd.Parameters.Add("@CUST_ITEM_CODE", SqlDbType.VarChar, 50)
                        With SprConsolidate
                            SqlConnectionclass.BeginTrans()
                            For intRow As Integer = 1 To .MaxRows
                                .Row = intRow
                                .Col = enmConsolidate.CHK
                                If .Value = "1" Then
                                    blnSelected = True
                                Else
                                    .Col = enmConsolidate.ItemCode
                                    sqlCmd.Parameters("@ITEM_CODE").Value = .Text.Trim
                                    .Col = enmConsolidate.CustItemCode
                                    sqlCmd.Parameters("@CUST_ITEM_CODE").Value = .Text.Trim
                                    SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                                End If
                            Next
                            SqlConnectionclass.CommitTran()
                        End With
                    End Using

                    If blnSelected = False Then
                        MsgBox("Atleast one item must be selected to save records.", MsgBoxStyle.Exclamation, ResolveResString(100))
                        Return
                    End If
                    Using sqlCmdSave As SqlCommand = New SqlCommand
                        With sqlCmdSave
                            .CommandText = "USP_SAVE_SCHEDULE_VEHICLEBOM"
                            .CommandTimeout = 0
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                            .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                            .Parameters.Add("@CUSTOMER_CODE", SqlDbType.VarChar, 10).Value = txtCustCode.Text.Trim
                            .Parameters.Add("@SCHEDULE_MONTH_YYYYMM", SqlDbType.Int).Value = dtpSchMonth.Value.ToString("yyyyMM")
                            .Parameters.Add("@USER_ID", SqlDbType.VarChar, 16).Value = mP_User
                            SqlConnectionclass.ExecuteNonQuery(sqlCmdSave)
                        End With
                    End Using
                    MsgBox("Records Saved Successfully.", MsgBoxStyle.Information, ResolveResString(100))
                    txtCustCode.Text = ""
                    txtCustName.Text = ""
                    SprConsolidate.MaxRows = 0
                    SprBreakup.MaxRows = 0
                    cmdCustHelp.Enabled = False
                    SprBreakup.Enabled = False
                    SprConsolidate.Enabled = False
                    btnConsolidateSchedule.Enabled = False
                    btnSelectModelVariant.Enabled = False
                    btnEditSch.Enabled = True
                    btnSaveBreakupSch.Enabled = False
                    txtConsolidateSchSearch.Text = ""
                    ListModelVariant = New List(Of ModelVariantStructure)
                    btnGrp.Revert()
                    txtBreakupItemCode.Text = ""
                    txtCustItemCode.Text = ""
                    txtConsolidateQty.Text = ""
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                    If MsgBox("Do you want to revert all changes/selection done in screen?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.No Then
                        Return
                    End If
                    txtCustCode.Text = ""
                    txtCustName.Text = ""
                    SprConsolidate.MaxRows = 0
                    SprBreakup.MaxRows = 0
                    cmdCustHelp.Enabled = False
                    SprBreakup.Enabled = False
                    SprConsolidate.Enabled = False
                    btnConsolidateSchedule.Enabled = False
                    btnSelectModelVariant.Enabled = False
                    btnEditSch.Enabled = True
                    btnSaveBreakupSch.Enabled = False
                    txtConsolidateSchSearch.Text = ""
                    ListModelVariant = New List(Of ModelVariantStructure)
                    btnGrp.Revert()
                    txtBreakupItemCode.Text = ""
                    txtCustItemCode.Text = ""
                    txtConsolidateQty.Text = ""
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                    Me.Close()
            End Select
        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub btnSelectModelVariant_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSelectModelVariant.Click
        Try
            If btnGrp.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                MsgBox("Action is allowed only in New Mode!", MsgBoxStyle.Information, ResolveResString(100))
                Return
            End If
            If txtCustCode.Text.Trim = "" Then
                MsgBox("Please select customer first !", MsgBoxStyle.Information, ResolveResString(100))
                Return
            End If

            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
            Dim frmModelVariant As New FRMMKTTRN0088A
            With frmModelVariant
                .CustomerCode = txtCustCode.Text
                .CustomerName = txtCustName.Text
                .SelectedModelVariant = ListModelVariant
                .ShowDialog()
                If .DialogResult = Windows.Forms.DialogResult.OK Then
                    ListModelVariant = .SelectedModelVariant
                End If
            End With
            frmModelVariant = Nothing
        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub btnConsolidateSchedule_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConsolidateSchedule.Click
        Dim strSQL As String = String.Empty
        Try
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
            SprConsolidate.MaxRows = 0
            SprBreakup.MaxRows = 0
            btnEditSch.Enabled = True
            btnSaveBreakupSch.Enabled = False
            'GENERATE SCHEDULE
            strSQL = "DELETE TMP_CUST_MODEL_VARIANT WHERE UNIT_CODE='" & gstrUNITID & "' AND IP_ADDRESS='" & gstrIpaddressWinSck & "'"
            SqlConnectionclass.ExecuteNonQuery(strSQL)
            Dim ITEMS As IEnumerator = ListModelVariant.GetEnumerator
            While ITEMS.MoveNext
                Dim ModelVariant As ModelVariantStructure = ITEMS.Current
                strSQL = "INSERT INTO TMP_CUST_MODEL_VARIANT(UNIT_CODE,IP_ADDRESS,CUST_CODE,MODEL_CODE,VARIANT_CODE,VOLUME)" & _
                "values ('" & gstrUNITID & "','" & gstrIpaddressWinSck & "','" & txtCustCode.Text.Trim & "','" & ModelVariant.ModelCode & "','" & ModelVariant.Variantcode & "'," & ModelVariant.Volume & " )"
                SqlConnectionclass.ExecuteNonQuery(strSQL)
            End While
            Using sqlcmd As SqlCommand = New SqlCommand
                With sqlcmd
                    .CommandText = "USP_GENERATE_SCHEDULE_VEHICLEBOM"
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 0
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@CUSTOMER_CODE", SqlDbType.VarChar, 10).Value = txtCustCode.Text.Trim
                    .Parameters.Add("@SCHEDULE_MONTH_YYYYMM", SqlDbType.Int).Value = dtpSchMonth.Value.ToString("yyyyMM")
                    .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                    SqlConnectionclass.ExecuteNonQuery(sqlcmd)
                End With
            End Using

            'FILL CONSOLIDATE SCHEDULE
            strSQL = "SELECT * FROM TMP_CONSOLIDATE_SCHEDULE WHERE UNIT_CODE = '" & gstrUNITID & "' AND IP_ADDRESS = '" & gstrIpaddressWinSck & "' "
            Using dtbConsolidateSch As DataTable = SqlConnectionclass.GetDataTable(strSQL)
                If dtbConsolidateSch.Rows.Count = 0 Then
                    MsgBox("No record found !", MsgBoxStyle.Information, ResolveResString(100))
                    Return
                End If
                chkSelAll.CheckState = CheckState.Checked
                With SprConsolidate
                    .MaxRows = 0
                    For Each row As DataRow In dtbConsolidateSch.Rows
                        AddBlankRowForConsolidate()
                        .Row = .MaxRows
                        .Col = enmConsolidate.ItemCode : .Text = Convert.ToString(row("ITEM_CODE"))
                        .Col = enmConsolidate.ItemDesc : .Text = Convert.ToString(row("ITEM_DESC"))
                        .Col = enmConsolidate.CustItemCode : .Text = Convert.ToString(row("CUST_ITEM_CODE"))
                        .Col = enmConsolidate.CustItemDesc : .Text = Convert.ToString(row("CUST_ITEM_DESC"))
                        .Col = enmConsolidate.Qty : .Text = Convert.ToString(row("TOTAL_REQD_QTY"))
                    Next
                End With

            End Using

            'FILL DAILY BREACK-UP SCHEDULE
            strSQL = "SELECT * FROM TMP_DAILY_SCHEDULE WHERE UNIT_CODE = '" & gstrUNITID & "' AND IP_ADDRESS = '" & gstrIpaddressWinSck & "' ORDER BY ITEM_CODE,CUST_ITEM_CODE,TRANS_DATE"
            mdtbBreakupSch = New DataTable
            mdtbBreakupSch = SqlConnectionclass.GetDataTable(strSQL)
            If mdtbBreakupSch.Rows.Count = 0 Then
                MsgBox("Daily Breack-Up Schedule Not Generated.", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return
            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub SprConsolidate_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles SprConsolidate.ButtonClicked
        Try
            With SprConsolidate
                If e.col = enmConsolidate.CHK Then
                    .Row = e.row
                    .Col = e.col
                    If .Value = "1" Then
                        .Col = enmConsolidate.DailyBreakUp : .Col2 = enmConsolidate.DailyBreakUp
                        .Row = e.row : .Row2 = e.row
                        .BlockMode = True
                        .Lock = False
                        .BlockMode = False
                    Else
                        .Col = enmConsolidate.DailyBreakUp : .Col2 = enmConsolidate.DailyBreakUp
                        .Row = e.row : .Row2 = e.row
                        .BlockMode = True
                        .Lock = True
                        .BlockMode = False
                    End If
                ElseIf e.col = enmConsolidate.DailyBreakUp Then
                    Dim strItem, strCustItem As String
                    If IsNothing(mdtbBreakupSch) = True Then Return
                    If mdtbBreakupSch.Rows.Count > 0 Then
                        With SprConsolidate
                            .Row = e.row
                            .Col = enmConsolidate.ItemCode
                            strItem = .Text.Trim
                            .Col = enmConsolidate.CustItemCode
                            strCustItem = .Text.Trim

                            txtBreakupItemCode.Text = strItem
                            .Col = enmConsolidate.Qty
                            txtConsolidateQty.Text = Val(.Value)
                            .Col = enmConsolidate.CustItemCode
                            txtCustItemCode.Text = .Text.Trim
                        End With

                        Dim rows() As DataRow = mdtbBreakupSch.Select("ITEM_CODE ='" & strItem.Trim & "' AND CUST_ITEM_CODE='" & strCustItem.Trim & "'")

                        With SprBreakup
                            .MaxRows = 0
                            For Each row As DataRow In rows
                                AddBlankRowForDailyBreakup()
                                .Row = .MaxRows
                                .Col = enmBreakup.SchDate : .Text = Convert.ToDateTime(row("TRANS_DATE")).ToString(gstrDateFormat)
                                .Col = enmBreakup.ItemCode : .Text = Convert.ToString(row("ITEM_CODE"))
                                .Col = enmBreakup.ItemDesc : .Text = Convert.ToString(row("ITEM_DESC"))
                                .Col = enmBreakup.CustItemCode : .Text = Convert.ToString(row("CUST_ITEM_CODE"))
                                .Col = enmBreakup.CustItemDesc : .Text = Convert.ToString(row("CUST_ITEM_DESC"))
                                .Col = enmBreakup.Qty : .Text = Convert.ToString(row("SCH_QTY"))
                            Next
                        End With
                        btnSaveBreakupSch.Enabled = False
                        btnEditSch.Enabled = True
                    End If
                End If

            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub btnSaveBreakupSch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSaveBreakupSch.Click
        Dim dblQtySum As Double
        Dim strSQL As String
        Dim dtSchDate As DateTime
        Dim strItem, strCustItem As String
        Dim dblQty As Double
        Try
            With SprBreakup
                If SprBreakup.MaxRows = 0 Then
                    MsgBox("No record found to save!", MsgBoxStyle.Exclamation, ResolveResString(100))
                    Return
                End If
                For intRow As Integer = 1 To SprBreakup.MaxRows
                    .Row = intRow
                    .Col = enmBreakup.Qty
                    dblQtySum += Val(.Value)
                Next

                If Val(txtConsolidateQty.Text) <> dblQtySum Then
                    MsgBox("Total of Day Wise Schedule Qty.[ " & dblQtySum & " ]  must be equal to Item's Consolidate Qty. [ " & txtConsolidateQty.Text & " ]", MsgBoxStyle.Exclamation, ResolveResString(100))
                    Return
                End If

                strSQL = "UPDATE TMP_DAILY_SCHEDULE SET SCH_QTY=@SCH_QTY WHERE UNIT_CODE = @UNIT_CODE AND IP_ADDRESS =@IP_ADDRESS AND ITEM_CODE =@ITEM_CODE AND TRANS_DATE = @SCH_DATE"
                Using sqlCmd As SqlCommand = New SqlCommand
                    SqlConnectionclass.BeginTrans()
                    sqlCmd.CommandText = strSQL
                    sqlCmd.CommandType = CommandType.Text
                    sqlCmd.CommandTimeout = 0
                    sqlCmd.Parameters.Add("@SCH_QTY", SqlDbType.Money)
                    sqlCmd.Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10)
                    sqlCmd.Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20)
                    sqlCmd.Parameters.Add("@ITEM_CODE", SqlDbType.VarChar, 16)
                    sqlCmd.Parameters.Add("@SCH_DATE", SqlDbType.DateTime)
                    For intRow As Integer = 1 To .MaxRows
                        .Row = intRow
                        .Col = enmBreakup.SchDate
                        dtSchDate = Convert.ToDateTime(.Value)
                        .Col = enmBreakup.ItemCode
                        strItem = .Text.Trim
                        .Col = enmBreakup.Qty
                        dblQty = Val(.Value)
                        .Col = enmBreakup.CustItemCode
                        strCustItem = .Text.Trim

                        sqlCmd.Parameters("@SCH_QTY").Value = dblQty
                        sqlCmd.Parameters("@UNIT_CODE").Value = gstrUNITID
                        sqlCmd.Parameters("@IP_ADDRESS").Value = gstrIpaddressWinSck
                        sqlCmd.Parameters("@ITEM_CODE").Value = strItem
                        sqlCmd.Parameters("@SCH_DATE").Value = getDateForDB(dtSchDate)
                        SqlConnectionclass.ExecuteNonQuery(sqlCmd)

                        Dim rows() As DataRow = mdtbBreakupSch.Select("ITEM_CODE ='" & strItem.Trim & "' AND CUST_ITEM_CODE='" & strCustItem.Trim & "' and TRANS_DATE ='" & getDateForDB(dtSchDate) & "'")
                        For Each row As DataRow In rows
                            row("SCH_QTY") = dblQty
                            Exit For
                        Next
                    Next
                    SqlConnectionclass.CommitTran()
                End Using

            End With
            MsgBox("Daily Breakup Schedule has been Changed", MsgBoxStyle.Information, ResolveResString(100))
            SprBreakup.MaxRows = 0
            With SprConsolidate
                .Row = 1 : .Row2 = .MaxRows
                .Col = 1 : .Col2 = .MaxCols
                .BlockMode = True
                .Lock = False
                .BackColor = Color.White
                .BlockMode = False

                For intRow As Integer = 1 To .MaxRows
                    .Row = intRow
                    .Col = enmConsolidate.CHK
                    If .Value = "0" Then
                        .Col = enmConsolidate.DailyBreakUp : .Lock = True
                    Else
                        .Col = enmConsolidate.DailyBreakUp : .Lock = False
                    End If
                Next
            End With

            txtBreakupItemCode.Text = ""
            txtConsolidateQty.Text = ""
            txtCustItemCode.Text = ""
            btnEditSch.Enabled = True
            btnSaveBreakupSch.Enabled = False
        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub btnEditSch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEditSch.Click
        Try
            If SprBreakup.MaxRows = 0 Then
                MsgBox("No record found to edit.", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return
            End If
            With SprConsolidate
                .Row = 1 : .Row2 = .MaxRows
                .Col = 1 : .Col2 = .MaxCols
                .BlockMode = True
                .Lock = True
                .BackColor = Color.LightGray
                .BlockMode = False
            End With

            With SprBreakup
                .Row = 1 : .Row2 = .MaxRows
                .Col = enmBreakup.Qty : .Col2 = enmBreakup.Qty
                .BlockMode = True
                .Lock = False
                .BlockMode = False
            End With
            btnEditSch.Enabled = False
            btnSaveBreakupSch.Enabled = True
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub BtnCancelSch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnCancelSch.Click
        Try
            If SprBreakup.MaxRows > 0 Then
                If MsgBox("Do you want to cancel the changes in daily breakup schedule?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.No Then
                    Return
                End If
            End If
            SprBreakup.MaxRows = 0
            With SprConsolidate
                .Row = 1 : .Row2 = .MaxRows
                .Col = 1 : .Col2 = .MaxCols
                .BlockMode = True
                .Lock = False
                .BackColor = Color.White
                .BlockMode = False

                For intRow As Integer = 1 To .MaxRows
                    .Row = intRow
                    .Col = enmConsolidate.CHK
                    If .Value = "0" Then
                        .Col = enmConsolidate.DailyBreakUp : .Lock = True
                    Else
                        .Col = enmConsolidate.DailyBreakUp : .Lock = False
                    End If
                Next
            End With
            txtBreakupItemCode.Text = ""
            txtConsolidateQty.Text = ""
            txtCustItemCode.Text = ""
            btnEditSch.Enabled = True
            btnSaveBreakupSch.Enabled = False
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtConsolidateSchSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtConsolidateSchSearch.TextChanged
        Dim intCounter As Integer
        Dim strText As String
        Dim Col As enmConsolidate
        Col = DirectCast(cboSearchCategory.SelectedIndex + 1, enmConsolidate)
        For intCounter = 1 To SprConsolidate.MaxRows
            With Me.SprConsolidate
                .Row = intCounter : .Col = Col : strText = Trim(.Text)
                If .FontBold = True Then
                    .FontBold = False
                    .Refresh()
                End If
            End With
        Next
        If Len(txtConsolidateSchSearch.Text) = 0 Then Exit Sub
        For intCounter = 1 To SprConsolidate.MaxRows
            With Me.SprConsolidate
                .Row = intCounter : .Col = Col : strText = Trim(.Text)
                If Trim(UCase(Mid(strText, 1, Len(txtConsolidateSchSearch.Text)))) = Trim(UCase(txtConsolidateSchSearch.Text)) Then
                    .Row = intCounter : .Col = Col : .FontBold = True
                    .Row = intCounter : .Col = Col : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    Exit For
                End If
            End With
        Next
    End Sub

    Private Sub chkSelAll_CheckStateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSelAll.CheckStateChanged
        Try
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
            With SprConsolidate
                For intRow As Integer = 1 To SprConsolidate.MaxRows
                    .Row = intRow
                    .Col = enmConsolidate.CHK
                    .Value = IIf(chkSelAll.CheckState = CheckState.Checked, "1", "0")
                Next
            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub dtpSchMonth_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpSchMonth.ValueChanged
        Try
            SprBreakup.MaxRows = 0
            SprConsolidate.MaxRows = 0
            btnConsolidateSchedule.Enabled = True
            btnSelectModelVariant.Enabled = True
            btnEditSch.Enabled = True
            btnSaveBreakupSch.Enabled = False
            txtBreakupItemCode.Text = ""
            txtConsolidateQty.Text = ""
            txtConsolidateSchSearch.Text = ""
            txtCustItemCode.Text = ""
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

#End Region
End Class