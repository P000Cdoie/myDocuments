'--------------------------------------------------------------------------------------------------
'COPYRIGHT      :   MIND
'CREATED BY     :   VINOD SINGH
'CREATED DATE   :   19/02/2018
'SCREEN         :   DISPATCH ADVICE AGAINST PENDING SCHEDULE
'PURPOSE        :   GENERATES PICK LIST AGAINST ITEM-CUSTOMER-PENDING SCHEDULE
'ISSUE ID       :   101462526 
'--------------------------------------------------------------------------------------------------
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Linq
Imports System.Collections.Generic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared


Public Class FRMMKTTRN0104

    Dim mintFormIndex As Integer
    Dim blnCheckPackQtymultiple_onDA As Boolean = False

    Private Enum enmCol
        CustCode = 1
        CustName
        DrgNo
        DrgDesc
        ForecastQty
        BackLogQty
        PendingDSAsOnDate
        AllowedDispatchQty
        For_The_Month_Schedule
        NextDaySch
        DispatchQty
        PkgStdQty
        PercAsPerBSR
        PickListNo
        PendingPicklistQty
        NeedByDate
        FIFOQty
    End Enum
    
#Region "Form level Events"
    Private Sub form_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub form_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Tag) = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub form_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            Me.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))            
        End Try
    End Sub
    Private Sub form_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
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
    Private Sub form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Call FitToClient(Me, GrpMain, ctlHeader, btnGrp, 500)
            Me.MdiParent = mdifrmMain
            mintFormIndex = mdifrmMain.AddFormNameToWindowList(Me.ctlHeader.Tag)
            SetGridsHeader()
            FillItemSearchCategory()
            btnGrp.ShowButtons(True, False, True, False)
            EnableControls(False, Me, True)
            txtDocNo.Enabled = True
            cmdDocHelp.Enabled = True
            sprPendingSchedule.Enabled = True
            txtDocNo.BackColor = Color.White
            txtDocNo.Focus()
            sprPendingSchedule.EditModePermanent = True
            sprPendingSchedule.EditModeReplace = True
            blnCheckPackQtymultiple_onDA = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("SELECT isnull(CheckPackQtymultiple_onDA,0) FROM sales_parameter WHERE UNIT_CODE ='" & gstrUNITID & "'"))

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub
#End Region

#Region "Methods"
    Private Sub SetGridsHeader()
        Try
            With sprPendingSchedule
                .MaxRows = 0
                .MaxCols = [Enum].GetNames(GetType(enmCol)).Count
                .Row = 0
                .set_RowHeight(0, 25)
                .Col = 0 : .set_ColWidth(0, 3)
                .Col = enmCol.PickListNo : .Text = "Pick List No" : .set_ColWidth(enmCol.PickListNo, 8) : .ColHidden = True
                .Col = enmCol.CustCode : .Text = "Customer Code" : .set_ColWidth(enmCol.CustCode, 8)
                .Col = enmCol.CustName : .Text = "Customer Name" : .set_ColWidth(enmCol.CustName, 30)
                .Col = enmCol.DrgNo : .Text = "Cust. Part No." : .set_ColWidth(enmCol.DrgNo, 12)
                .Col = enmCol.DrgDesc : .Text = "Cust.Part Desc." : .set_ColWidth(enmCol.DrgDesc, 15)
                .Col = enmCol.ForecastQty : .Text = "Forecast Qty." : .set_ColWidth(enmCol.ForecastQty, 8)
                .Col = enmCol.BackLogQty : .Text = "Backlog" : .set_ColWidth(enmCol.BackLogQty, 8)
                .Col = enmCol.PendingDSAsOnDate : .Text = "Pending DS As On Date + 3 Days" : .set_ColWidth(enmCol.PendingDSAsOnDate, 12)
                .Col = enmCol.NextDaySch : .Text = "Next Date Schedule" : .set_ColWidth(enmCol.NextDaySch, 10) : .ColHidden = True
                .Col = enmCol.PkgStdQty : .Text = "Std.Pack Qty." : .set_ColWidth(enmCol.PkgStdQty, 6)
                .Col = enmCol.DispatchQty : .Text = "Dispatch Qty." : .set_ColWidth(enmCol.DispatchQty, 12)
                .Col = enmCol.PercAsPerBSR : .Text = "% As Per BSR Stock" : .set_ColWidth(enmCol.PercAsPerBSR, 10) : .ColHidden = True
                .Col = enmCol.AllowedDispatchQty : .Text = "Total Allowed Dispatch Qty" : .set_ColWidth(enmCol.AllowedDispatchQty, 10)
                .Col = enmCol.For_The_Month_Schedule : .Text = "For the Month Schedule" : .set_ColWidth(enmCol.For_The_Month_Schedule, 10)
                .Col = enmCol.NeedByDate : .Text = "Need By Date" : .set_ColWidth(enmCol.NeedByDate, 8)
                .Col = enmCol.FIFOQty : .Text = "FIFO Qty" : .set_ColWidth(enmCol.FIFOQty, 8)
                .Col = enmCol.PendingPicklistQty : .Text = "Pending PickList Qty." : .set_ColWidth(enmCol.PendingPicklistQty, 8)

            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub AddBlankRow()
        Try
            With Me.sprPendingSchedule
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .BackColorStyle = FPSpreadADO.BackColorStyleConstants.BackColorStyleUnderGrid
                .set_RowHeight(.Row, 15)
                .Col = enmCol.PickListNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .Col = enmCol.CustCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .Col = enmCol.CustName : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .Col = enmCol.DrgNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .Col = enmCol.DrgDesc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .Col = enmCol.ForecastQty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Col = enmCol.BackLogQty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Col = enmCol.PendingDSAsOnDate : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Col = enmCol.NextDaySch : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .ColHidden = True
                .Col = enmCol.PkgStdQty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Col = enmCol.DispatchQty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMax = 999999999.99 : .TypeFloatMin = 0
                .Col = enmCol.PercAsPerBSR : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .ColHidden = True
                .Col = enmCol.AllowedDispatchQty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Col = enmCol.For_The_Month_Schedule : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Col = enmCol.NeedByDate : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .Col = enmCol.FIFOQty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Col = enmCol.PendingPicklistQty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight

            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub FillItemSearchCategory()
        Try
            With cboSearch
                .DataSource = Nothing
                .Items.Clear()
                .DataSource = [Enum].GetNames(GetType(enmCol))
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function ValidateData() As Boolean
        Dim dblBSRStock As Double
        Dim dblDispatchQty As Double
        Dim dblPendingPickList As Double
        Dim dblTotalDispatch As Double
        Dim dblAllowedDispatch As Double
        Dim dblPackQty As Double
        Dim Qty As Integer
        Dim dblActualDispatchQty As Double = 0
        Try
            If txtItemCode.Text.Trim = "" Then
                MsgBox("Kindly select Item Code First !", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return False
            ElseIf Val(lblBSRStock.Text) <= 0 Then
                MsgBox("BSR Stock must be greater than zero !", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return False
            ElseIf sprPendingSchedule.MaxRows = 0 Then
                MsgBox("No Pending Schedule found to save !", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return False
            End If
            dblBSRStock = GetStockInHand("01B1", txtItemCode.Text.Trim)
            lblBSRStock.Text = dblBSRStock
            With sprPendingSchedule
                For intRow As Integer = 1 To .MaxRows
                    .Row = intRow
                    .Col = enmCol.DispatchQty
                    dblDispatchQty += Val(.Value)
                    dblAllowedDispatch += Val(.Value)
                    dblActualDispatchQty = Val(.Value)
                    .Col = enmCol.PendingPicklistQty
                    dblPendingPickList += Val(.Value)
                    If dblAllowedDispatch < 0 Then
                        MsgBox("Dispatch Qty should be greater than 0 !", MsgBoxStyle.Exclamation, ResolveResString(100))
                        Return False
                    End If
                    If blnCheckPackQtymultiple_onDA Then
                        If dblActualDispatchQty > 0 Then
                            .Col = enmCol.PkgStdQty
                            dblPackQty = Val(.Value)
                            Qty = dblActualDispatchQty Mod dblPackQty

                            If Qty > 0 Then
                                MsgBox("Dispatch Qty should be multiple of Pack Qty !", MsgBoxStyle.Exclamation, ResolveResString(100))
                                Return False
                            End If

                        End If
                    End If
                   
                Next
                dblTotalDispatch = dblDispatchQty + dblPendingPickList
            End With
            Dim strDispatch As String = dblDispatchQty
            Dim strPendingPickList As String = dblPendingPickList
            Dim strStock As String = dblBSRStock
            'If Math.Round(dblBSRStock, 4) < Math.Round(dblTotalDispatch) Then
            '    MsgBox("Total Dispatch Qty - " & strDispatch & " " + Environment.NewLine + "Total Pending PickList Qty - " & strPendingPickList & "  " + Environment.NewLine + "Total New Stock - " & strStock & "  " + Environment.NewLine + "Dispatch + Pending picklist Qty must be equal or less than BSR stock !", MsgBoxStyle.Exclamation, ResolveResString(100))
            '    'MsgBox("Total Dispatch Qty - " & strDispatch & "   Total Pending PickList Qty - " & strPendingPickList & "            Total Stock - " & strStock & "    Dispatch + Pending picklist Qty must be equal or less than BSR stock !", MsgBoxStyle.Exclamation, ResolveResString(100))
            '    sprPendingSchedule.Row = sprPendingSchedule.MaxRows : sprPendingSchedule.Col = enmCol.DispatchQty
            '    sprPendingSchedule.Action = FPSpreadADO.ActionConstants.ActionActiveCell
            '    sprPendingSchedule.Focus()
            '    Return False
            'End If
            If Math.Round(dblDispatchQty, 4) = 0 Then
                MsgBox("Dispatch qty is 0.You cannot create dispatch.", MsgBoxStyle.Information, ResolveResString(100))
                sprPendingSchedule.Row = sprPendingSchedule.MaxRows : sprPendingSchedule.Col = enmCol.DispatchQty
                sprPendingSchedule.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                sprPendingSchedule.Focus()
                Return False
            End If
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function ValidateStock() As Boolean
        Dim dblBSRStock As Double
        Dim dblnetBSRStock As Double
        Dim dblDispatchQty As Double
        Dim dblPendingPickList As Double
        Dim dblTotalDispatch As Double
        Try
            
            dblnetBSRStock = lblNetStock.Text
            With sprPendingSchedule
                For intRow As Integer = 1 To .MaxRows
                    .Row = intRow
                    .Col = enmCol.DispatchQty
                    dblDispatchQty += Val(.Value)
                   
                Next
                dblTotalDispatch = dblDispatchQty
            End With
            Dim strDispatch As String = dblTotalDispatch
            Dim strNetStock As String = dblnetBSRStock
            If Math.Round(dblnetBSRStock, 4) < Math.Round(dblTotalDispatch) Then
                MsgBox("Total Dispatch Qty - " & strDispatch & " " + Environment.NewLine + "Total net Stock - " & strNetStock & "  " + Environment.NewLine + "Dispatch Qty must be equal or less than Net stock !", MsgBoxStyle.Exclamation, ResolveResString(100))
                'MsgBox("Total Dispatch Qty - " & strDispatch & "   Total Pending PickList Qty - " & strPendingPickList & "            Total Stock - " & strStock & "    Dispatch + Pending picklist Qty must be equal or less than BSR stock !", MsgBoxStyle.Exclamation, ResolveResString(100))
                sprPendingSchedule.Row = sprPendingSchedule.MaxRows : sprPendingSchedule.Col = enmCol.DispatchQty
                sprPendingSchedule.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                sprPendingSchedule.Focus()
                Return False
            End If
           
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub FillDispatchAdviceDetails()
        Try
            sprPendingSchedule.MaxRows = 0
            With sprPendingSchedule
                .Row = 0
                .Col = enmCol.BackLogQty : .ColHidden = True
                .Col = enmCol.ForecastQty : .ColHidden = True
                .Col = enmCol.NextDaySch : .ColHidden = True
                .Col = enmCol.PendingDSAsOnDate : .ColHidden = True
                .Col = enmCol.PercAsPerBSR : .ColHidden = True
                .Col = enmCol.NeedByDate : .ColHidden = True
                .Col = enmCol.FIFOQty : .ColHidden = True
                .Col = enmCol.PendingPicklistQty : .ColHidden = True
                .Col = enmCol.AllowedDispatchQty : .ColHidden = True
                .Col = enmCol.For_The_Month_Schedule : .ColHidden = True
            End With

            Using sqlCmd As SqlCommand = New SqlCommand
                With sqlCmd
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 0
                    .CommandText = "USP_GET_DISP_ADVICE_DTL_FOR_PENDING_SCH"
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@DISP_ADVICE_NO", SqlDbType.Int).Value = Val(txtDocNo.Text)
                    Using dt As DataTable = SqlConnectionclass.GetDataTable(sqlCmd)
                        If dt.Rows.Count > 0 Then
                            txtItemCode.Text = Convert.ToString(dt.Rows(0)("ITEM_CODE"))
                            lblItemDesc.Text = Convert.ToString(dt.Rows(0)("DESCRIPTION"))
                            lblBSRStock.Text = Convert.ToString(dt.Rows(0)("BSR_STOCK"))
                            'lblDocDate.Text = dt.Rows(0)("ADVICE_DATE").ToString(gstrDateFormat)
                            chkClosed.Checked = Convert.ToBoolean(dt.Rows(0)("PICKLIST_CLOSED"))
                            With sprPendingSchedule
                                For Each row As DataRow In dt.Rows
                                    AddBlankRow()
                                    .Row = .MaxRows
                                    .Col = enmCol.PickListNo : .Text = Convert.ToString(row("PICKLIST_NO"))
                                    .Col = enmCol.CustCode : .Text = Convert.ToString(row("CUST_CODE"))
                                    .Col = enmCol.CustName : .Text = Convert.ToString(row("CUST_NAME"))
                                    .Col = enmCol.DrgNo : .Text = Convert.ToString(row("CUST_DRG_NO"))
                                    .Col = enmCol.DrgDesc : .Text = Convert.ToString(row("DRG_DESC"))
                                    .Col = enmCol.DispatchQty : .Text = Convert.ToString(row("DISPATCH_QTY"))
                                    .Col = enmCol.PkgStdQty : .Text = Convert.ToString(row("STD_PKG_QTY"))
                                Next
                                .Row = 1 : .Row2 = .MaxRows
                                .Col = 1 : .Col2 = .MaxCols
                                .BlockMode = True
                                .Lock = True
                                .BlockMode = False
                            End With
                        Else
                            MsgBox("No Record Found !", MsgBoxStyle.Exclamation, ResolveResString(100))
                            Return
                        End If
                    End Using
                End With
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub RefreshScreen()
        Try
            EnableControls(False, Me, True)
            txtDocNo.Enabled = True
            cmdDocHelp.Enabled = True
            txtDocNo.BackColor = Color.White
            sprPendingSchedule.MaxRows = 0
            sprPendingSchedule.Enabled = True
            txtDocNo.Enabled = True
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function IsDispatchedToday(ByVal strItemCode As String) As Boolean
        Try
            Dim strSQL As String = String.Empty
            strSQL = "SELECT TOP 1 1 FROM VW_BSR_CURRENT_DAY_DISPATCH WHERE UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE='" & strItemCode.Trim & "'"
            Return IsRecordExists(strSQL)
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Control's Events" 

    Private Sub cmdDocHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDocHelp.Click

        Dim strDocNo() As String
        Dim strQry As String
        ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
        Try
            If btnGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                strQry = "SELECT [ADVICE_NO],[ADVICE_DATE],[ITEM_CODE],[DESCRIPTION] FROM VW_DISPATCH_ADVICE_FOR_PENDING_SCHEDULE WHERE [UNIT_CODE]='" & gstrUNITID & "' ORDER BY [ADVICE_NO] DESC,[ADVICE_DATE] DESC"
                strDocNo = ctlHelp1.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
                If IsNothing(strDocNo) = True Then Exit Sub
                If strDocNo.GetUpperBound(0) <> -1 Then
                    If (Len(strDocNo(0)) >= 1) And strDocNo(0) = "0" Then
                        MsgBox("No Record found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                        Exit Sub
                    Else
                        EnableControls(False, Me, True)
                        txtDocNo.Enabled = True
                        txtDocNo.BackColor = Color.White
                        cmdDocHelp.Enabled = True
                        txtDocNo.Text = strDocNo(0)
                        lblDocDate.Text = strDocNo(1).ToString()
                        txtItemCode.Text = strDocNo(2)
                        lblItemDesc.Text = strDocNo(3)
                        lblBSRStock.Text = "0.00"
                        FillDispatchAdviceDetails()
                        cboSearch.Enabled = True
                        txtSearch.Enabled = True
                        txtSearch.BackColor = Color.White
                        sprPendingSchedule.Enabled = True
                        With sprPendingSchedule
                            .Row = 1 : .Row2 = .MaxRows
                            .Col = enmCol.DispatchQty : .Col2 = enmCol.DispatchQty
                            .BlockMode = True
                            .Lock = True
                            .BlockMode = False
                        End With
                        txtDocNo.Focus()
                    End If

                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub cmdItemHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdItemHelp.Click
        Dim strDocNo() As String
        Dim strQry As String
        ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
        Try
            If btnGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                Using sqlCmd As SqlCommand = New SqlCommand
                    With sqlCmd
                        .CommandType = CommandType.StoredProcedure
                        .CommandText = "USP_PENDING_SCHEDULE_ASONDATE_FOR_DISP_ADVICE_NEW"
                        .CommandTimeout = 0
                        .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUnitId
                        .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                        SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                    End With
                End Using
                txtItemCode.BackColor = Color.White
                txtItemCode.ForeColor = Color.Black
                strQry = "SELECT A.CUST_DRGNO,A.ITEM_CODE,A.DESCRIPTION,A.BSR_STOCK FROM VW_DISP_ADVICE_ITEM_HELP A INNER JOIN TEMP_DISPATCH_ADVICE_NEXT_ITEM T ON T.UNIT_CODE=A.UNIT_CODE AND T.ITEM_CODE=A.ITEM_CODE AND T.CUST_DRGNO=A.CUST_DRGNO WHERE A.UNIT_CODE ='" & gstrUnitId & "' AND T.UNIT_CODE ='" & gstrUnitId & "' AND T.IP_ADDRESS='" & gstrIpaddressWinSck & "' GROUP BY A.ITEM_CODE,A.DESCRIPTION,A.BSR_STOCK,A.CUST_DRGNO ORDER BY A.CUST_DRGNO,A.ITEM_CODE "
                strDocNo = ctlHelp1.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
                If IsNothing(strDocNo) = True Then Exit Sub
                If strDocNo.GetUpperBound(0) <> -1 Then
                    If (Len(strDocNo(0)) >= 1) And strDocNo(0) = "0" Then
                        MsgBox("No Record found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                        Exit Sub
                    Else
                        txtItemCode.Text = strDocNo(1)
                        lblItemDesc.Text = strDocNo(2)
                        lblBSRStock.Text = Val(strDocNo(3))
                        'SendKeys.Send("{Tab}")
                        cmdShowSchedule.PerformClick()
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub cmdShowSchedule_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdShowSchedule.Click
        Try
            If btnGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                Dim dblPendSchAsOnDate As Double = 0
                Dim dblBackLogQty As Double = 0
                Dim dblNextDaySchedule As Double = 0
                Dim dblTotalAllowedDispatch As Double = 0
                With sprPendingSchedule
                    .Row = 0
                    .Col = enmCol.BackLogQty : .ColHidden = False
                    .Col = enmCol.ForecastQty : .ColHidden = False
                    .Col = enmCol.NextDaySch : .ColHidden = False
                    .Col = enmCol.PendingDSAsOnDate : .ColHidden = False
                    .Col = enmCol.PercAsPerBSR : .ColHidden = False
                    .Col = enmCol.AllowedDispatchQty : .ColHidden = False
                End With
                ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
                sprPendingSchedule.MaxRows = 0
                txtAccumuDispQty.Text = ""
                txtRemanDisQty.Text = ""
                If Not IsRecordExists("SELECT TOP 1 1 FROM ITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE ='" & txtItemCode.Text.Trim & "'") Then
                    MsgBox("Invalid Item Code !", MsgBoxStyle.Exclamation, ResolveResString(100))
                    txtItemCode.Text = String.Empty
                    txtItemCode.Focus()
                    Return
                End If

                If txtItemCode.Text.Trim.Length > 0 Then
                    lblBSRStock.Text = GetStockInHand("01B1", txtItemCode.Text.Trim)
                End If
                If txtItemCode.Text.Trim.Length = 0 Then
                    MsgBox("Kindly select Item Code first !", MsgBoxStyle.Exclamation, ResolveResString(100))
                    txtItemCode.Text = String.Empty
                    txtItemCode.Focus()
                    Return
                ElseIf Val(lblBSRStock.Text) <= 0 Then
                    'MsgBox("BSR stock must be greater than zero !", MsgBoxStyle.Exclamation, ResolveResString(100))
                    Dim strSQL = "Update TEMP_DISPATCH_ADVICE_NEXT_ITEM set ItemStatus=1 WHERE item_code='" & txtItemCode.Text & "' and UNIT_CODE='" & gstrUNITID & "' AND IP_ADDRESS='" & gstrIpaddressWinSck & "'"
                    SqlConnectionclass.ExecuteNonQuery(strSQL)
                    AutoNewMode()
                    Return
                End If
                If String.IsNullOrEmpty(lblItemDesc.Text.Trim) Then
                    lblItemDesc.Text = SqlConnectionclass.ExecuteScalar("SELECT DESCRIPTION FROM ITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE ='" & txtItemCode.Text.Trim & "'")
                End If

                If IsDispatchedToday(txtItemCode.Text.Trim) = True Then
                    txtItemCode.BackColor = Color.Green
                    txtItemCode.ForeColor = Color.White
                Else
                    txtItemCode.BackColor = Color.White
                    txtItemCode.ForeColor = Color.Black
                End If

                Using sqlCmd As SqlCommand = New SqlCommand
                    With sqlCmd
                        .CommandType = CommandType.StoredProcedure
                        .CommandText = "USP_PENDING_SCHEDULE_ASONDATE_FOR_DISP_ADVICE_NEW"
                        .CommandTimeout = 0
                        .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                        .Parameters.Add("@ITEM_CODE", SqlDbType.VarChar, 16).Value = txtItemCode.Text.Trim
                        .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                        SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                        Dim strSQL As String = "SELECT * FROM TMP_BSR_PENDINGSCHEDULEFORDISPADVICE WHERE UNIT_CODE ='" & gstrUNITID & "' AND IP_ADDRESS ='" & gstrIpaddressWinSck & "'"
                        Using dtSch As DataTable = SqlConnectionclass.GetDataTable(strSQL)
                            If dtSch.Rows.Count > 0 Then
                                'sprPendingSchedule.Row = 0 : sprPendingSchedule.Col = enmCol.PickListNo : sprPendingSchedule.ColHidden = True
                                For Each Row As DataRow In dtSch.Rows
                                    With sprPendingSchedule
                                        AddBlankRow()
                                        .Row = .MaxRows
                                        .Col = enmCol.CustCode : .Text = Convert.ToString(Row("CUSTOMER_CODE"))
                                        .Col = enmCol.CustName : .Text = Convert.ToString(Row("CUST_NAME"))
                                        .Col = enmCol.DrgNo : .Text = Convert.ToString(Row("CUST_DRGNO"))
                                        .Col = enmCol.DrgDesc : .Text = Convert.ToString(Row("DRG_DESC"))
                                        .Col = enmCol.ForecastQty : .Text = Convert.ToString(Row("FORECAST_QTY"))
                                        .Col = enmCol.BackLogQty : .Text = Convert.ToString(Row("BACKLOG_QTY"))
                                        dblBackLogQty = Val(.Text)
                                        .Col = enmCol.PendingDSAsOnDate : .Text = Convert.ToString(Row("PENDING_AS_ON_DATE"))
                                        dblPendSchAsOnDate = Val(.Text)
                                        .Col = enmCol.NextDaySch : .Text = Convert.ToString(Row("NEXT_DAY_SCHEDULE"))
                                        dblNextDaySchedule = Val(.Text)
                                        '.Col = enmCol.DispatchQty : .TypeFloatMin = 0 : .TypeFloatMax = (dblPendSchAsOnDate + dblBackLogQty + dblNextDaySchedule)
                                        .Col = enmCol.PkgStdQty : .Text = Convert.ToString(Row("STD_PKG_QTY"))
                                        .Col = enmCol.AllowedDispatchQty : .Text = Convert.ToString(Row("TotalAllowedDispatchasOnDate"))
                                        dblTotalAllowedDispatch += Val(.Text)
                                        .Col = enmCol.For_The_Month_Schedule : .Text = Convert.ToString(Row("FOR_THE_MONTH_SCHEDULE"))
                                        .Col = enmCol.NeedByDate
                                        If IsDBNull(Row("NEED_BY_DATE")) Then
                                            .Text = String.Empty
                                        Else
                                            .Text = Convert.ToDateTime(Row("NEED_BY_DATE")).ToString(gstrDateFormat)
                                        End If
                                        .Col = enmCol.FIFOQty : .Text = Convert.ToString(Row("FIFO_QTY"))
                                        .Col = enmCol.PercAsPerBSR : .Text = "0.00"
                                        .Col = enmCol.PendingPicklistQty : .Text = Convert.ToString(Row("PENDING_PICKLIST_QTY"))
                                        .BlockMode = True
                                        .Row = .MaxRows : .Row2 = .MaxRows
                                        .Col = enmCol.DispatchQty : .Col2 = enmCol.DispatchQty : .TypeFloatMin = 0 : .TypeFloatMax = Val(Row("TotalAllowedDispatchasOnDate"))
                                        .BlockMode = False
                                    End With
                                Next
                                If txtItemCode.Text.Trim.Length > 0 Then
                                    Dim dblStockQty As Int64 = SqlConnectionclass.ExecuteScalar("SELECT sum(PENDING_PICKLIST_QTY) FROM TMP_BSR_PENDINGSTOCKFORDISPADVICE WHERE UNIT_CODE ='" & gstrUNITID & "' AND IP_ADDRESS ='" & gstrIpaddressWinSck & "'")
                                    LblpendingStock.Text = dblStockQty
                                    lblNetStock.Text = lblBSRStock.Text - LblpendingStock.Text
                                End If
                                If dblTotalAllowedDispatch <= 0 Then
                                    MsgBox("Dispatch Advice has been created for pending schedule.You cannot create more dispatch advice.", MsgBoxStyle.Information, ResolveResString(100))
                                End If
                            Else
                                'MsgBox("No Delivery Schedule found !", MsgBoxStyle.Exclamation, ResolveResString(100))
                                strSQL = "Update TEMP_DISPATCH_ADVICE_NEXT_ITEM set ItemStatus=1 WHERE item_code='" & txtItemCode.Text & "' and UNIT_CODE='" & gstrUNITID & "' AND IP_ADDRESS='" & gstrIpaddressWinSck & "'"
                                SqlConnectionclass.ExecuteNonQuery(strSQL)
                                AutoNewMode()
                                'txtItemCode.Text = String.Empty
                                'txtItemCode.Focus()
                            End If
                        End Using
                    End With
                End Using
                With sprPendingSchedule
                    If .MaxRows > 0 Then
                        .Row = 1 : .Col = enmCol.DispatchQty
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        .Focus()
                    End If
                End With
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub txtSearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearch.TextChanged
        Dim intCounter As Integer
        Dim strText As String
        Dim Col As enmCol
        Try
            Col = DirectCast(Me.cboSearch.SelectedIndex + 1, enmCol)
            For intCounter = 1 To sprPendingSchedule.MaxRows
                With Me.sprPendingSchedule
                    .Row = intCounter : .Col = Col : strText = Trim(.Text)
                    If .FontBold = True Then
                        .FontBold = False
                        .Refresh()
                    End If
                End With
            Next
            If Len(txtSearch.Text) = 0 Then Exit Sub
            For intCounter = 1 To sprPendingSchedule.MaxRows
                With Me.sprPendingSchedule
                    .Row = intCounter : .Col = Col : strText = Trim(.Text)
                    If Trim(UCase(Mid(strText, 1, Len(txtSearch.Text)))) = Trim(UCase(txtSearch.Text)) Then
                        .Row = intCounter : .Col = Col : .FontBold = True
                        .Row = intCounter : .Col = Col : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        Exit For
                    End If
                End With
            Next
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnGrp_ButtonClick(ByVal Sender As System.Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles btnGrp.ButtonClick
        Dim strSQL As String = String.Empty
        Dim strCustCode As String = String.Empty
        Dim strDrgNo As String = String.Empty
        Dim strDrgDesc As String = String.Empty
        Dim dblForeCastQty, dblDispatchQty, dblPendingAsOnDate, _
            dblBackLogQty, dblNextDaySch, dblStdPkgQty, dblPercBSRStock As Double
        Dim intDispatchAdviceNo As Integer
        Dim dblForTheMonthSch As Double, dblFIFOQty As Double, dblAllowedbldispatchqty As Double
        Dim strNeedByDate As Object

        Try
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
            Select Case e.Button
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                    txtDocNo.Text = ""
                    lblDocDate.Text = ""
                    lblBSRStock.Text = "0"
                    LblpendingStock.Text = "0"
                    lblNetStock.Text = "0"
                    txtAccumuDispQty.Text = "0"
                    txtRemanDisQty.Text = "0"
                    txtItemCode.Text = ""
                    lblItemDesc.Text = ""
                    sprPendingSchedule.MaxRows = 0
                    EnableControls(True, Me, True)
                    txtDocNo.Enabled = False
                    cmdDocHelp.Enabled = False
                    cmdShowSchedule.Enabled = True
                    With sprPendingSchedule
                        .Row = 0
                        .Col = enmCol.BackLogQty : .ColHidden = False
                        .Col = enmCol.ForecastQty : .ColHidden = False
                        .Col = enmCol.NextDaySch : .ColHidden = False
                        .Col = enmCol.PendingDSAsOnDate : .ColHidden = False
                        .Col = enmCol.PercAsPerBSR : .ColHidden = False
                        .Col = enmCol.NeedByDate : .ColHidden = False
                        .Col = enmCol.FIFOQty : .ColHidden = False
                        .Col = enmCol.PendingPicklistQty : .ColHidden = False
                        .Col = enmCol.For_The_Month_Schedule : .ColHidden = False
                    End With
                    If txtItemCode.Enabled Then txtItemCode.Focus()

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                    If ValidateData() = False Then
                        Exit Sub
                    End If
                    If MsgBox("Do you want to save record(s)", MsgBoxStyle.Question + MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.No Then
                        sprPendingSchedule.Row = sprPendingSchedule.MaxRows : sprPendingSchedule.Col = enmCol.DispatchQty
                        sprPendingSchedule.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        sprPendingSchedule.Focus()
                        Return
                    End If

                    'Priti Starts for schedule shortage issue With out transaction.
                    Dim blnCheckSchedulOnSaveWOTrans As Boolean = False
                    blnCheckSchedulOnSaveWOTrans = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("SELECT isnull(CheckSchedulOnSaveWOTrans,0) FROM sales_parameter WHERE UNIT_CODE ='" & gstrUNITID & "'"))
                    If blnCheckSchedulOnSaveWOTrans Then

                        Using sqlCmd As SqlCommand = New SqlCommand
                            With sqlCmd
                                .CommandType = CommandType.StoredProcedure
                                .CommandText = "USP_PENDING_SCHEDULE_ASONDATE_FOR_DISP_ADVICE_NEW"
                                .CommandTimeout = 0
                                .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                                .Parameters.Add("@ITEM_CODE", SqlDbType.VarChar, 16).Value = txtItemCode.Text.Trim
                                .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                                SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                                'strSQL = SqlConnectionclass.ExecuteNonQuery("
                                Using dtSch As DataTable = SqlConnectionclass.GetDataTable("SELECT CUSTOMER_CODE, isnull(TotalAllowedDispatchasOnDate,0) as TotalAllowedDispatchasOnDate FROM TMP_BSR_PENDINGSTOCKFORDISPADVICE WHERE UNIT_CODE ='" & gstrUNITID & "' AND IP_ADDRESS ='" & gstrIpaddressWinSck & "'")
                                    If dtSch.Rows.Count > 0 Then
                                        For Each Row As DataRow In dtSch.Rows
                                            Dim strCust = Convert.ToString(Row("CUSTOMER_CODE"))
                                            Dim dblAllowDispatchQty As Double = Convert.ToDouble(Row("TotalAllowedDispatchasOnDate"))
                                            With sprPendingSchedule
                                                Dim InnerCust As String = ""
                                                Dim dblDispatch As Double = 0
                                                For intRow As Integer = 1 To .MaxRows
                                                    .Row = intRow
                                                    .Col = enmCol.CustCode
                                                    InnerCust = (.Text)
                                                    If UCase(Trim(strCust)) = UCase(Trim(InnerCust)) Then
                                                        .Row = intRow
                                                        .Col = enmCol.DispatchQty
                                                        dblDispatch = Val(.Value)
                                                        If dblAllowDispatchQty < dblDispatch Then
                                                            MsgBox("Only " & Convert.ToSingle(dblAllowDispatchQty) & " Schedule Qty is avaiable for Customer " & InnerCust & " .You cannot save Dispatch Advice !", MsgBoxStyle.Critical, ResolveResString(100))
                                                            Exit Sub
                                                        End If
                                                    End If

                                                Next
                                            End With

                                        Next
                                    End If

                                End Using
                            End With
                        End Using
                    End If
                    'Priti ends for schedule shortage issue.

                    'Priti Starts for Stock issue With out transaction.
                    Dim blnCheckStockOnSaveDispAdv As Boolean = False
                    blnCheckStockOnSaveDispAdv = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("SELECT isnull(CheckStockOnSaveDispAdv,0) FROM sales_parameter WHERE UNIT_CODE ='" & gstrUNITID & "'"))
                    If blnCheckStockOnSaveDispAdv Then
                        If txtItemCode.Text.Trim.Length > 0 Then
                            Dim dblStockQty As Int64 = SqlConnectionclass.ExecuteScalar("SELECT sum(PENDING_PICKLIST_QTY) FROM TMP_BSR_PENDINGSTOCKFORDISPADVICE WHERE UNIT_CODE ='" & gstrUNITID & "' AND IP_ADDRESS ='" & gstrIpaddressWinSck & "'")
                            LblpendingStock.Text = dblStockQty
                            lblNetStock.Text = lblBSRStock.Text - LblpendingStock.Text
                        End If
                        If ValidateStock() = False Then
                            Exit Sub
                        End If
                    End If

                    strSQL = "DELETE TMP_DISP_ADVICE_FOR_PENDING_SCH WHERE UNIT_CODE='" & gstrUNITID & "' AND IP_ADDRESS='" & gstrIpaddressWinSck & "'"
                    SqlConnectionclass.ExecuteNonQuery(strSQL)
                    SqlConnectionclass.BeginTrans()

                    'Priti Starts for schedule shortage issue With in transaction.
                    Dim dblCheckSchedulOnSave As Boolean = False
                    dblCheckSchedulOnSave = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("SELECT isnull(CheckSchedulOnSave,0) FROM sales_parameter WHERE UNIT_CODE ='" & gstrUNITID & "'"))
                    If dblCheckSchedulOnSave Then

                        Using sqlCmd As SqlCommand = New SqlCommand
                            With sqlCmd
                                .CommandType = CommandType.StoredProcedure
                                .CommandText = "USP_PENDING_SCHEDULE_ASONDATE_FOR_DISP_ADVICE_NEW"
                                .CommandTimeout = 0
                                .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                                .Parameters.Add("@ITEM_CODE", SqlDbType.VarChar, 16).Value = txtItemCode.Text.Trim
                                .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                                SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                                'strSQL = SqlConnectionclass.ExecuteNonQuery("
                                Using dtSch As DataTable = SqlConnectionclass.GetDataTable("SELECT CUSTOMER_CODE, isnull(TotalAllowedDispatchasOnDate,0) as TotalAllowedDispatchasOnDate FROM TMP_BSR_PENDINGSCHEDULEFORDISPADVICE WHERE UNIT_CODE ='" & gstrUNITID & "' AND IP_ADDRESS ='" & gstrIpaddressWinSck & "'")
                                    If dtSch.Rows.Count > 0 Then
                                        For Each Row As DataRow In dtSch.Rows
                                            Dim strCust = Convert.ToString(Row("CUSTOMER_CODE"))
                                            Dim dblAllowDispatchQty As Double = Convert.ToDouble(Row("TotalAllowedDispatchasOnDate"))
                                            With sprPendingSchedule
                                                Dim InnerCust As String = ""
                                                Dim dblDispatch As Double = 0
                                                For intRow As Integer = 1 To .MaxRows
                                                    .Row = intRow
                                                    .Col = enmCol.CustCode
                                                    InnerCust = (.Text)
                                                    If UCase(Trim(strCust)) = UCase(Trim(InnerCust)) Then
                                                        .Row = intRow
                                                        .Col = enmCol.DispatchQty
                                                        dblDispatch = Val(.Value)
                                                        If dblAllowDispatchQty < dblDispatch Then
                                                            SqlConnectionclass.RollbackTran()
                                                            MsgBox("Schedule is not avaiable.Unable to save Dispatch Advice for customer " & InnerCust & " !", MsgBoxStyle.Critical, ResolveResString(100))
                                                            Exit Sub
                                                        End If
                                                    End If

                                                Next
                                            End With

                                        Next
                                    End If

                                End Using
                            End With
                        End Using
                    End If
                    'Priti ends for schedule shortage issue.

                    With sprPendingSchedule
                        If .MaxRows > 0 Then
                            Using sqlcmd As SqlCommand = New SqlCommand
                                strSQL = "INSERT TMP_DISP_ADVICE_FOR_PENDING_SCH([UNIT_CODE],[IP_ADDRESS],[CUST_CODE],[CUST_DRG_NO],[FORECAST_QTY],[BACKLOG_QTY],[PENDING_DS_ASONDATE],[NEXT_DAY_SCH],[STD_PKG_QTY],[DISPATCH_QTY],FOR_THE_MONTH_SCH,NEED_BY_DATE,FIFO_QTY,TotalAllowedDispatchasOnDate)"
                                strSQL += " VALUES (@UNIT_CODE,@IP_ADDRESS,@CUST_CODE,@CUST_DRGNO,@FORECASTQTY,@BACKLOGQTY,@PENDINGASONDATE,@NEXTDATESCHEDULE,@PACKQTY,@DISPATCHQTY,@FORTHEMONTHSCHEDULE,@NEED_BY_DATE,@FIFOQTY,@TotalAllowedDispatchasOnDate)"
                                With sqlcmd
                                    .CommandText = strSQL
                                    .CommandType = CommandType.Text
                                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10)
                                    .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20)
                                    .Parameters.Add("@CUST_CODE", SqlDbType.VarChar, 10)
                                    .Parameters.Add("@CUST_DRGNO", SqlDbType.VarChar, 50)
                                    .Parameters.Add("@FORECASTQTY", SqlDbType.Money)
                                    .Parameters.Add("@BACKLOGQTY", SqlDbType.Money)
                                    .Parameters.Add("@PENDINGASONDATE", SqlDbType.Money)
                                    .Parameters.Add("@NEXTDATESCHEDULE", SqlDbType.Money)
                                    .Parameters.Add("@PACKQTY", SqlDbType.Money)
                                    .Parameters.Add("@DISPATCHQTY", SqlDbType.Money)
                                    .Parameters.Add("@FORTHEMONTHSCHEDULE", SqlDbType.Money)
                                    .Parameters.Add("@NEED_BY_DATE", SqlDbType.DateTime)
                                    .Parameters.Add("@FIFOQTY", SqlDbType.Money)
                                    .Parameters.Add("@TotalAllowedDispatchasOnDate", SqlDbType.Money)
                                End With

                                For intRow As Integer = 1 To .MaxRows
                                    .Row = intRow
                                    .Col = enmCol.DispatchQty
                                    If Val(.Text) > 0 Then
                                        dblDispatchQty = Val(.Text)
                                        .Col = enmCol.CustCode : strCustCode = .Text.Trim
                                        .Col = enmCol.DrgNo : strDrgNo = .Text.Trim
                                        .Col = enmCol.ForecastQty : dblForeCastQty = Val(.Text)
                                        .Col = enmCol.BackLogQty : dblBackLogQty = Val(.Text)
                                        .Col = enmCol.PendingDSAsOnDate : dblPendingAsOnDate = Val(.Text)
                                        .Col = enmCol.NextDaySch : dblNextDaySch = Val(.Text)
                                        .Col = enmCol.PercAsPerBSR : dblPercBSRStock = Val(.Text)
                                        .Col = enmCol.PkgStdQty : dblStdPkgQty = Val(.Text)
                                        .Col = enmCol.AllowedDispatchQty : dblAllowedbldispatchqty = Val(.Text)
                                        .Col = enmCol.For_The_Month_Schedule : dblForTheMonthSch = Val(.Text)
                                        .Col = enmCol.FIFOQty : dblFIFOQty = Val(.Text)

                                        .Col = enmCol.NeedByDate
                                        If String.IsNullOrEmpty(.Text) Then
                                            strNeedByDate = DBNull.Value
                                        Else
                                            strNeedByDate = getDateForDB(strNeedByDate)
                                        End If

                                        With sqlcmd
                                            .Parameters("@UNIT_CODE").Value = gstrUNITID
                                            .Parameters("@IP_ADDRESS").Value = gstrIpaddressWinSck
                                            .Parameters("@CUST_CODE").Value = strCustCode
                                            .Parameters("@CUST_DRGNO").Value = strDrgNo
                                            .Parameters("@FORECASTQTY").Value = dblForeCastQty
                                            .Parameters("@BACKLOGQTY").Value = dblBackLogQty
                                            .Parameters("@PENDINGASONDATE").Value = dblPendingAsOnDate
                                            .Parameters("@NEXTDATESCHEDULE").Value = dblNextDaySch
                                            .Parameters("@PACKQTY").Value = dblStdPkgQty
                                            .Parameters("@DISPATCHQTY").Value = dblDispatchQty
                                            .Parameters("@FORTHEMONTHSCHEDULE").Value = dblForTheMonthSch
                                            .Parameters("@NEED_BY_DATE").Value = strNeedByDate
                                            .Parameters("@FIFOQTY").Value = dblFIFOQty
                                            .Parameters("@TotalAllowedDispatchasOnDate").Value = dblAllowedbldispatchqty
                                            SqlConnectionclass.ExecuteNonQuery(sqlcmd)
                                        End With


                                    End If
                                Next
                            End Using
                            Using sqlCmd As SqlCommand = New SqlCommand
                                With sqlCmd
                                    .CommandText = "USP_SAVE_DISPATCH_ADVICE_FOR_PENDING_SCHEDULE"
                                    .CommandTimeout = 0
                                    .CommandType = CommandType.StoredProcedure
                                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                                    .Parameters.Add("@ITEM_CODE", SqlDbType.VarChar, 16).Value = txtItemCode.Text.Trim
                                    .Parameters.Add("@BSR_STOCK", SqlDbType.Money).Value = Val(lblBSRStock.Text)
                                    .Parameters.Add("@IP_ADDR", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                                    .Parameters.Add("@USER_ID", SqlDbType.VarChar, 20).Value = mP_User
                                    .Parameters.Add("@DISPATCH_ADVICE_NO", SqlDbType.Int).Direction = ParameterDirection.Output
                                    SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                                    intDispatchAdviceNo = .Parameters("@DISPATCH_ADVICE_NO").Value
                                    If intDispatchAdviceNo <= 0 Then
                                        SqlConnectionclass.RollbackTran()
                                        MsgBox("Unable to save Dispatch Advice !", MsgBoxStyle.Critical, ResolveResString(100))
                                        Exit Sub
                                    End If
                                End With
                            End Using
                        End If
                    End With
                    SqlConnectionclass.CommitTran()
                    MsgBox("Dispatch Advice Saved Successfully with Document No : " & intDispatchAdviceNo & "", MsgBoxStyle.Information, ResolveResString(100))
                    AutoNewMode()
                    Exit Sub

                    btnGrp.Revert()
                    RefreshScreen()
                    txtDocNo.Text = intDispatchAdviceNo
                    lblDocDate.Text = GetServerDate.ToString(gstrDateFormat)
                    If Val(txtDocNo.Text) > 0 Then
                        FillDispatchAdviceDetails()
                    End If
                    Exit Sub

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
                    If Val(txtDocNo.Text) = 0 Then
                        MsgBox("Kindly select a valid Dispatch Advice first !", MsgBoxStyle.Exclamation, ResolveResString(100))
                        Return
                    End If
                    If chkClosed.Checked = True Then
                        MsgBox("Dispatch Advice already has been closed or deleted !", MsgBoxStyle.Exclamation, ResolveResString(100))
                        Return
                    End If
                    If MsgBox("Are you sure to delete this Dispatch Advice ?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                        Dim dtItemCode As New DataTable
                        Using sqlcmd As SqlCommand = New SqlCommand
                            With sqlcmd
                                .CommandText = "USP_DELETE_DISP_ADVICE_FOR_PENDING_SCHEDULE"
                                .CommandTimeout = 0
                                .CommandType = CommandType.StoredProcedure
                                .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                                .Parameters.Add("@DISPATCH_ADVICE_NO", SqlDbType.Int).Value = Val(txtDocNo.Text)
                                .Parameters.Add("@USER_ID", SqlDbType.VarChar, 20).Value = mP_User
                                .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
                                dtItemCode.Load(SqlConnectionclass.ExecuteReader(sqlcmd))
                                If Convert.ToString(.Parameters("@MESSAGE").Value).Length > 0 Then
                                    MessageBox.Show(Convert.ToString(.Parameters("@MESSAGE").Value), "eMPro", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                                    Exit Sub
                                End If
                                'SqlConnectionclass.ExecuteNonQuery(sqlcmd)
                            End With
                        End Using
                        MsgBox("Dispatch Advice has been deleted successfully.", MsgBoxStyle.Information, ResolveResString(100))
                        btnGrp.Revert()
                        RefreshScreen()
                    End If
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                    If MsgBox("Do you want to cancel the current process?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, ResolveResString(1000)) = MsgBoxResult.Yes Then
                        btnGrp.Revert()
                        EnableControls(False, Me, True)
                        lblDocDate.Text = ""
                        lblBSRStock.Text = "0"
                        LblpendingStock.Text = "0"
                        lblNetStock.Text = "0"
                        txtAccumuDispQty.Text = "0"
                        txtRemanDisQty.Text = "0"
                        lblItemDesc.Text = ""
                        sprPendingSchedule.MaxRows = 0
                        txtDocNo.Enabled = True
                        txtDocNo.BackColor = Color.White
                        txtDocNo.Focus()
                        cmdDocHelp.Enabled = True
                    End If
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                    If MsgBox("Do you want to close this screen?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, ResolveResString(1000)) = MsgBoxResult.Yes Then
                        Me.Close()
                    End If
            End Select
        Catch ex As Exception
            SqlConnectionclass.RollbackTran()
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    
    Private Sub txtItemCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtItemCode.KeyUp
        Try
            If e.KeyCode = Keys.F1 Then
                cmdItemHelp.PerformClick()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtItemCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtItemCode.TextChanged
        Try
            lblBSRStock.Text = "0"
            LblpendingStock.Text = "0"
            lblNetStock.Text = "0"
            txtAccumuDispQty.Text = "0"
            txtRemanDisQty.Text = "0"
            lblItemDesc.Text = ""
            sprPendingSchedule.MaxRows = 0
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub sprPendingSchedule_CircularFormula(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_CircularFormulaEvent) Handles sprPendingSchedule.CircularFormula

    End Sub


    Private Sub sprPendingSchedule_ClickEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles sprPendingSchedule.ClickEvent
        Dim dblValue As Double
        Dim dblAllowedDispatchQty As Double
        Dim strCustomer As String
        Dim strItem As String
        Try
            With sprPendingSchedule
                If btnGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    If e.col = enmCol.DispatchQty Then
                        .Row = e.row
                        .Col = e.col
                        .Col = enmCol.AllowedDispatchQty : dblAllowedDispatchQty = .Text.Trim
                        .Col = enmCol.CustCode : strCustomer = .Text.Trim
                        If dblAllowedDispatchQty = 0 Then
                            strItem = txtItemCode.Text
                            Dim strQuery = SqlConnectionclass.ExecuteScalar("SELECT substring( stuff(( select  ' ' + convert(varchar(20),  b.DISP_ADVICE_NO)  + '  Qty - ' + convert(varchar(25), b.PICKLIST_QTY)+ ' ,  ' " & _
                                  "FROM TMP_BSR_PENDINGSCHEDULEFORDISPADVICE A " & _
                                  "LEFT OUTER JOIN DISP_ADVICE_PENDING_SCHEDULE_PICKLIST B " & _
                                  "ON	A.UNIT_CODE = B.UNIT_CODE " & _
                               "AND A.CUSTOMER_CODE = B.CUST_CODE " & _
                               "AND A.ITEM_CODE = B.ITEM_CODE " & _
                               "AND A.CUST_DRGNO = B.CUST_DRG_NO " & _
                               "AND B.CLOSED = 0 " & _
                                  "WHERE A.UNIT_CODE ='" & gstrUNITID & "'  and a.ITEM_CODE='" & strItem & "' and a.IP_ADDRESS='" & gstrIpaddressWinSck & "' and a.CUSTOMER_CODE='" & strCustomer & "'  for xml path ('')),1,1,'' ) ,0,len(stuff(( select  ' ' + convert(varchar(20),  b.DISP_ADVICE_NO)  + '  Qty - ' + convert(varchar(25), b.PICKLIST_QTY)+ ' ,  ' " & _
                                  "FROM TMP_BSR_PENDINGSCHEDULEFORDISPADVICE A " & _
                                     "LEFT OUTER JOIN DISP_ADVICE_PENDING_SCHEDULE_PICKLIST B " & _
                                  "ON	A.UNIT_CODE = B.UNIT_CODE " & _
                               "AND A.CUSTOMER_CODE = B.CUST_CODE " & _
                               "AND A.ITEM_CODE = B.ITEM_CODE " & _
                               "AND A.CUST_DRGNO = B.CUST_DRG_NO " & _
                               "AND B.CLOSED = 0 " & _
                                  "WHERE A.UNIT_CODE ='" & gstrUNITID & "'  and a.ITEM_CODE='" & strItem & "' and a.IP_ADDRESS='" & gstrIpaddressWinSck & "' and a.CUSTOMER_CODE='" & strCustomer & "'  for xml path ('')),1,1,'' )) -1 )")
                            MsgBox("PickList is already generated to the customer code " & strCustomer & " with avalaible BSR stock . " + Environment.NewLine + "Pick List No - " & strQuery & " ", MsgBoxStyle.Information, ResolveResString(100))
                        End If

                       
                    End If
                End If

            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub sprPendingSchedule_EditChange(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles sprPendingSchedule.EditChange
        Dim dblValue As Double
        Try
            With sprPendingSchedule
                If btnGrp.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    If e.col = enmCol.DispatchQty Then
                        With sprPendingSchedule

                            Dim dblDispatch As Double = 0
                            For intRow As Integer = 1 To .MaxRows
                                .Row = intRow
                                .Col = enmCol.DispatchQty

                                dblDispatch += Val(.Value)

                            Next
                            txtAccumuDispQty.Text = dblDispatch
                            txtRemanDisQty.Text = Val(lblNetStock.Text) - Val(txtAccumuDispQty.Text)
                        End With

                        .Row = e.row
                        .Col = e.col
                        If .Value > 0 And Val(lblBSRStock.Text) > 0 Then
                            dblValue = Val(.Value)
                            .Col = enmCol.PercAsPerBSR : .Text = Math.Round(dblValue / Val(lblBSRStock.Text) * 100, 4)

                        End If

                    End If
                End If

            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub txtDocNo_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtDocNo.KeyUp
        Try
            If e.KeyCode = Keys.F1 Then
                cmdDocHelp.PerformClick()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtItemCode_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtItemCode.KeyPress
        Dim keyasii As Integer = Asc(e.KeyChar)
        Try
            If keyasii = 13 Then
                If String.IsNullOrEmpty(txtItemCode.Text.Trim) = False Then
                    cmdShowSchedule.PerformClick()
                    With sprPendingSchedule
                        If .MaxRows > 0 Then
                            .Row = 1 : .Col = enmCol.DispatchQty
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                        End If
                    End With
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

#End Region
    Private Sub AutoNewMode()
        Dim dtItemCode As New DataTable
        Try
            Dim itemCode As String = Trim(txtItemCode.Text)
            txtDocNo.Text = ""
            lblDocDate.Text = ""
            lblBSRStock.Text = "0"
            LblpendingStock.Text = "0"
            lblNetStock.Text = "0"
            txtItemCode.Text = ""
            lblItemDesc.Text = ""
            txtAccumuDispQty.Text = "0"
            txtRemanDisQty.Text = "0"
            sprPendingSchedule.MaxRows = 0
            EnableControls(True, Me, True)
            txtDocNo.Enabled = False
            cmdDocHelp.Enabled = False
            cmdShowSchedule.Enabled = True
            With sprPendingSchedule
                .Row = 0
                .Col = enmCol.BackLogQty : .ColHidden = False
                .Col = enmCol.ForecastQty : .ColHidden = False
                .Col = enmCol.NextDaySch : .ColHidden = False
                .Col = enmCol.PendingDSAsOnDate : .ColHidden = False
                .Col = enmCol.PercAsPerBSR : .ColHidden = False
                .Col = enmCol.NeedByDate : .ColHidden = False
                .Col = enmCol.FIFOQty : .ColHidden = False
                .Col = enmCol.PendingPicklistQty : .ColHidden = False
                .Col = enmCol.For_The_Month_Schedule : .ColHidden = False
            End With
            Using sqlCmd As SqlCommand = New SqlCommand
                With sqlCmd
                    .CommandText = "USP_DISPATCH_ADVICE_GET_NEXT_ITEM"
                    .CommandTimeout = 0
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUnitId
                    .Parameters.Add("@ITEM_CODE", SqlDbType.VarChar, 16).Value = itemCode
                    .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 25).Value = gstrIpaddressWinSck
                    .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                    dtItemCode.Load(SqlConnectionclass.ExecuteReader(sqlCmd))
                    If Convert.ToString(.Parameters("@MESSAGE").Value).Length > 0 Then
                        MessageBox.Show(Convert.ToString(.Parameters("@MESSAGE").Value), "eMPro", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                        Exit Sub
                    End If
                End With
            End Using         
            If dtItemCode IsNot Nothing AndAlso dtItemCode.Rows.Count > 0 Then
                txtItemCode.Text = Convert.ToString(dtItemCode.Rows(0)("ITEM_CODE"))
                lblItemDesc.Text = Convert.ToString(dtItemCode.Rows(0)("DESCRIPTION"))
            Else
                Dim strSQL = "Update TEMP_DISPATCH_ADVICE_NEXT_ITEM set ItemStatus=0 WHERE  UNIT_CODE='" & gstrUNITID & "' AND IP_ADDRESS='" & gstrIpaddressWinSck & "'"
                SqlConnectionclass.ExecuteNonQuery(strSQL)
            End If
            If txtItemCode.Text.Length > 0 Then
                cmdShowSchedule.PerformClick()
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            dtItemCode.Dispose()
        End Try
    End Sub

    Private Sub sprPendingSchedule_KeyDownEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles sprPendingSchedule.KeyDownEvent
        Try
            If e.keyCode = Keys.Enter Then
                If sprPendingSchedule.ActiveCol = enmCol.DispatchQty Then
                    If sprPendingSchedule.ActiveRow = sprPendingSchedule.MaxRows Then
                        SendKeys.Send("{Tab}")
                    Else
                        sprPendingSchedule.Row = sprPendingSchedule.ActiveRow + 1 : sprPendingSchedule.Col = enmCol.DispatchQty
                        sprPendingSchedule.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        sprPendingSchedule.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
End Class