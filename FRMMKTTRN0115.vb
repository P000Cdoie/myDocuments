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


Public Class FRMMKTTRN0115
    Dim intSave As Integer = 0
    Dim mintFormIndex As Integer
    Dim dtAuto As DataTable
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
        Mode
    End Enum
    Private Enum enmColItem
        Status = 0
        ItemCode
        ItemDesc
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
            Call FitToClient(Me, GrpMain1, ctlHeader, btnGrp1, 500)
            Me.MdiParent = mdifrmMain
            mintFormIndex = mdifrmMain.AddFormNameToWindowList(Me.ctlHeader.Tag)
            SetGridsHeader()
            FillItemSearchCategory()
            FillCustomerSearchCategory()
            btnGrp1.ShowButtons(True, False, True, False)
            EnableControls(False, Me, True)
            txtDocNo.Enabled = True
            cmdDocHelp.Enabled = True
            sprPendingSchedule.Enabled = True
            txtDocNo.BackColor = Color.White
            txtDocNo.Focus()
            sprPendingSchedule.EditModePermanent = True
            sprPendingSchedule.EditModeReplace = True
            lblNote.Visible = False
            OptCurrent.Enabled = True
            OptFuture.Enabled = True
            OptCurrent.Checked = True
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
                .Col = enmCol.CustName : .Text = "Customer Name" : .set_ColWidth(enmCol.CustName, 25)
                .Col = enmCol.DrgNo : .Text = "Cust. Part No." : .set_ColWidth(enmCol.DrgNo, 8)
                .Col = enmCol.DrgDesc : .Text = "Cust.Part Desc." : .set_ColWidth(enmCol.DrgDesc, 10)
                .Col = enmCol.ForecastQty : .Text = "Forecast Qty." : .set_ColWidth(enmCol.ForecastQty, 8) : .ColHidden = True
                .Col = enmCol.BackLogQty : .Text = "Backlog" : .set_ColWidth(enmCol.BackLogQty, 8)
                .Col = enmCol.PendingDSAsOnDate : .Text = "Pending DS As On Date + 3 Days" : .set_ColWidth(enmCol.PendingDSAsOnDate, 8)
                .Col = enmCol.NextDaySch : .Text = "Next Date Schedule" : .set_ColWidth(enmCol.NextDaySch, 10) : .ColHidden = True
                .Col = enmCol.PkgStdQty : .Text = "Std.Pack Qty." : .set_ColWidth(enmCol.PkgStdQty, 6)
                .Col = enmCol.DispatchQty : .Text = "Dispatch Qty." : .set_ColWidth(enmCol.DispatchQty, 12)
                .Col = enmCol.PercAsPerBSR : .Text = "% As Per BSR Stock" : .set_ColWidth(enmCol.PercAsPerBSR, 10) : .ColHidden = True
                .Col = enmCol.AllowedDispatchQty : .Text = "Total Allowed Dispatch Qty" : .set_ColWidth(enmCol.AllowedDispatchQty, 10)
                .Col = enmCol.For_The_Month_Schedule : .Text = "For the Month Schedule" : .set_ColWidth(enmCol.For_The_Month_Schedule, 10)
                .Col = enmCol.NeedByDate : .Text = "Need By Date" : .set_ColWidth(enmCol.NeedByDate, 8)
                .Col = enmCol.FIFOQty : .Text = "FIFO Qty" : .set_ColWidth(enmCol.FIFOQty, 8)
                .Col = enmCol.PendingPicklistQty : .Text = "Pending PickList Qty." : .set_ColWidth(enmCol.PendingPicklistQty, 8)
                .Col = enmCol.Mode : .Text = "Mode" : .set_ColWidth(enmCol.PendingPicklistQty, 8)
                
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub AddBlankRow()
        Try
            Dim strComboLst As String = String.Empty
            strComboLst = "Road" & Chr(9) & "Air"
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
                .Col = enmCol.ForecastQty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .ColHidden = True
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
                .Col = enmCol.Mode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .TypeComboBoxList = strComboLst
                .TypeComboBoxEditable = 0
                .TypeComboBoxIndex = 1
                .Value = "Road"
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
        Dim dblActualDispatchQty As Double = 0
        Dim dblAllowedDispatch As Double = 0
        Dim dblPackQty As Double
        Dim Qty As Integer
        Try
            If sprPendingSchedule.MaxRows = 0 Then
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
            If btnGrp1.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                strQry = "SELECT [ADVICE_NO],[ADVICE_DATE],[ITEM_CODE],[DESCRIPTION] FROM VW_DISPATCH_ADVICE_FOR_PENDING_SCHEDULE WHERE [UNIT_CODE]='" & gstrUNITID & "' ORDER BY [ADVICE_NO] DESC,[ADVICE_DATE] DESC"
                strDocNo = ctlHelp11.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
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
            If btnGrp1.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                Dim strDSCriteria As String = ""
                If OptCurrent.Checked Then
                    strDSCriteria = "Current"
                Else
                    strDSCriteria = "Future"
                End If
                Using sqlCmd As SqlCommand = New SqlCommand
                    With sqlCmd
                        .CommandType = CommandType.StoredProcedure
                        .CommandText = "USP_PENDING_SCHEDULE_ASONDATE_FOR_DISP_ADVICE_NEWAuto"
                        .CommandTimeout = 0
                        .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                        .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                        .Parameters.Add("@DSCriteria", SqlDbType.VarChar, 20).Value = strDSCriteria

                        SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                    End With
                End Using
                txtItemCode.BackColor = Color.White
                txtItemCode.ForeColor = Color.Black
                strQry = "SELECT A.CUST_DRGNO,A.ITEM_CODE,A.DESCRIPTION,A.BSR_STOCK FROM VW_DISP_ADVICE_ITEM_HELP A INNER JOIN TEMP_DISPATCH_ADVICE_NEXT_ITEM T ON T.UNIT_CODE=A.UNIT_CODE AND T.ITEM_CODE=A.ITEM_CODE AND T.CUST_DRGNO=A.CUST_DRGNO WHERE A.UNIT_CODE ='" & gstrUNITID & "' AND T.UNIT_CODE ='" & gstrUNITID & "' AND T.IP_ADDRESS='" & gstrIpaddressWinSck & "' GROUP BY A.ITEM_CODE,A.DESCRIPTION,A.BSR_STOCK,A.CUST_DRGNO ORDER BY A.CUST_DRGNO,A.ITEM_CODE "
                strDocNo = ctlHelp11.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
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
                        Dim strText As String = ""
                        Dim intCounter As Integer = 0
                        For intCounter = 0 To dgvItemDetail.Rows.Count - 1
                            strText = dgvItemDetail.Rows(intCounter).Cells(enmColItem.ItemCode).Value
                            If Trim(UCase(Mid(strText, 1, Len(txtItemCode.Text)))) = Trim(UCase(txtItemCode.Text)) Then
                                dgvItemDetail.CurrentCell = dgvItemDetail.Rows(intCounter).Cells(enmColItem.ItemCode)
                                Exit For
                            End If
                        Next
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
            If btnGrp1.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                Dim dblPendSchAsOnDate As Double = 0
                Dim dblBackLogQty As Double = 0
                Dim dblNextDaySchedule As Double = 0
                Dim dblTotalAllowedDispatch As Double = 0
                With sprPendingSchedule
                    .Row = 0
                    .Col = enmCol.BackLogQty : .ColHidden = False
                    .Col = enmCol.ForecastQty : .ColHidden = True
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
                Dim strDSCriteria As String = ""
                If OptCurrent.Checked Then
                    strDSCriteria = "Current"
                Else
                    strDSCriteria = "Future"
                End If
                Using sqlCmd As SqlCommand = New SqlCommand
                    With sqlCmd
                        .CommandType = CommandType.StoredProcedure
                        .CommandText = "USP_PENDING_SCHEDULE_ASONDATE_FOR_DISP_ADVICE_NEWAuto"
                        .CommandTimeout = 0
                        .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                        .Parameters.Add("@ITEM_CODE", SqlDbType.VarChar, 16).Value = txtItemCode.Text.Trim
                        .Parameters.Add("@DSCriteria", SqlDbType.VarChar, 20).Value = strDSCriteria

                        If chkOptimized.Checked = True Then
                            .Parameters.Add("@EntryType", SqlDbType.VarChar, 10).Value = "Auto"
                        End If

                        .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                        SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                        Dim strSQL As String = "SELECT * FROM TMP_BSR_PENDINGSCHEDULEFORDISPADVICE WHERE UNIT_CODE ='" & gstrUNITID & "' AND IP_ADDRESS ='" & gstrIpaddressWinSck & "'"
                        Using dtSch As DataTable = SqlConnectionclass.GetDataTable(strSQL)
                            If dtSch.Rows.Count > 0 Then
                                'sprPendingSchedule.Row = 0 : sprPendingSchedule.Col = enmCol.PickListNo : sprPendingSchedule.ColHidden = True
                                Dim dblNetStock As Integer = 0
                                Dim dblPendingNetStock As Integer = 0
                                If txtItemCode.Text.Trim.Length > 0 Then
                                    Dim dblStockQty As Int64 = SqlConnectionclass.ExecuteScalar("SELECT sum(PENDING_PICKLIST_QTY) FROM TMP_BSR_PENDINGSTOCKFORDISPADVICE WHERE UNIT_CODE ='" & gstrUNITID & "' AND IP_ADDRESS ='" & gstrIpaddressWinSck & "'")
                                    LblpendingStock.Text = dblStockQty
                                    lblNetStock.Text = lblBSRStock.Text - LblpendingStock.Text
                                    dblNetStock = lblNetStock.Text
                                    dblPendingNetStock = dblNetStock
                                End If
                                Dim dblDispatchQty As Double = 0
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
                                        dblTotalAllowedDispatch = Val(.Text)
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
                                        .Col = enmCol.Mode : .Text = "Road"
                                        If dblPendingNetStock > 0 Then
                                            If dblTotalAllowedDispatch < dblPendingNetStock Then
                                                .Col = enmCol.DispatchQty : .Text = Convert.ToString(Row("TotalAllowedDispatchasOnDate"))
                                                dblPendingNetStock = dblPendingNetStock - dblTotalAllowedDispatch
                                            Else
                                                .Col = enmCol.DispatchQty : .Text = dblPendingNetStock
                                                dblPendingNetStock = 0
                                            End If
                                        End If
                                        .Col = enmCol.DispatchQty
                                        dblDispatchQty += Val(.Text)
                                        txtAccumuDispQty.Text = dblDispatchQty


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
    Private Sub SetItemGridsHeader()
        Try
            
            dgvItemDetail.Columns.Clear()

            Dim objChkBox As New DataGridViewCheckBoxColumn
            objChkBox.Name = "Selection"
            objChkBox.HeaderText = "Select"

            dgvItemDetail.Columns.Add(objChkBox)
            dgvItemDetail.Columns.Add("ItemCode", "Item Code")
            dgvItemDetail.Columns.Add("ItemDesc", "Item Desc")


            dgvItemDetail.Columns(enmColItem.Status).Visible = False
            dgvItemDetail.Columns(enmColItem.Status).Width = 50
            dgvItemDetail.Columns(enmColItem.ItemCode).Width = 150
            dgvItemDetail.Columns(enmColItem.ItemDesc).Width = 300

            dgvItemDetail.Columns(enmColItem.Status).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvItemDetail.Columns(enmColItem.ItemCode).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvItemDetail.Columns(enmColItem.ItemDesc).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            
            dgvItemDetail.Columns(enmColItem.ItemCode).ReadOnly = True
            dgvItemDetail.Columns(enmColItem.ItemDesc).ReadOnly = True
            
            dgvItemDetail.Columns(enmColItem.Status).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvItemDetail.Columns(enmColItem.ItemCode).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvItemDetail.Columns(enmColItem.ItemDesc).SortMode = DataGridViewColumnSortMode.NotSortable

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
   
    Private Sub FillItemGrid()
        Dim ds As New DataSet
        Try
            SetItemGridsHeader()
            Dim i As Integer = 0
            Dim strDSCriteria As String = ""
            If OptCurrent.Checked Then
                strDSCriteria = "Current"
            Else
                strDSCriteria = "Future"
            End If
            Using sqlCmd As SqlCommand = New SqlCommand
                With sqlCmd
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "USP_PENDING_SCHEDULE_ASONDATE_FOR_DISP_ADVICE_NEWAuto"
                    '.CommandText = "USP_PENDING_SCHEDULE_ASONDATE_FOR_DISP_ADVICE_NEWpriti"
                    .CommandTimeout = 0
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                    .Parameters.Add("@DSCriteria", SqlDbType.VarChar, 20).Value = strDSCriteria

                    SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                End With
            End Using

            Using sqlCmd As SqlCommand = New SqlCommand
                With sqlCmd
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "DISPATCH_ADVICE_ITEMLIST"
                    .CommandTimeout = 0
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                    SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                End With
            End Using
         
            Dim strSQL As String = "SELECT * FROM DISPATCH_ADVICE_ITEM_LIST where UNIT_CODE ='" & gstrUNITID & "'  AND IP_ADDRESS='" & gstrIpaddressWinSck & "'"
            Dim dtSch As DataTable = SqlConnectionclass.GetDataTable(strSQL)
            If dtSch.Rows.Count > 0 Then

                dgvItemDetail.Rows.Clear()
                dgvItemDetail.Rows.Add(dtSch.Rows.Count)
                For Each dr As DataRow In dtSch.Rows
                    dgvItemDetail.Rows(i).Cells(enmColItem.Status).Value = False
                    dgvItemDetail.Rows(i).Cells(enmColItem.ItemCode).Value = dr("ITEM_CODE")
                    dgvItemDetail.Rows(i).Cells(enmColItem.ItemDesc).Value = dr("ITEM_DESC")
                    i += 1

                Next
            Else
                MsgBox("No Item found.", MsgBoxStyle.Information, ResolveResString(100))
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub
    Private Sub btnGrp_ButtonClick(ByVal Sender As System.Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles btnGrp1.ButtonClick
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
                    'EnableControls(True, Me, True)
                    txtDocNo.Enabled = False
                    cmdDocHelp.Enabled = False
                    cmdShowSchedule.Enabled = True
                    With sprPendingSchedule
                        .Row = 0
                        .Col = enmCol.BackLogQty : .ColHidden = False
                        .Col = enmCol.ForecastQty : .ColHidden = True
                        .Col = enmCol.NextDaySch : .ColHidden = False
                        .Col = enmCol.PendingDSAsOnDate : .ColHidden = False
                        .Col = enmCol.PercAsPerBSR : .ColHidden = False
                        .Col = enmCol.NeedByDate : .ColHidden = False
                        .Col = enmCol.FIFOQty : .ColHidden = False
                        .Col = enmCol.PendingPicklistQty : .ColHidden = False
                        .Col = enmCol.For_The_Month_Schedule : .ColHidden = False
                    End With
                    If txtItemCode.Enabled Then txtItemCode.Focus()
                    If MsgBox("Are your sure with selected DS schedule criteria ", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                        FillItemGrid()
                        chkOptimized.Checked = True
                        chkOptimized.Enabled = True
                    Else
                        btnGrp1.Revert()
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
                        OptCurrent.Enabled = True
                        OptFuture.Enabled = True
                        OptCurrent.Checked = True
                        dgvItemDetail.Columns.Clear()
                    End If


                  
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                    If MsgBox("Do you want to save record(s)", MsgBoxStyle.Question + MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then

                        AutoSaveData()
                    End If

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
                        btnGrp1.Revert()
                        RefreshScreen()
                    End If
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                    If MsgBox("Do you want to cancel the current process?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, ResolveResString(1000)) = MsgBoxResult.Yes Then
                        btnGrp1.Revert()
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
                        OptCurrent.Enabled = True
                        OptFuture.Enabled = True
                        OptCurrent.Checked = True
                        dgvItemDetail.Columns.Clear()
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
            lblItemDesc.Text = ""
            txtAccumuDispQty.Text = "0"
            txtRemanDisQty.Text = "0"
            sprPendingSchedule.MaxRows = 0
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub


    Private Sub sprPendingSchedule_ClickEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles sprPendingSchedule.ClickEvent
       
    End Sub

    Private Sub sprPendingSchedule_EditChange(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles sprPendingSchedule.EditChange
        Dim dblValue As Double
        Try
            With sprPendingSchedule
                If btnGrp1.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
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

  
    Private Sub GvItem_ClickEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles GvItem.ClickEvent
        Dim strStatus As String
        With GvItem
            If e.col = enmColItem.Status Then
                .Row = e.row
                .Col = e.col
                .Col = enmColItem.Status : strStatus = .Text.Trim
                If strStatus = 1 Then
                    MsgBox("click")
                End If
            End If
        End With
    End Sub

    Private Sub GvItem_EditChange(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles GvItem.EditChange
        Dim strStatus As String
        With GvItem
            If e.col = enmColItem.Status Then
                .Row = e.row
                .Col = e.col
                .Col = enmColItem.Status : strStatus = .Text.Trim
                If strStatus = 1 Then
                    MsgBox("editchange")
                End If
            End If
        End With
    End Sub

    Private Sub dgvItemDetail_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemDetail.CellDoubleClick
        Dim intCounter As Integer = 0
        If chkOptimized.Checked = True Then
            If e.RowIndex >= 0 AndAlso e.ColumnIndex = 1 Then
                Dim row As DataGridViewRow = dgvItemDetail.Rows(e.RowIndex)
                'row.Cells(enmColItem.Status).Value = Convert.ToBoolean(row.Cells(enmColItem.Status).EditedFormattedValue)
                'If Convert.ToBoolean(row.Cells(enmColItem.Status).Value) Then
                Dim strItemCode As String = row.Cells(enmColItem.ItemCode).Value
                Dim strItemdesc As String = row.Cells(enmColItem.ItemDesc).Value
                txtItemCode.Text = strItemCode
                lblItemDesc.Text = strItemdesc
                FillCustomerGrid(strItemCode)
                'End If
            End If
        End If
    End Sub

   

    Private Sub dgvItemDetail_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvItemDetail.CellValueChanged
        Dim intCounter As Integer = 0
        If chkOptimized.Checked = True Then
            If e.RowIndex >= 0 AndAlso e.ColumnIndex = 0 Then
                Dim row As DataGridViewRow = dgvItemDetail.Rows(e.RowIndex)
                row.Cells(enmColItem.Status).Value = Convert.ToBoolean(row.Cells(enmColItem.Status).EditedFormattedValue)
                If Convert.ToBoolean(row.Cells(enmColItem.Status).Value) Then
                    Dim strItemCode As String = row.Cells(enmColItem.ItemCode).Value
                    Dim strItemdesc As String = row.Cells(enmColItem.ItemDesc).Value
                    txtItemCode.Text = strItemCode
                    lblItemDesc.Text = strItemdesc
                    FillCustomerGrid(strItemCode)
                End If
            End If
        End If
    End Sub
    Private Sub FillCustomerGrid(ByVal strItemCode As String)
        Try
            If btnGrp1.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                Dim dblPendSchAsOnDate As Double = 0
                Dim dblBackLogQty As Double = 0
                Dim dblNextDaySchedule As Double = 0
                Dim dblTotalAllowedDispatch As Double = 0
                With sprPendingSchedule
                    .Row = 0
                    .Col = enmCol.BackLogQty : .ColHidden = False
                    .Col = enmCol.ForecastQty : .ColHidden = True
                    .Col = enmCol.NextDaySch : .ColHidden = False
                    .Col = enmCol.PendingDSAsOnDate : .ColHidden = False
                    .Col = enmCol.PercAsPerBSR : .ColHidden = False
                    .Col = enmCol.AllowedDispatchQty : .ColHidden = False
                End With
                ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
                sprPendingSchedule.MaxRows = 0

                If Not IsRecordExists("SELECT TOP 1 1 FROM ITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE ='" & txtItemCode.Text.Trim & "'") Then
                    MsgBox("Invalid Item Code !", MsgBoxStyle.Exclamation, ResolveResString(100))
                    txtItemCode.Text = String.Empty
                    txtItemCode.Focus()
                    Return
                End If

                If txtItemCode.Text.Trim.Length > 0 Then
                    lblBSRStock.Text = GetStockInHand("01B1", txtItemCode.Text.Trim)
                End If
                Dim strDSCriteria As String = ""
                If OptCurrent.Checked Then
                    strDSCriteria = "Current"
                Else
                    strDSCriteria = "Future"
                End If

                Using sqlCmd As SqlCommand = New SqlCommand
                    With sqlCmd
                        .CommandType = CommandType.StoredProcedure
                        .CommandText = "USP_PENDING_SCHEDULE_ASONDATE_FOR_DISP_ADVICE_NEWAuto"
                        '.CommandText = "USP_PENDING_SCHEDULE_ASONDATE_FOR_DISP_ADVICE_NEWpriti"
                        .CommandTimeout = 0
                        .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                        .Parameters.Add("@ITEM_CODE", SqlDbType.VarChar, 16).Value = strItemCode
                        .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                        .Parameters.Add("@DSCriteria", SqlDbType.VarChar, 20).Value = strDSCriteria

                        SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                        Dim strSQL As String = "SELECT * FROM TMP_BSR_PENDINGSCHEDULEFORDISPADVICE WHERE UNIT_CODE ='" & gstrUNITID & "' AND IP_ADDRESS ='" & gstrIpaddressWinSck & "' and TotalAllowedDispatchasOnDate > 0   order by TotalAllowedDispatchasOnDate desc"

                        If chkOptimized.Checked = True Then
                            Using dtSch As DataTable = SqlConnectionclass.GetDataTable(strSQL)

                                If dtSch.Rows.Count > 0 Then
                                    Dim dblNetStock As Integer = 0
                                    Dim dblPendingNetStock As Integer = 0
                                    If txtItemCode.Text.Trim.Length > 0 Then
                                        Dim dblStockQty As Int64 = SqlConnectionclass.ExecuteScalar("SELECT sum(PENDING_PICKLIST_QTY) FROM TMP_BSR_PENDINGSTOCKFORDISPADVICE WHERE UNIT_CODE ='" & gstrUNITID & "' AND IP_ADDRESS ='" & gstrIpaddressWinSck & "'")
                                        LblpendingStock.Text = dblStockQty
                                        lblNetStock.Text = lblBSRStock.Text - LblpendingStock.Text
                                        dblNetStock = lblNetStock.Text
                                        dblPendingNetStock = dblNetStock
                                    End If
                                    Dim dblDispatchQty As Double = 0
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
                                            dblTotalAllowedDispatch = Val(.Text)
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
                                            .Col = enmCol.Mode : .Text = "Road"
                                            '.BlockMode = True
                                            '.Row = .MaxRows : .Row2 = .MaxRows
                                            '.Col = enmCol.DispatchQty : .Col2 = enmCol.DispatchQty : .TypeFloatMin = 0 : .TypeFloatMax = Val(Row("TotalAllowedDispatchasOnDate"))
                                            '.BlockMode = False
                                            If dblPendingNetStock > 0 Then
                                                If dblTotalAllowedDispatch < dblPendingNetStock Then
                                                    .Col = enmCol.DispatchQty : .Text = Convert.ToString(Row("TotalAllowedDispatchasOnDate"))
                                                    dblPendingNetStock = dblPendingNetStock - dblTotalAllowedDispatch
                                                Else
                                                    .Col = enmCol.DispatchQty : .Text = dblPendingNetStock
                                                    dblPendingNetStock = 0
                                                End If
                                            End If
                                            .Col = enmCol.DispatchQty
                                            dblDispatchQty += Val(.Text)
                                            txtAccumuDispQty.Text = dblDispatchQty
                                        End With


                                    Next
                                    If chkOptimized.Checked = True Then
                                        strSQL = "insert into  TMP_BSR_PENDINGSCHEDULEFORDISPADVICELog ([UNIT_CODE],[IP_ADDRESS],[ITEM_CODE],[CUSTOMER_CODE], " & _
                                        "[CUST_NAME],[CUST_DRGNO],[DRG_DESC],[FORECAST_QTY],[BACKLOG_QTY],[PENDING_AS_ON_DATE],[NEXT_DAY_SCHEDULE], " & _
                                        "[CURRENT_MONTH_FUTURE_SCHEDULE],[FOR_THE_MONTH_SCHEDULE],[ALLOWED_PICKLIST_QTY],[PENDING_PICKLIST_QTY],[STD_PKG_QTY], " & _
                                        "[NEED_BY_DATE],[FIFO_QTY],[TotalAllowedDispatchasOnDate],[AllowedSOQty],[OpenAllowedSOQty],[CloseAllowedSOQty], " & _
                                        "[User_Id],[Ent_dt])  " & _
                                        " select  [UNIT_CODE],[IP_ADDRESS],[ITEM_CODE],[CUSTOMER_CODE],[CUST_NAME],[CUST_DRGNO],[DRG_DESC],[FORECAST_QTY], " & _
                                        "[BACKLOG_QTY],[PENDING_AS_ON_DATE],[NEXT_DAY_SCHEDULE],[CURRENT_MONTH_FUTURE_SCHEDULE],[FOR_THE_MONTH_SCHEDULE], " & _
                                        "[ALLOWED_PICKLIST_QTY],[PENDING_PICKLIST_QTY],[STD_PKG_QTY],[NEED_BY_DATE],[FIFO_QTY],[TotalAllowedDispatchasOnDate], " & _
                                        "[AllowedSOQty],[OpenAllowedSOQty],[CloseAllowedSOQty],'" & mP_User & "',getdate() from TMP_BSR_PENDINGSCHEDULEFORDISPADVICE WHERE UNIT_CODE ='" & gstrUNITID & "' AND IP_ADDRESS ='" & gstrIpaddressWinSck & "' order by TotalAllowedDispatchasOnDate desc"
                                        SqlConnectionclass.ExecuteNonQuery(strSQL)
                                    End If
                                Else
                                    If chkOptimized.Checked = True Then
                                        MsgBox("No Delivery Schedule found !", MsgBoxStyle.Exclamation, ResolveResString(100))
                                    End If
                                End If
                            End Using
                        Else
                            '' Auto Save start here
                            dtAuto = SqlConnectionclass.GetDataTable(strSQL)
                        End If
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

    Private Sub AutoSaveData()
        Try
            intSave = 0
            Dim intProgess As Integer = 0
            If chkOptimized.Checked = False Then
                If dgvItemDetail.Rows.Count > 0 Then
                    ProgressBar1.Value = 0
                    With dgvItemDetail

                        For i As Integer = 0 To dgvItemDetail.Rows.Count - 1
                            Dim strItemCode As String = Convert.ToString(dgvItemDetail.Rows(i).Cells(enmColItem.ItemCode).Value).Trim()
                            Dim strItemdesc As String = Convert.ToString(dgvItemDetail.Rows(i).Cells(enmColItem.ItemDesc).Value).Trim()
                            txtItemCode.Text = strItemCode
                            lblItemDesc.Text = strItemdesc
                            FillCustomerGrid(strItemCode)
                            DataTableSavedData()
                            Dim percentage As Double = (i / dgvItemDetail.Rows.Count) * 100
                            ProgressBar1.Value = Int32.Parse(Math.Truncate(percentage).ToString())
                        Next
                        dgvItemDetail.Columns.Clear()
                    End With
                Else
                    MsgBox("No data found to save")
                End If
            Else
                GridSavedData()

            End If
            'If chkOptimized.Checked = False Then
            If intSave > 0 Then
                MsgBox("Data Saved successfully")
            Else
                MsgBox("No Data found to Save")
            End If
            'End If
            If chkOptimized.Checked = False Then
                btnGrp1.Revert()
            End If

            'EnableControls(False, Me, True)
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
            ProgressBar1.Value = 0
        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub
    Private Function ValidateOptimization()
        If ValidateData() = False Then
            Return False
            Exit Function
        End If
        'If MsgBox("Do you want to save record(s)", MsgBoxStyle.Question + MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.No Then
        '    sprPendingSchedule.Row = sprPendingSchedule.MaxRows : sprPendingSchedule.Col = enmCol.DispatchQty
        '    sprPendingSchedule.Action = FPSpreadADO.ActionConstants.ActionActiveCell
        '    sprPendingSchedule.Focus()
        '    Return False
        'End If

        'Priti Starts for schedule shortage issue With out transaction.
        Dim strDSCriteria As String = ""
        If OptCurrent.Checked Then
            strDSCriteria = "Current"
        Else
            strDSCriteria = "Future"
        End If
        Dim blnCheckSchedulOnSaveWOTrans As Boolean = False
        blnCheckSchedulOnSaveWOTrans = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("SELECT isnull(CheckSchedulOnSaveWOTrans,0) FROM sales_parameter WHERE UNIT_CODE ='" & gstrUNITID & "'"))
        If blnCheckSchedulOnSaveWOTrans Then

            Using sqlCmd As SqlCommand = New SqlCommand
                With sqlCmd
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "USP_PENDING_SCHEDULE_ASONDATE_FOR_DISP_ADVICE_NEWAuto"
                    .CommandTimeout = 0
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@ITEM_CODE", SqlDbType.VarChar, 16).Value = txtItemCode.Text.Trim
                    .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                    .Parameters.Add("@DSCriteria", SqlDbType.VarChar, 20).Value = strDSCriteria
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
                                                MsgBox("Only " & Convert.ToSingle(dblAllowDispatchQty) & " Schedule Qty is avaiable for Customer " & InnerCust & " .You cannot save Dispatch Advice !", MsgBoxStyle.Critical, ResolveResString(100))
                                                Return False
                                                Exit Function
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
                Return False
                Exit Function
            End If
        End If
       
        Return True
    End Function

    Private Sub GridSavedData()
        Dim strSQL As String = String.Empty
        Dim strCustCode As String = String.Empty
        Dim strMode As String = String.Empty
        Dim strDrgNo As String = String.Empty
        Dim strDrgDesc As String = String.Empty
        Dim dblForeCastQty, dblDispatchQty, dblPendingAsOnDate, _
            dblBackLogQty, dblNextDaySch, dblStdPkgQty, dblPercBSRStock As Double
        Dim intDispatchAdviceNo As Integer
        Dim dblForTheMonthSch As Double, dblFIFOQty As Double, dblAllowedbldispatchqty As Double
        Dim strNeedByDate As Object
        Try
            Dim strDSCriteria As String = ""
            If OptCurrent.Checked Then
                strDSCriteria = "Current"
            Else
                strDSCriteria = "Future"
            End If
            If chkOptimized.Checked = True Then
                If ValidateOptimization() = False Then
                    Exit Sub
                End If
            End If
            strSQL = "DELETE TMP_DISP_ADVICE_FOR_PENDING_SCH WHERE UNIT_CODE='" & gstrUNITID & "' AND IP_ADDRESS='" & gstrIpaddressWinSck & "'"
            SqlConnectionclass.ExecuteNonQuery(strSQL)
            If txtAccumuDispQty.Text > 0 Then
                SqlConnectionclass.BeginTrans()
                With sprPendingSchedule
                    If .MaxRows > 0 Then
                        Using sqlcmd As SqlCommand = New SqlCommand
                            strSQL = "INSERT TMP_DISP_ADVICE_FOR_PENDING_SCH([UNIT_CODE],[IP_ADDRESS],[CUST_CODE],[CUST_DRG_NO],[FORECAST_QTY],[BACKLOG_QTY],[PENDING_DS_ASONDATE],[NEXT_DAY_SCH],[STD_PKG_QTY],[DISPATCH_QTY],FOR_THE_MONTH_SCH,NEED_BY_DATE,FIFO_QTY,TotalAllowedDispatchasOnDate,Mode)"
                            strSQL += " VALUES (@UNIT_CODE,@IP_ADDRESS,@CUST_CODE,@CUST_DRGNO,@FORECASTQTY,@BACKLOGQTY,@PENDINGASONDATE,@NEXTDATESCHEDULE,@PACKQTY,@DISPATCHQTY,@FORTHEMONTHSCHEDULE,@NEED_BY_DATE,@FIFOQTY,@TotalAllowedDispatchasOnDate,@Mode)"
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
                                .Parameters.Add("@Mode", SqlDbType.VarChar, 20)
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
                                    .Col = enmCol.Mode : strMode = .Text.Trim
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
                                        .Parameters("@Mode").Value = strMode
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
                                .Parameters.Add("@ScheduleType", SqlDbType.VarChar, 20).Value = strDSCriteria
                                If chkOptimized.Checked = False Then
                                    .Parameters.Add("@EntryType", SqlDbType.VarChar, 10).Value = "AUTO"
                                Else
                                    .Parameters.Add("@EntryType", SqlDbType.VarChar, 10).Value = "OPT"
                                End If
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
                intSave += 1
                SqlConnectionclass.CommitTran()
            End If


            'txtItemCode.Text = ""
            'lblItemDesc.Text = ""
            'lblBSRStock.Text = ""
            'lblNetStock.Text = ""
            'LblpendingStock.Text = ""
            'txtAccumuDispQty.Text = ""
            'txtRemanDisQty.Text = ""
            'sprPendingSchedule.MaxRows = 0
        Catch ex As Exception
            Throw ex
        End Try


    End Sub

    Private Sub DataTableSavedData()
        Dim strSQL As String = String.Empty
        Dim strCustCode As String = String.Empty
        Dim strMode As String = String.Empty
        Dim strDrgNo As String = String.Empty
        Dim strDrgDesc As String = String.Empty
        Dim dblForeCastQty, dblDispatchQty, dblPendingAsOnDate, _
            dblBackLogQty, dblNextDaySch, dblStdPkgQty, dblPercBSRStock As Double
        Dim intDispatchAdviceNo As Integer
        Dim dblForTheMonthSch As Double, dblFIFOQty As Double, dblAllowedbldispatchqty As Double
        Dim strNeedByDate As Object
        Dim dblTotalDispatch As Double = 0

        Dim DSFIFo As New DataTable
        DSFIFo.Columns.Add(New DataColumn("DISP_ADVICE_NO"))
        DSFIFo.Columns.Add(New DataColumn("ITEM_CODE"))
        DSFIFo.Columns.Add(New DataColumn("CUST_DRGNO"))
        DSFIFo.Columns.Add(New DataColumn("CUSTOMER_CODE"))
        DSFIFo.Columns.Add(New DataColumn("DSDateTime"))
        DSFIFo.Columns.Add(New DataColumn("Avaiable_Quantity", System.Type.GetType("System.Double")))
        DSFIFo.Columns.Add(New DataColumn("DispatchQty", System.Type.GetType("System.Double")))

        Try
            Dim strDSCriteria As String = ""
            If OptCurrent.Checked Then
                strDSCriteria = "Current"
            Else
                strDSCriteria = "Future"
            End If

            strSQL = "DELETE TMP_DISP_ADVICE_FOR_PENDING_SCH WHERE UNIT_CODE='" & gstrUNITID & "' AND IP_ADDRESS='" & gstrIpaddressWinSck & "'"
            SqlConnectionclass.ExecuteNonQuery(strSQL)


            Using sqlcmd As SqlCommand = New SqlCommand
                strSQL = "INSERT TMP_DISP_ADVICE_FOR_PENDING_SCH([UNIT_CODE],[IP_ADDRESS],[CUST_CODE],[CUST_DRG_NO],[FORECAST_QTY],[BACKLOG_QTY],[PENDING_DS_ASONDATE],[NEXT_DAY_SCH],[STD_PKG_QTY],[DISPATCH_QTY],FOR_THE_MONTH_SCH,NEED_BY_DATE,FIFO_QTY,TotalAllowedDispatchasOnDate,Mode)"
                strSQL += " VALUES (@UNIT_CODE,@IP_ADDRESS,@CUST_CODE,@CUST_DRGNO,@FORECASTQTY,@BACKLOGQTY,@PENDINGASONDATE,@NEXTDATESCHEDULE,@PACKQTY,@DISPATCHQTY,@FORTHEMONTHSCHEDULE,@NEED_BY_DATE,@FIFOQTY,@TotalAllowedDispatchasOnDate,@Mode)"
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
                    .Parameters.Add("@Mode", SqlDbType.VarChar, 20)
                End With

                Dim dblTotalAllowedDispatch As Double = 0

                If dtAuto.Rows.Count > 0 Then
                    Dim dblNetStock As Integer = 0
                    Dim dblPendingNetStock As Integer = 0
                    If txtItemCode.Text.Trim.Length > 0 Then
                        Dim dblStockQty As Int64 = SqlConnectionclass.ExecuteScalar("SELECT sum(isnull(PENDING_PICKLIST_QTY,0)) FROM TMP_BSR_PENDINGSTOCKFORDISPADVICE WHERE UNIT_CODE ='" & gstrUNITID & "' AND IP_ADDRESS ='" & gstrIpaddressWinSck & "'")
                        LblpendingStock.Text = dblStockQty
                        lblNetStock.Text = lblBSRStock.Text - LblpendingStock.Text
                        dblNetStock = lblNetStock.Text
                        dblPendingNetStock = dblNetStock
                    End If
                    If dblPendingNetStock > 0 Then

                        '' START
                        Dim strQuery As String = ""
                        
                        Using sqlCmd1 As SqlCommand = New SqlCommand
                            With sqlCmd1
                                .CommandType = CommandType.StoredProcedure
                                .CommandText = "GETFIFODSDATA_AUTODISPATCHADVICE"
                                .CommandTimeout = 0
                                .Parameters.Add("@UNITCODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                                .Parameters.Add("@ITEM_CODE", SqlDbType.VarChar, 16).Value = txtItemCode.Text
                                .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                                .Parameters.Add("@DSCriteria", SqlDbType.VarChar, 20).Value = strDSCriteria
                                SqlConnectionclass.ExecuteNonQuery(sqlCmd1)
                            End With
                        End Using
                        strQuery = "select dsdatetime,Schedule_Quantity -  isnull(PendingQty,0) as Schedule,CUSTOMER_CODE as account_code from TMP_FIFOWISEDS " & _
                        "where UNIT_CODE ='" & gstrUNITID & "' AND IP_ADDRESS ='" & gstrIpaddressWinSck & "' order by dsdatetime asc "
                      

                        Dim dt As DataTable = SqlConnectionclass.GetDataTable(strQuery)
                        Dim strCustDrgNo As String = ""
                        Dim DtCustomer As String = ""
                        Dim dtSchedule As Double = 0
                        Dim dtDSdate As String = ""
                        Dim dtAllowedQTy As Double = 0
                        Dim dtDispatchQty As Double = 0
                        Dim dtCustDispatchQty As Double = 0
                        Dim dtUpdateDispatchQty As Double = 0
                        Dim dblScheduleQty As Double = 0
                        Dim Qty As Integer = 0
                        Dim dblPackQty As Integer = 0


                        dtAuto.Columns.Add(New DataColumn("DispatchQty"))
                        For Each Row As DataRow In dt.Rows
                            If dblPendingNetStock > 0 Then
                                DtCustomer = Convert.ToString(Row("account_code"))
                                dtSchedule = Convert.ToDouble(Row("Schedule"))
                                dblScheduleQty = Convert.ToDouble(Row("Schedule"))
                                dtDSdate = Convert.ToString(Row("dsdatetime"))
                                For Each row1 As DataRow In dtAuto.Select("CUSTOMER_CODE ='" + Convert.ToString(DtCustomer) + "'")
                                    dtCustDispatchQty = 0
                                    Qty = 0
                                    dtAllowedQTy = row1("TotalAllowedDispatchasOnDate")
                                    dblPackQty = row1("STD_PKG_QTY")
                                    strCustDrgNo = row1("CUST_DRGNO")
                                    If IsDBNull(row1("DispatchQty")) Then
                                        dtCustDispatchQty = 0
                                    Else
                                        dtCustDispatchQty = row1("DispatchQty")
                                    End If

                                    If dtDispatchQty <> 0 Then
                                        dtAllowedQTy = dtAllowedQTy - dtCustDispatchQty
                                    End If
                                    'If blnCheckPackQtymultiple_onDA Then
                                    '    Qty = Math.Floor(dtSchedule / dblPackQty)
                                    '    dtSchedule = Qty * dblPackQty
                                    '    If dtSchedule < dblPendingNetStock And dtSchedule < dtAllowedQTy Then
                                    '        Qty = 0
                                    '        Qty = Math.Floor(dtAllowedQTy / dblPackQty)
                                    '        dtAllowedQTy = Qty * dblPackQty
                                    '    End If

                                    'End If
                                    Exit For
                                Next

                                If dtSchedule < dblPendingNetStock And dtSchedule < dtAllowedQTy Then
                                    If blnCheckPackQtymultiple_onDA Then
                                        Qty = Math.Floor(dtSchedule / dblPackQty)
                                        dtSchedule = Qty * dblPackQty
                                    End If
                                    dtDispatchQty = dtSchedule
                                    dblPendingNetStock = dblPendingNetStock - dtSchedule
                                Else
                                    If dtAllowedQTy < dblPendingNetStock Then
                                        If blnCheckPackQtymultiple_onDA Then
                                            Qty = Math.Floor(dtAllowedQTy / dblPackQty)
                                            dtAllowedQTy = Qty * dblPackQty
                                        End If
                                        dtDispatchQty = dtAllowedQTy
                                        dblPendingNetStock = dblPendingNetStock - dtDispatchQty
                                    Else
                                        If blnCheckPackQtymultiple_onDA Then
                                            Qty = Math.Floor(dblPendingNetStock / dblPackQty)
                                            dtAllowedQTy = Qty * dblPackQty
                                        End If
                                        dtDispatchQty = dtAllowedQTy
                                        dblPendingNetStock = 0
                                        End If
                                End If

                                Dim myRow() As Data.DataRow
                                myRow = dtAuto.Select("CUSTOMER_CODE ='" + Convert.ToString(DtCustomer) + "'")
                                myRow(0)("DispatchQty") = dtDispatchQty + dtCustDispatchQty

                                Dim insertRow As DataRow
                                insertRow = DSFIFo.NewRow
                                insertRow("ITEM_CODE") = txtItemCode.Text
                                insertRow("CUST_DRGNO") = strCustDrgNo
                                insertRow("CUSTOMER_CODE") = DtCustomer
                                insertRow("Avaiable_Quantity") = dblScheduleQty
                                insertRow("DispatchQty") = dtDispatchQty
                                insertRow("DSDateTime") = dtDSdate
                                DSFIFo.Rows.Add(insertRow)
                            Else
                                Exit For
                            End If
                        Next
                        '' END


                        dblTotalDispatch = 0
                        For Each Row As DataRow In dtAuto.Rows
                            dblDispatchQty = 0
                            strCustCode = Convert.ToString(Row("CUSTOMER_CODE"))
                            strDrgNo = Convert.ToString(Row("CUST_DRGNO"))
                            If IsDBNull(Row("FORECAST_QTY")) Then
                                dblForeCastQty = 0
                            Else
                                dblForeCastQty = Convert.ToDouble(Row("FORECAST_QTY"))
                            End If

                            dblBackLogQty = Convert.ToDouble(Row("BACKLOG_QTY"))
                            dblPendingAsOnDate = Convert.ToDouble(Row("PENDING_AS_ON_DATE"))
                            If IsDBNull(Row("NEXT_DAY_SCHEDULE")) Then
                                dblNextDaySch = 0
                            Else
                                dblNextDaySch = Convert.ToDouble(Row("NEXT_DAY_SCHEDULE"))
                            End If

                            dblPercBSRStock = "0.00"
                            dblStdPkgQty = Convert.ToDouble(Row("STD_PKG_QTY"))
                            dblAllowedbldispatchqty = Convert.ToDouble(Row("TotalAllowedDispatchasOnDate"))
                            dblForTheMonthSch = Convert.ToDouble(Row("FOR_THE_MONTH_SCHEDULE"))
                            dblFIFOQty = Convert.ToDouble(Row("FIFO_QTY"))


                            If IsDBNull(Row("NEED_BY_DATE")) Then
                                strNeedByDate = DBNull.Value
                            Else
                                strNeedByDate = Convert.ToDateTime(Row("NEED_BY_DATE")).ToString(gstrDateFormat)
                            End If
                            strMode = "Road"
                            If blnCheckPackQtymultiple_onDA Then

                            End If
                            If IsDBNull(Row("DispatchQty")) Then
                                dblDispatchQty = 0
                            Else
                                dblDispatchQty = Convert.ToString(Row("DispatchQty"))
                            End If

                            'If dblPendingNetStock > 0 Then
                            '    If dblAllowedbldispatchqty < dblPendingNetStock Then
                            '        dblDispatchQty = Convert.ToString(Row("TotalAllowedDispatchasOnDate"))
                            '        dblPendingNetStock = dblPendingNetStock - dblDispatchQty
                            '    Else
                            '        dblDispatchQty = dblPendingNetStock
                            '        dblPendingNetStock = 0
                            '    End If
                            'End If

                            dblTotalDispatch += dblDispatchQty
                            txtAccumuDispQty.Text = dblTotalDispatch
                            If dblDispatchQty > 0 Then
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
                                    .Parameters("@Mode").Value = strMode
                                    SqlConnectionclass.ExecuteNonQuery(sqlcmd)
                                End With
                            End If
                        Next
                    End If
                End If
            End Using

            If dblTotalDispatch > 0 Then
                SqlConnectionclass.BeginTrans()
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
                        .Parameters.Add("@ScheduleType", SqlDbType.VarChar, 20).Value = strDSCriteria
                        If chkOptimized.Checked = False Then
                            .Parameters.Add("@EntryType", SqlDbType.VarChar, 10).Value = "AUTO"
                        Else
                            .Parameters.Add("@EntryType", SqlDbType.VarChar, 10).Value = "OPT"
                        End If
                        SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                        intDispatchAdviceNo = .Parameters("@DISPATCH_ADVICE_NO").Value
                        If intDispatchAdviceNo <= 0 Then
                            SqlConnectionclass.RollbackTran()
                            MsgBox("Unable to save Dispatch Advice !", MsgBoxStyle.Critical, ResolveResString(100))
                            Exit Sub
                        Else
                            If DSFIFo.Rows.Count > 0 Then
                                Dim dtDate As DateTime
                                For Each Row As DataRow In DSFIFo.Rows
                                    dtDate = Convert.ToDateTime(Row("DSDateTime"))
                                    strSQL = "INSERT INTO TMP_FIFOWISEDispatchAdvice (DISP_ADVICE_NO,ITEM_CODE,CUST_DRGNO,CUSTOMER_CODE, " & _
                                    "Avaiable_Quantity,Despatch_Qty,DSDateTime,UNIT_CODE,IP_ADDRESS ,User_Id,Ent_dt) Values (" & intDispatchAdviceNo & " , " & _
                                    "'" & Convert.ToString(Row("ITEM_CODE")) & "', " & _
                                    "'" & Convert.ToString(Row("CUST_DRGNO")) & "'," & _
                                    "'" & Convert.ToString(Row("CUSTOMER_CODE")) & "', " & _
                                    "'" & Convert.ToDouble(Row("Avaiable_Quantity")) & "'," & _
                                    "'" & Convert.ToString(Row("DispatchQty")) & "', " & _
                                    "'" & dtDate.ToString("dd MMM yyyy HH:mm") & "'," & _
                                    "'" & gstrUNITID & "'," & _
                                    "'" & gstrIpaddressWinSck & "' , " & _
                                    "'" & mP_User & "' , getdate()) "
                                    SqlConnectionclass.ExecuteNonQuery(strSQL)
                                Next
                            End If
                        End If
                    End With
                End Using


                intSave += 1
                SqlConnectionclass.CommitTran()
            End If


            'txtItemCode.Text = ""
            'lblItemDesc.Text = ""
            'lblBSRStock.Text = ""
            'lblNetStock.Text = ""
            'LblpendingStock.Text = ""
            'txtAccumuDispQty.Text = ""
            'txtRemanDisQty.Text = ""
            'sprPendingSchedule.MaxRows = 0
        Catch ex As Exception
            Throw ex
        End Try


    End Sub
    Private Sub FillItemSearchCategory()
        Try
            With cmbItem
                .DataSource = Nothing
                .Items.Clear()
                .DataSource = [Enum].GetNames(GetType(enmColItem))
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub FillCustomerSearchCategory()
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

   
    Private Sub txtItemsearch_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtItemsearch.TextChanged
        Dim intCounter As Integer
        Dim strText As String
        Dim Col As enmColItem
        Try
            'Col = DirectCast(Me.cmbItem.SelectedIndex + 1, enmColItem)
            'For intCounter = 1 To dgvItemDetail.Rows.Count - 1
            '    strText = dgvItemDetail.Rows(intCounter).Cells(enmColItem.ItemCode).Value

            'Next
            If Len(txtItemsearch.Text) = 0 Then Exit Sub
            For intCounter = 0 To dgvItemDetail.Rows.Count - 1
                If cmbItem.Text = "ItemCode" Then
                    strText = dgvItemDetail.Rows(intCounter).Cells(enmColItem.ItemCode).Value
                    If Trim(UCase(Mid(strText, 1, Len(txtItemsearch.Text)))) = Trim(UCase(txtItemsearch.Text)) Then
                        dgvItemDetail.CurrentCell = dgvItemDetail.Rows(intCounter).Cells(enmColItem.ItemCode)
                        Exit For
                    End If
                ElseIf cmbItem.Text = "ItemDesc" Then
                    strText = dgvItemDetail.Rows(intCounter).Cells(enmColItem.ItemDesc).Value
                    If Trim(UCase(Mid(strText, 1, Len(txtItemsearch.Text)))) = Trim(UCase(txtItemsearch.Text)) Then
                        dgvItemDetail.CurrentCell = dgvItemDetail.Rows(intCounter).Cells(enmColItem.ItemDesc)
                        Exit For
                    End If
                End If
                

            Next
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

 

    Private Sub chkOptimized_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOptimized.CheckedChanged
       

        If chkOptimized.Checked Then
            lblNote.Visible = True
            'EnableControls(True, Me, True)
            sprPendingSchedule.Enabled = True
            txtItemCode.Enabled = True
            txtItemCode.BackColor = Color.White
            cmdItemHelp.Enabled = True
            cmdShowSchedule.Enabled = True
            cboSearch.Enabled = True
            txtSearch.Enabled = True
            txtSearch.BackColor = Color.White
            cmbItem.Enabled = True
            txtItemsearch.Enabled = True
            txtItemsearch.BackColor = Color.White
           
        Else
            'btnGrp1.Revert()
            lblNote.Visible = False
            'EnableControls(False, Me, True)
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
            chkOptimized.Enabled = True
            'dgvItemDetail.Columns.Clear()
        End If
    End Sub

    Private Sub OptCurrent_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptCurrent.CheckedChanged
        ClearDataonDSCriteria()
    End Sub
    Private Sub ClearDataonDSCriteria()
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
    End Sub
    Private Sub OptFuture_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptFuture.CheckedChanged
       ClearDataonDSCriteria
    End Sub

End Class