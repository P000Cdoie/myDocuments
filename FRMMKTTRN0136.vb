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


Public Class FRMMKTTRN0136
    Dim intSave As Integer = 0
    Dim mintFormIndex As Integer
    Dim dtAuto As DataTable
    Dim blnCheckPackQtymultiple_onDA As Boolean = False

    Private Enum enmColItem
        Status = 0
        CustDrgNo
        ItemCode
        Qty1
        Qty2
        Qty3
        Qty4
        Qty5
        Qty6
        Qty7

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
            Call FitToClient(Me, GrpMain1, ctlHeader, grpButton, 500)
            Me.MdiParent = mdifrmMain
            mintFormIndex = mdifrmMain.AddFormNameToWindowList(Me.ctlHeader.Tag)

            'btnGrp1.ShowButtons(True, False, True, False)
            'EnableControls(False, Me, True)           
            RefreshScreen()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub
#End Region

#Region "Methods"


    Private Sub RefreshScreen()
        Try
            txtCustItem.Text = ""
            lblCusttemName.Text = ""
            txtcustomerhelp.Text = ""
            lblCustName.Text = ""
            txtDocNo.Text = ""
            txtDocDate.Text = ""
            dtfromdate.Format = DateTimePickerFormat.Custom
            dtfromdate.CustomFormat = gstrDateFormat
            dtfromdate.Value = GetServerDate()
            dttodate.Format = DateTimePickerFormat.Custom
            dttodate.CustomFormat = gstrDateFormat
            dttodate.Value = GetServerDate()

            CDPFromDate.Format = DateTimePickerFormat.Custom
            CDPFromDate.CustomFormat = gstrDateFormat
            CDPFromDate.Value = GetServerDate()
            CDPEndDate.Format = DateTimePickerFormat.Custom
            CDPEndDate.CustomFormat = gstrDateFormat
            CDPEndDate.Value = GetServerDate()

            dgvSODetail.Rows.Clear()
            dgvSODetail.Columns.Clear()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub


#End Region



    Private Sub cmdShowSchedule_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdShowSaleOrder.Click
        SetItemGridsHeader()
    End Sub

  
    Private Sub SetItemGridsHeader()
        Try
            If txtcustomerhelp.Text = "" Then
                MsgBox("Select Customer First", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            ElseIf txtDocNo.Text = "" Then
                MsgBox("Select Document No First", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            ElseIf txtCustItem.Text = "" Then
                MsgBox("Select Customer Drg No First", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If

            If DateDiff(DateInterval.Day, dtfromdate.Value, dttodate.Value) > 7 Then
                'MsgBox("Select only 7 days data.", MsgBoxStyle.Information, ResolveResString(100))
                'Exit Sub
            End If

            dgvSODetail.Columns.Clear()
            dgvSODetail.Rows.Clear()

            Dim StrSql As String = "select DISTINCT trans_date from custitem_mst C with (nolock),DailyMktSchedule_tempCDP COVI with (nolock) WHERE " & _
               "C.UNIT_CODE = COVI.UNIT_CODE and  C.Cust_Drgno = COVI.Cust_Drgno  AND C.active = 1 AND SCHUPLDREQD = 1 and  C.unit_code='" & gstrUNITID & "' and  " & _
               "C.ACCOUNT_CODE = '" & txtcustomerhelp.Text & "' AND COVI.Schedule_Flag=1 and Status=1 and Is_Distributed=0 and " & _
               "C.CUST_DRGNO ='" & txtCustItem.Text & "' AND doc_no='" & txtDocNo.Text & "' and Schedule_Quantity > 0 and trans_date  > ='" & VB6.Format(Me.dtfromdate.Value, "dd/mmm/yyyy") & "'" & _
               "and trans_date  < ='" & VB6.Format(Me.dttodate.Value, "dd/mmm/yyyy") & "' " & _
               "and trans_date  not in  (Select trans_date from DailyMktSchedule where unit_code='" & gstrUNITID & "'  and Account_Code='" & txtcustomerhelp.Text & "' and status=1 and  DOC_NO = '" & txtDocNo.Text & "'  and cust_drgno='" & txtCustItem.Text & "'  ) ORDER BY trans_date asc "
            Dim dt As DataTable = SqlConnectionclass.GetDataTable(StrSql)
            Dim intCount As Integer = 0
            Dim strQuery As String = ""
            Dim strFinalQuery As String = ""
            Dim strPivotFinalQuery As String = ""
            Dim strDate As String
            Dim strPivotDate As Date
            If dt.Rows.Count > 0 Then
                Dim objChkBox As New DataGridViewCheckBoxColumn
                objChkBox.Name = "Selection"
                objChkBox.HeaderText = "Select"

                dgvSODetail.Columns.Add(objChkBox)
                dgvSODetail.Columns.Add("CustDrgNo", "CustDrgNo")
                dgvSODetail.Columns.Add("ItemCode", "ItemCode")

                dgvSODetail.Columns(enmColItem.Status).Width = 50
                dgvSODetail.Columns(enmColItem.CustDrgNo).Width = 150
                dgvSODetail.Columns(enmColItem.ItemCode).Width = 100

                dgvSODetail.Columns(enmColItem.Status).Visible = False
                dgvSODetail.Columns(enmColItem.CustDrgNo).Frozen = True
                dgvSODetail.Columns(enmColItem.ItemCode).Frozen = True

                dgvSODetail.Columns(enmColItem.CustDrgNo).ReadOnly = True
                dgvSODetail.Columns(enmColItem.ItemCode).ReadOnly = True

                dgvSODetail.Columns(enmColItem.CustDrgNo).SortMode = DataGridViewColumnSortMode.NotSortable
                dgvSODetail.Columns(enmColItem.ItemCode).SortMode = DataGridViewColumnSortMode.NotSortable

                For Each dr As DataRow In dt.Rows
                    intCount = intCount + 1
                    dgvSODetail.Columns.Add("Qty" + Convert.ToString(intCount), "")
                    dgvSODetail.Columns("Qty" + Convert.ToString(intCount)).Width = 100
                    dgvSODetail.Columns("Qty" + Convert.ToString(intCount)).Visible = False

                    strDate = dr("trans_date")
                    strPivotDate = dr("trans_date")
                    strDate = DateFormateYYYYMMDD(strPivotDate)
                    dgvSODetail.Columns("Qty" + Convert.ToString(intCount)).HeaderText = strDate
                    dgvSODetail.Columns("Qty" + Convert.ToString(intCount)).Visible = True
                    dgvSODetail.Columns("Qty" + Convert.ToString(intCount)).SortMode = DataGridViewColumnSortMode.NotSortable
                    strQuery = strQuery + "[" + strDate + "]" + ","
                Next
                strFinalQuery = strQuery.Substring(0, strQuery.Length - 1)

            End If

            If dt.Rows.Count > 0 Then
            
                StrSql = "Select Cust_Drgno,item_code," & strFinalQuery & " from  ( " & _
                  "SELECT DISTINCT Cust_Drgno,C.item_code,convert(date,Shipment_Dt,103) as DeliveryDate,Shipment_Qty as qty from custitem_mst C with (nolock), " & _
                  "ScheduleProposalCalculations COVI with (nolock) WHERE " & _
                  "C.UNIT_CODE = COVI.UNIT_CODE and  C.Cust_Drgno = COVI.item_code   AND C.Account_Code=COVI.CONSIGNEE_CODE AND " & _
                  "C.active = 1 AND SCHUPLDREQD = 1 and  C.unit_code='" & gstrUNITID & "' AND C.ACCOUNT_CODE = '" & txtcustomerhelp.Text & "' AND doc_no='" & txtDocNo.Text & "'  and Cust_Drgno='" & txtCustItem.Text & "' " & _
                  "AND CONVERT(DATETIME,convert(varchar(11),Shipment_Dt,106),106)  between Product_Start_date and Product_End_date ) as SourceTable " & _
                  "Pivot (sum(qty) for DeliveryDate in (" & strFinalQuery & " )) as Pivottable"

                dt = New DataTable
                dt = SqlConnectionclass.GetDataTable(StrSql)
                If dt.Rows.Count > 0 Then
                    dgvSODetail.Rows.Clear()
                    dgvSODetail.Rows.Add(dt.Rows.Count)
                    i = 0
                    For Each dr As DataRow In dt.Rows
                        dgvSODetail.Rows(i).Cells(enmColItem.Status).Value = False
                        dgvSODetail.Rows(i).Cells(enmColItem.CustDrgNo).Value = dr("Cust_Drgno")
                        dgvSODetail.Rows(i).Cells(enmColItem.ItemCode).Value = dr("item_code")
                        For Int As Integer = 1 To intCount
                            dgvSODetail.Rows(i).Cells("Qty" + Convert.ToString(Int)).Value = dr(Int + 1)
                            If IsDBNull(dr(Int + 1)) = True Then
                                dgvSODetail.Rows(i).Cells("Qty" + Convert.ToString(Int)).Value = 0
                                dgvSODetail.Rows(i).Cells("Qty" + Convert.ToString(Int)).ReadOnly = True
                            End If
                        Next
                        i += 1
                    Next
                End If
            Else
                MsgBox("No Data found for this date range", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Function DateFormateYYYYMMDD(ByVal Dtp As Date) As String
        Try
            Dim StrDate, StrYear, StrMonth, StrDay As String
            StrDate = FormatDateTime(Dtp, DateFormat.ShortDate)
            StrMonth = Month(Dtp)
            If StrMonth.Length = 1 Then
                StrMonth = "0" + StrMonth
            End If
            StrDay = Convert.ToString(Dtp.Day)
            If StrDay.Length = 1 Then
                StrDay = "0" + StrDay
            End If
            StrYear = Year(Dtp)
            StrDate = StrYear + "-" + StrMonth + "-" + StrDay

            Return StrDate
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function Validations() As Boolean
        Try
            If txtcustomerhelp.Text = "" Then
                MsgBox("Select Customer First", MsgBoxStyle.Information, ResolveResString(100))
                Return False
            ElseIf txtDocNo.Text = "" Then
                MsgBox("Select Document No First", MsgBoxStyle.Information, ResolveResString(100))
                Return False
            ElseIf txtCustItem.Text = "" Then
                MsgBox("Select Customer Drg No First", MsgBoxStyle.Information, ResolveResString(100))
                Return False
            End If
            If dgvSODetail.Rows.Count = 0 Then
                MsgBox("There is no Schedule for Update.", MsgBoxStyle.Information, ResolveResString(100))
                Return False
            End If
            For intColumn As Integer = 0 To dgvSODetail.Columns.Count - 1
                Dim intTotal As Integer = 0
                If intColumn > 2 Then
                    For intRow As Integer = 0 To dgvSODetail.Rows.Count - 1
                        intTotal += dgvSODetail.Rows(intRow).Cells(intColumn).Value
                    Next
                    Dim strColumnHeader As String = dgvSODetail.Columns(intColumn).HeaderText
                    Dim strSql As String = "select Top 1 Shipment_Qty from ScheduleProposalCalculations COVI with (nolock) " & _
                    "WHERE  unit_code='" & gstrUNITID & "' and  CONSIGNEE_CODE = '" & txtcustomerhelp.Text & "' AND item_code ='" & txtCustItem.Text & "' " & _
                    "AND doc_no='" & txtDocNo.Text & "' and Shipment_Dt='" & strColumnHeader & "'"
                    Dim intTotalSchedule As Integer = SqlConnectionclass.ExecuteScalar(strSql)
                    If intTotal > intTotalSchedule Then
                        MsgBox("Total Schedule should not be greater than " & intTotalSchedule & " For date " & strColumnHeader & " .", MsgBoxStyle.Information, ResolveResString(100))
                        Return False
                    ElseIf intTotal < intTotalSchedule Then
                        MsgBox("Total Schedule should not be Less than " & intTotalSchedule & " For date " & strColumnHeader & " .", MsgBoxStyle.Information, ResolveResString(100))
                        Return False
                    End If
                End If
            Next
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub cmdcusthelp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdcusthelp.Click
        Try
            Dim strCustHelp() As String
            Dim strsql As String
            If DateDiff(DateInterval.Day, dtfromdate.Value, dttodate.Value) > 7 Then
                'MsgBox("Select only 7 days data.", MsgBoxStyle.Information, ResolveResString(100))
                'Exit Sub
            End If

            strCustHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct SUC.CONSIGNEE_CODE as Customer, " & _
            "CUST_NAME as CustomerName from SCHEDULE_UPLOAD_COVISINT_DTL  (NOLOCK) SUC,Customer_mst CM  where SUC.CONSIGNEE_CODE=CM.Customer_Code and " & _
            "SUC.UNIT_CODE=CM.UNIT_CODE  and   Convert(DATE,SUC.ent_dt,103)  > ='" & VB6.Format(Me.CDPFromDate.Value, "dd/mmm/yyyy") & "'  AND " & _
            "Convert(DATE,SUC.ent_dt,103)    < ='" & VB6.Format(Me.CDPEndDate.Value, "dd/mmm/yyyy") & "' AND SUC.UNIT_CODE='" & gstrUNITID & "'", "Customer Codes List", 1)
            If UBound(strCustHelp) <> -1 Then
                If strCustHelp(0) <> "0" Then
                    txtcustomerhelp.Text = Trim(strCustHelp(0))
                    lblCustName.Text = strCustHelp(1)
                    txtCustItem.Text = ""
                    lblCusttemName.Text = ""
                    txtDocNo.Text = ""
                    txtDocDate.Text = ""
                    dgvSODetail.RowCount = 0
                Else
                    txtcustomerhelp.Text = ""
                    MsgBox("No Record Available", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub btnRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        Try
            RefreshScreen()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub


    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub



    Private Sub cmdCustItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCustItem.Click
        Try
            Dim strCustItemHelp() As String
            Dim strsql As String
            If txtcustomerhelp.Text = "" Then
                MsgBox("Select Customer First", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            ElseIf txtDocNo.Text = "" Then
                MsgBox("Select Document No First", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If

            strsql = "SELECT DISTINCT  C.cust_drgno as Custdrgno,Account_code as Customer from DailyMktSchedule_tempCDP C with (nolock) " & _
            "where   unit_code='" & gstrUNITID & "' and C.ACCOUNT_CODE ='" & txtcustomerhelp.Text & "' AND DOC_NO = '" & txtDocNo.Text & "'"
            strCustItemHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strsql, "CustItem List", 1)
            If UBound(strCustItemHelp) <> -1 Then
                If strCustItemHelp(0) <> "0" Then
                    txtCustItem.Text = Trim(strCustItemHelp(0))
                Else

                    MsgBox("No Record Available", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub



    Private Sub cmdDocNoHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdDocNoHelp.Click
        Try
            Dim strDocHelp() As String
            Dim strsql As String
            If txtcustomerhelp.Text = "" Then
                MsgBox("Select Customer First", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If

            strsql = "select distinct doc_no as DocumentNo,convert(date,ent_dt,103) as DocumentDate from " & _
            "DailyMktSchedule_tempCDP (NOLOCK)  where Status=1 and Schedule_Flag=1 and  Account_Code ='" & txtcustomerhelp.Text & "'  and " & _
            "Convert(DATE,ent_dt,103)  > ='" & VB6.Format(Me.CDPFromDate.Value, "dd/mmm/yyyy") & "'  and " & _
            "Convert(DATE,ent_dt,103)  < ='" & VB6.Format(Me.CDPEndDate.Value, "dd/mmm/yyyy") & "'   AND UNIT_CODE='" & gstrUNITID & "'  and Is_Distributed=0 "
            strDocHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strsql, "Document No List", 1)
            If UBound(strDocHelp) <> -1 Then
                If strDocHelp(0) <> "0" Then
                    txtDocNo.Text = Trim(strDocHelp(0))
                    txtDocDate.Text = Trim(strDocHelp(1))
                Else

                    MsgBox("No Record Available", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            Dim strSql As String = ""
            Dim StrDrgNo As String = ""
            Dim strItemCode As String = ""
            Dim intQty As Integer = 0
            Dim intQty2 As Integer = 0
            Dim intQty3 As Integer = 0
            Dim intQty4 As Integer = 0
            Dim intQty5 As Integer = 0
            Dim intQty6 As Integer = 0
            Dim intQty7 As Integer = 0
            Dim strDate As String = ""
            If Validations() = False Then
                Exit Sub
            End If

            If MsgBox("Please review again because once data saved than it will not show again or not available for edit.", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, ResolveResString(100)) = MsgBoxResult.No Then
                Exit Sub
            End If

            SqlConnectionclass.CloseGlobalConnection()
            SqlConnectionclass.OpenGlobalConnection()
            SqlConnectionclass.BeginTrans()

            Dim intDocNo As Integer = SqlConnectionclass.ExecuteScalar("SELECT ISNULL(MAX(DOC_NO),0) + 1 FROM DailyMkt_ManualCDPItem ")

            For intcounter As Integer = 0 To dgvSODetail.Rows.Count - 1
                StrDrgNo = dgvSODetail.Rows(intcounter).Cells(enmColItem.CustDrgNo).Value
                strItemCode = dgvSODetail.Rows(intcounter).Cells(enmColItem.ItemCode).Value
                Dim int As Integer = 1
                For intCount As Integer = 1 To dgvSODetail.ColumnCount - 3
                    strDate = dgvSODetail.Columns("Qty" + Convert.ToString(int)).HeaderText
                    intQty = dgvSODetail.Rows(intcounter).Cells("Qty" + Convert.ToString(int)).Value
                    If strDate.Length > 0 Then
                        strSql = "INSERT INTO DailyMkt_ManualCDPItem (DocumentNo,ACCOUNT_CODE,Trans_date,Item_code," & _
                                "Cust_Drgno,Schedule_Flag,Schedule_Quantity,Despatch_Qty,Status,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId," & _
                                "consignee_code,filetype,doc_no, unit_code) Values ('" & intDocNo & "', '" & txtcustomerhelp.Text & "'," & _
                                "'" & strDate & "','" & strItemCode & "', " & _
                                "'" & StrDrgNo & "'," & _
                                "1,'" & intQty & "',0,1,getdate(),'" & mP_User & "',getdate(),'" & mP_User & "'," & _
                                "'" & txtcustomerhelp.Text & "', 'COVISINT', '" & txtDocNo.Text & "', '" & gstrUNITID & "')"
                        SqlConnectionclass.ExecuteNonQuery(strSql)
                        int = int + 1
                    End If
                    Dim strUpdate = "Update DailyMktSchedule_tempCDP set Is_Distributed=1 where unit_code='" & gstrUNITID & "' and account_code='" & txtcustomerhelp.Text & "' and " & _
                    "trans_date='" & strDate & "' and Cust_Drgno='" & StrDrgNo & "'  and Item_code='" & strItemCode & "'"
                    SqlConnectionclass.ExecuteNonQuery(strUpdate)
                Next

            

            Next

            strSql = " INSERT INTO DailyMktschedule ( [Account_Code],[Trans_date],[Item_code],[Cust_Drgno],[Schedule_Flag],[Schedule_Quantity], " & _
                     "[Despatch_Qty],[Status],[Ent_dt],[Ent_UserId],[Upd_dt],[Upd_UserId],[Consignee_Code],[FILETYPE],[DOC_NO],[UNIT_CODE] )  " & _
                     "SELECT [Account_Code],[Trans_date],[Item_code],[Cust_Drgno],[Schedule_Flag],[Schedule_Quantity],[Despatch_Qty],[Status]," & _
                     "[Ent_dt],[Ent_UserId],[Upd_dt],[Upd_UserId],[Consignee_Code],[FILETYPE],[DOC_NO],[UNIT_CODE]  FROM DailyMkt_ManualCDPItem  " & _
                     "where DocumentNo='" & intDocNo & "'"
            SqlConnectionclass.ExecuteNonQuery(strSql)

            SqlConnectionclass.CommitTran()
            MsgBox("Schedule  Updated Successfully ", MsgBoxStyle.Information, ResolveResString(100))
            RefreshScreen()

        Catch ex As Exception
            SqlConnectionclass.RollbackTran()
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    'Private Sub dtfromdate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtfromdate.TextChanged
    '    txtCustItem.Text = ""
    '    lblCusttemName.Text = ""
    '    txtcustomerhelp.Text = ""
    '    lblCustName.Text = ""
    '    txtDocNo.Text = ""
    '    txtDocDate.Text = ""
    'End Sub

    'Private Sub dttodate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dttodate.TextChanged
    '    txtCustItem.Text = ""
    '    lblCusttemName.Text = ""
    '    txtcustomerhelp.Text = ""
    '    lblCustName.Text = ""
    '    txtDocNo.Text = ""
    '    txtDocDate.Text = ""

    '    dgvSODetail.Columns.Clear()
    '    dgvSODetail.Rows.Clear()
    'End Sub

    Private Sub CDPFromDate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CDPFromDate.TextChanged
        txtCustItem.Text = ""
        lblCusttemName.Text = ""
        txtcustomerhelp.Text = ""
        lblCustName.Text = ""
        txtDocNo.Text = ""
        txtDocDate.Text = ""
    End Sub

    Private Sub CDPEndDate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CDPEndDate.TextChanged
        txtCustItem.Text = ""
        lblCusttemName.Text = ""
        txtcustomerhelp.Text = ""
        lblCustName.Text = ""
        txtDocNo.Text = ""
        txtDocDate.Text = ""

        dgvSODetail.Columns.Clear()
        dgvSODetail.Rows.Clear()
    End Sub
End Class