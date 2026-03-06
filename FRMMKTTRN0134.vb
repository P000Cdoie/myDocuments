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


Public Class FRMMKTTRN0134
    Dim intSave As Integer = 0
    Dim mintFormIndex As Integer
    Dim dtAuto As DataTable
    Dim blnCheckPackQtymultiple_onDA As Boolean = False
   
    Private Enum enmColItem
        Status = 0
        PDSNo
        SCH_DATE
        PDSTYPE
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
            SetItemGridsHeader()
            FillItemSearchCategory()
            'btnGrp1.ShowButtons(True, False, True, False)
            'EnableControls(False, Me, True)           
            RESETDATA()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub
#End Region

#Region "Methods"
   

    'Private Sub RefreshScreen()
    '    Try
    '        EnableControls(False, Me, True)
    '        txtDocNo.Enabled = True
    '        cmdDocHelp.Enabled = True
    '        txtDocNo.BackColor = Color.White
    '        sprPendingSchedule.MaxRows = 0
    '        sprPendingSchedule.Enabled = True
    '        txtDocNo.Enabled = True
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    
#End Region



    Private Sub cmdShowSchedule_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdShowPDS.Click
        ShowPDSDATA()
    End Sub

    Private Sub ShowPDSDATA()
        Try
            Dim StrSql As String = ""
            Dim i As Integer = 0
            If txtcustomerhelp.Text = "" Then
                MsgBox("Select Customer First", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
            Using sqlCmd1 As SqlCommand = New SqlCommand
                With sqlCmd1
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "USP_GETPDS_DATA_DELETION"
                    .CommandTimeout = 0
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@CUSTOMER_CODE", SqlDbType.VarChar, 8).Value = txtcustomerhelp.Text
                    .Parameters.Add("@FROMDATE", SqlDbType.Char, 11).Value = VB6.Format(Me.dtfromdate.Value, "dd/mmm/yyyy")
                    .Parameters.Add("@ToDATE", SqlDbType.Char, 11).Value = VB6.Format(Me.dttodate.Value, "dd/mmm/yyyy")
                    .Parameters.Add("@USER_ID", SqlDbType.VarChar, 20).Value = mP_User
                    .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 25).Value = gstrIpaddressWinSck
                    .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
                    If Convert.ToString(.Parameters("@MESSAGE").Value).Length > 0 Then
                        MessageBox.Show(Convert.ToString(.Parameters("@MESSAGE").Value), "eMPro", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                        Exit Sub
                    End If
                    SqlConnectionclass.ExecuteNonQuery(sqlCmd1)
                    SetItemGridsHeader()
                    StrSql = "select DISTINCT PDS_NUMBER,SCH_DATE,PICKLIST_TYPE from TMP_PDS_DELETE WHERE UNIT_CODE='" & gstrUNITID & "' and CUSTOMER_CODE='" & txtcustomerhelp.Text & "' and IP_ADDRESS='" & gstrIpaddressWinSck & "'"
                    Dim dt As DataTable = SqlConnectionclass.GetDataTable(StrSql)
                    If dt.Rows.Count > 0 Then
                        dgvItemDetail.Rows.Clear()
                        dgvItemDetail.Rows.Add(dt.Rows.Count)
                        For Each dr As DataRow In dt.Rows
                            dgvItemDetail.Rows(i).Cells(enmColItem.Status).Value = False
                            dgvItemDetail.Rows(i).Cells(enmColItem.PDSNo).Value = dr("PDS_NUMBER")
                            dgvItemDetail.Rows(i).Cells(enmColItem.SCH_DATE).Value = dr("SCH_DATE")
                            dgvItemDetail.Rows(i).Cells(enmColItem.PDSTYPE).Value = dr("PICKLIST_TYPE")
                            i += 1
                        Next
                    End If

                End With
            End Using
        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub
    Private Sub SetItemGridsHeader()
        Try

            dgvItemDetail.Columns.Clear()

            Dim objChkBox As New DataGridViewCheckBoxColumn
            objChkBox.Name = "Selection"
            objChkBox.HeaderText = "Select"

            dgvItemDetail.Columns.Add(objChkBox)
            dgvItemDetail.Columns.Add("PDSNo", "PDS No")
            dgvItemDetail.Columns.Add("SCH_DATE", "SCH DATE")
            dgvItemDetail.Columns.Add("PDSTYPE", "INVOICE CREATED")
          
            dgvItemDetail.Columns(enmColItem.Status).Width = 50
            dgvItemDetail.Columns(enmColItem.PDSNo).Width = 150
            dgvItemDetail.Columns(enmColItem.SCH_DATE).Width = 150
            dgvItemDetail.Columns(enmColItem.PDSTYPE).Width = 150
           
            dgvItemDetail.Columns(enmColItem.Status).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvItemDetail.Columns(enmColItem.PDSNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvItemDetail.Columns(enmColItem.SCH_DATE).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvItemDetail.Columns(enmColItem.PDSTYPE).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
           
            dgvItemDetail.Columns(enmColItem.PDSNo).ReadOnly = True
            dgvItemDetail.Columns(enmColItem.SCH_DATE).ReadOnly = True
            dgvItemDetail.Columns(enmColItem.PDSTYPE).ReadOnly = True


            dgvItemDetail.Columns(enmColItem.Status).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvItemDetail.Columns(enmColItem.PDSNo).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvItemDetail.Columns(enmColItem.SCH_DATE).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvItemDetail.Columns(enmColItem.PDSTYPE).SortMode = DataGridViewColumnSortMode.NotSortable

        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    Private Sub cmdcusthelp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdcusthelp.Click
        Try
            Dim strCustHelp() As String
            Dim strsql As String
            strCustHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select customer_Code as CustomerCode, cust_name as CustomerName from customer_mst b(nolock)  " & _
                "where b.unit_code='" & gstrUNITID & "'  and isnull(PDSDelete,0)=1 and ((isnull(b.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= b.deactive_date)) ", "Customer Codes List", 1)
            If UBound(strCustHelp) <> -1 Then
                If strCustHelp(0) <> "0" Then
                    txtcustomerhelp.Text = Trim(strCustHelp(0))
                    lblCustName.Text = strCustHelp(1)
                    SetItemGridsHeader()
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
            RESETDATA()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub RESETDATA()
        Try
            txtcustomerhelp.Text = ""
            lblCustName.Text = ""
            OptSequencePDS.Checked = True
            dtfromdate.Format = DateTimePickerFormat.Custom
            dtfromdate.CustomFormat = gstrDateFormat
            dtfromdate.Value = GetServerDate()
            dttodate.Format = DateTimePickerFormat.Custom
            dttodate.CustomFormat = gstrDateFormat
            dttodate.Value = GetServerDate()
            SetItemGridsHeader()
        Catch ex As Exception
            RaiseException(ex)
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
    Private Sub txtItemsearch_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtItemsearch.TextChanged
        Dim intCounter As Integer
        Dim strText As String
        Dim Col As enmColItem
        Try
            If Len(txtItemsearch.Text) = 0 Then Exit Sub
            For intCounter = 0 To dgvItemDetail.Rows.Count - 1
                If cmbItem.Text = "PDSNo" Then
                    strText = dgvItemDetail.Rows(intCounter).Cells(enmColItem.PDSNo).Value
                    If Trim(UCase(Mid(strText, 1, Len(txtItemsearch.Text)))) = Trim(UCase(txtItemsearch.Text)) Then
                        dgvItemDetail.CurrentCell = dgvItemDetail.Rows(intCounter).Cells(enmColItem.PDSNo)
                        Exit For
                    End If
                    'ElseIf cmbItem.Text = "PicklistNo" Then
                    '    strText = dgvItemDetail.Rows(intCounter).Cells(enmColItem.PicklistNo).Value
                    '    If Trim(UCase(Mid(strText, 1, Len(txtItemsearch.Text)))) = Trim(UCase(txtItemsearch.Text)) Then
                    '        dgvItemDetail.CurrentCell = dgvItemDetail.Rows(intCounter).Cells(enmColItem.PicklistNo)
                    '        Exit For
                    '    End If
                End If


            Next
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            Dim Count As Integer = 0
            Dim strPDSNo As String = ""
            Dim strPicklistNo As String = ""
            Dim strPicklistType As String = ""
            Dim intCounter As Integer
            Dim ds As New DataSet

            If dgvItemDetail.Rows.Count = 0 Then
                MsgBox("There is no PDS for Delete", MsgBoxStyle.Information, Me.Text)
                Exit Sub
            ElseIf dgvItemDetail.Rows.Count > 0 Then
                For intCounter = 0 To dgvItemDetail.Rows.Count - 1
                    If dgvItemDetail.Rows(intCounter).Cells(enmColItem.Status).Value = True Then
                        Count = 1
                        Exit For
                    End If
                Next
                If Count = 0 Then
                    MsgBox("There is no PDS for Delete", MsgBoxStyle.Information, Me.Text)
                    Exit Sub
                End If
            End If
            Dim strMessage As String = ""
            Dim strMessageString As String = ""
            For intCounter = 0 To dgvItemDetail.Rows.Count - 1
                If dgvItemDetail.Rows(intCounter).Cells(enmColItem.Status).Value = True Then
                    strPDSNo = dgvItemDetail.Rows(intCounter).Cells(enmColItem.PDSNo).Value
                    'strPicklistNo = dgvItemDetail.Rows(intCounter).Cells(enmColItem.PicklistNo).Value
                    'strPicklistType = dgvItemDetail.Rows(intCounter).Cells(enmColItem.PickListType).Value
                    Using sqlCmd As SqlCommand = New SqlCommand
                        With sqlCmd
                            .CommandType = CommandType.StoredProcedure
                            .CommandText = "USP_DELETE_PDSDATA"
                            .CommandTimeout = 0
                            .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                            .Parameters.Add("@CUSTOMER_CODE", SqlDbType.VarChar, 8).Value = txtcustomerhelp.Text
                            .Parameters.Add("@PDS_NUMBER", SqlDbType.VarChar, 50).Value = strPDSNo
                            .Parameters.Add("@USER_ID", SqlDbType.VarChar, 20).Value = mP_User
                            .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 25).Value = gstrIpaddressWinSck
                            .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                            SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                            strMessage = Convert.ToString(.Parameters("@MESSAGE").Value)
                            If Convert.ToString(.Parameters("@MESSAGE").Value) <> "" Then
                                strMessage = Convert.ToString(.Parameters("@MESSAGE").Value)
                            End If
                            If strMessage <> "" Then
                                If strMessageString = "" Then
                                    strMessageString = "Below PDS are not deleted .It has been used in scanning or Invoice entry " & vbCrLf & strMessage
                                Else
                                    strMessageString += vbCrLf & strMessage
                                End If
                            End If
                        End With
                    End Using
                End If

            Next
            If strMessageString <> "" Then
                MsgBox(strMessageString, MsgBoxStyle.Information, Me.Text)
            Else
                MsgBox("PDS deleted successfully", MsgBoxStyle.Information, Me.Text)
                ShowPDSDATA()
            End If


        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub OptSequencePDS_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles OptSequencePDS.CheckedChanged, OptNonSequence.CheckedChanged
        Try
            If OptSequencePDS.Checked Or OptNonSequence.Checked Then
                txtcustomerhelp.Text = ""
                lblCustName.Text = ""
                dtfromdate.Format = DateTimePickerFormat.Custom
                dtfromdate.CustomFormat = gstrDateFormat
                dtfromdate.Value = GetServerDate()
                dttodate.Format = DateTimePickerFormat.Custom
                dttodate.CustomFormat = gstrDateFormat
                dttodate.Value = GetServerDate()
                SetItemGridsHeader()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
End Class