Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
'*********************************************************************************************************************
'Copyright(c)       - MIND
'Name of Module     - BSR Packing List Reprint
'Name of Form       - FRMMKTTRN0119  , BSR Packing List Reprint
'Created by         - Ashish sharma
'Created Date       - 13 OCT 2021
'description        - BSR Packing List Reprint
'*********************************************************************************************************************
Public Class FRMMKTTRN0119
    Private Enum GridPackingList
        Selection = 0
        PackingListNo
        GeneratedDate
        CustomerCode
        CustomerName
    End Enum

    Private Sub FRMMKTTRN0119_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            txtDocNo.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FRMMKTTRN0119_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Call FitToClient(Me, GrpMain, ctlHeader, GrpBoxButtons, 500)
            Me.MdiParent = mdifrmMain
            rdbPackingListWise.Checked = True
            ConfigureGridColumnns()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub ConfigureGridColumnns()
        Try
            dgvPackingDetail.Columns.Clear()
            Dim objChkBox As New DataGridViewCheckBoxColumn
            objChkBox.Name = "Selection"
            objChkBox.HeaderText = " "

            dgvPackingDetail.Columns.Add(objChkBox)
            dgvPackingDetail.Columns.Add("PackingListNo", "Packing List No.")
            dgvPackingDetail.Columns.Add("PackingListDate", "Generated Date")
            dgvPackingDetail.Columns.Add("CustomerCode", "Customer Code")
            dgvPackingDetail.Columns.Add("CustomerName", "Customer Name")

            dgvPackingDetail.Columns(GridPackingList.Selection).Width = 50
            dgvPackingDetail.Columns(GridPackingList.PackingListNo).Width = 115
            dgvPackingDetail.Columns(GridPackingList.GeneratedDate).Width = 120
            dgvPackingDetail.Columns(GridPackingList.CustomerCode).Width = 120
            dgvPackingDetail.Columns(GridPackingList.CustomerName).Width = 350

            dgvPackingDetail.Columns(GridPackingList.Selection).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvPackingDetail.Columns(GridPackingList.PackingListNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvPackingDetail.Columns(GridPackingList.GeneratedDate).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvPackingDetail.Columns(GridPackingList.CustomerCode).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvPackingDetail.Columns(GridPackingList.CustomerName).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            dgvPackingDetail.Columns(GridPackingList.PackingListNo).ReadOnly = True
            dgvPackingDetail.Columns(GridPackingList.GeneratedDate).ReadOnly = True
            dgvPackingDetail.Columns(GridPackingList.CustomerCode).ReadOnly = True
            dgvPackingDetail.Columns(GridPackingList.CustomerName).ReadOnly = True
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub rdbPackingListWise_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbPackingListWise.CheckedChanged
        Try
            If rdbPackingListWise.Checked Then
                lblDocNo.Text = "Packing List No."
                txtDocNo.Text = String.Empty
                dgvPackingDetail.Rows.Clear()
                txtDocNo.Focus()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub rdbInvoiceWise_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbInvoiceWise.CheckedChanged
        Try
            If rdbInvoiceWise.Checked Then
                lblDocNo.Text = "Invoice No."
                dgvPackingDetail.Rows.Clear()
                txtDocNo.Text = String.Empty
                txtDocNo.Focus()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub rdbCloseBoxWise_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbCloseBoxWise.CheckedChanged
        Try
            If rdbCloseBoxWise.Checked Then
                lblDocNo.Text = "Close Box No."
                dgvPackingDetail.Rows.Clear()
                txtDocNo.Text = String.Empty
                txtDocNo.Focus()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtDocNo_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDocNo.KeyPress
        Try
            If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) Then
                e.Handled = True
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            dgvPackingDetail.Rows.Clear()
            rdbPackingListWise.Checked = True
            rdbInvoiceCheckAll.Checked = False
            rdbInvoiceUncheckAll.Checked = False
            txtDocNo.Text = String.Empty
            txtDocNo.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub rdbInvoiceCheckAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbInvoiceCheckAll.CheckedChanged
        Try
            CheckUncheckInvoicesAll()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub CheckUncheckInvoicesAll()
        If dgvPackingDetail Is Nothing OrElse dgvPackingDetail.Rows.Count = 0 Then
            rdbInvoiceCheckAll.Checked = False
            rdbInvoiceUncheckAll.Checked = False
            Exit Sub
        End If
        With dgvPackingDetail
            For i As Integer = 0 To .Rows.Count - 1
                .Rows(i).Cells(GridPackingList.Selection).Value = rdbInvoiceCheckAll.Checked
            Next
        End With
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim dt As New DataTable
        Try
            Dim strMsg As String = String.Empty
            Dim strSearchOption As String = String.Empty
            If rdbPackingListWise.Checked Then
                strMsg = "Packing List No."
                strSearchOption = "P"
            ElseIf rdbInvoiceWise.Checked Then
                strMsg = "Invoice No."
                strSearchOption = "I"
            ElseIf rdbCloseBoxWise.Checked Then
                strMsg = "Close Box No."
                strSearchOption = "B"
            End If
            If String.IsNullOrEmpty(txtDocNo.Text) Then
                MsgBox("Please enter " & strMsg, MsgBoxStyle.Exclamation, ResolveResString(100))
                txtDocNo.Focus()
                Exit Sub
            End If

            Using sqlcmd As SqlCommand = New SqlCommand
                With sqlcmd
                    .CommandText = "USP_BSR_PACKING_LIST_REPRINT"
                    .CommandTimeout = 0
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Clear()
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUnitId
                    .Parameters.Add("@SERACH_OPTION", SqlDbType.VarChar, 1).Value = strSearchOption
                    .Parameters.Add("@DOC_NO", SqlDbType.VarChar, 18).Value = Convert.ToString(txtDocNo.Text)
                    .Parameters.Add("@USER_ID", SqlDbType.VarChar, 50).Value = mP_User
                    .Parameters.Add("@OPERATION", SqlDbType.VarChar, 50).Value = "SEARCH"
                    .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                    dt.Load(SqlConnectionclass.ExecuteReader(sqlcmd))
                    If Convert.ToString(.Parameters("@MESSAGE").Value) = String.Empty Then
                        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                            Dim i As Integer = 0
                            dgvPackingDetail.Rows.Add(dt.Rows.Count)
                            For Each dr As DataRow In dt.Rows
                                dgvPackingDetail.Rows(i).Cells(GridPackingList.Selection).Value = False
                                dgvPackingDetail.Rows(i).Cells(GridPackingList.PackingListNo).Value = dr("PACKING_LIST_NO")
                                dgvPackingDetail.Rows(i).Cells(GridPackingList.GeneratedDate).Value = dr("GENERATED_DATE")
                                dgvPackingDetail.Rows(i).Cells(GridPackingList.CustomerCode).Value = dr("CUSTOMER_CODE")
                                dgvPackingDetail.Rows(i).Cells(GridPackingList.CustomerName).Value = Convert.ToString(dr("CUSTOMER_NAME"))
                                i += 1
                            Next
                            dgvPackingDetail.Focus()
                            dgvPackingDetail.CurrentCell = dgvPackingDetail.Rows(0).Cells(GridPackingList.Selection)
                        Else
                            MsgBox("No record(s) found.", MsgBoxStyle.Exclamation, ResolveResString(100))
                            txtDocNo.Text = String.Empty
                            txtDocNo.Focus()
                        End If
                    Else
                        MsgBox(Convert.ToString(.Parameters("@MESSAGE").Value), MsgBoxStyle.Exclamation, ResolveResString(100))
                        txtDocNo.Text = String.Empty
                        txtDocNo.Focus()
                    End If
                End With
            End Using
        Catch ex As Exception
            RaiseException(ex)
        Finally
            If dt IsNot Nothing Then
                dt.Dispose()
            End If
        End Try
    End Sub

    Private Sub btnReprintPackingList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReprintPackingList.Click
        Try
            If dgvPackingDetail Is Nothing OrElse dgvPackingDetail.Rows.Count = 0 Then
                MsgBox("No Packing List found to print.", MsgBoxStyle.Exclamation, ResolveResString(100))
                txtDocNo.Focus()
                Exit Sub
            End If
            Dim flag As Boolean = False

            For i As Integer = 0 To dgvPackingDetail.Rows.Count - 1
                If Convert.ToBoolean(dgvPackingDetail.Rows(i).Cells(GridPackingList.Selection).Value) Then
                    flag = True
                    Exit For
                End If
            Next
            If Not flag Then
                MsgBox("Please select atleast one packing list to print.", MsgBoxStyle.Exclamation, ResolveResString(100))
                dgvPackingDetail.Focus()
                dgvPackingDetail.CurrentCell = dgvPackingDetail.Rows(i).Cells(GridPackingList.Selection)
                Exit Sub
            End If
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
            Dim packingListNo As String = String.Empty
            Dim blnRes As Boolean = False
            For i As Integer = 0 To dgvPackingDetail.Rows.Count - 1
                If Convert.ToBoolean(dgvPackingDetail.Rows(i).Cells(GridPackingList.Selection).Value) Then
                    packingListNo = Convert.ToString(dgvPackingDetail.Rows(i).Cells(GridPackingList.PackingListNo).Value)
                    Using sqlcmd As SqlCommand = New SqlCommand
                        With sqlcmd
                            .CommandText = "USP_BSR_PACKING_LIST_REPRINT"
                            .CommandTimeout = 0
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.Clear()
                            .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUnitId
                            .Parameters.Add("@PACKING_LIST_NO", SqlDbType.BigInt).Value = Convert.ToInt64(packingListNo)
                            .Parameters.Add("@USER_ID", SqlDbType.VarChar, 50).Value = mP_User
                            .Parameters.Add("@OPERATION", SqlDbType.VarChar, 50).Value = "REPRINT"
                            .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                            SqlConnectionclass.ExecuteNonQuery(sqlcmd)
                            If Convert.ToString(.Parameters("@MESSAGE").Value) <> String.Empty Then
                                MsgBox(Convert.ToString(.Parameters("@MESSAGE").Value), MsgBoxStyle.Exclamation, ResolveResString(100))
                            End If
                        End With
                    End Using
                    PrintPackingList(packingListNo, "R")
                    blnRes = True
                End If
            Next
            If blnRes Then
                MsgBox("Selected packing List(s) have been successfully printed.", MsgBoxStyle.Information, ResolveResString(100))
                dgvPackingDetail.Rows.Clear()
                txtDocNo.Text = String.Empty
                txtDocNo.Focus()
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub PrintPackingList(ByVal packingListNo As String, ByVal copyName As String)
        Try
            Dim Address As String = gstr_WRK_ADDRESS1 & gstr_WRK_ADDRESS2
            Dim objRpt As ReportDocument
            Dim frmReportViewer As New eMProCrystalReportViewer
            objRpt = frmReportViewer.GetReportDocument()
            objRpt.Load(My.Application.Info.DirectoryPath & "\Reports\rptBSRPackingList.rpt")
            objRpt.DataDefinition.FormulaFields("CompanyName").Text = "'" + gstrCOMPANY + "'"
            objRpt.DataDefinition.FormulaFields("CompanyAddress").Text = "'" + Address + "'"
            objRpt.DataDefinition.FormulaFields("COPY_NAME").Text = "'" + copyName + "'"
            objRpt.RecordSelectionFormula = "{BSR_PACKING_LIST_DTL.PACKING_LIST_NO}=" & packingListNo & " AND {BSR_PACKING_LIST_DTL.UNIT_CODE}='" & gstrUnitId & "'"
            frmReportViewer.ShowPrintButton = False
            frmReportViewer.ShowTextSearchButton = False
            frmReportViewer.ShowZoomButton = False
            frmReportViewer.SetReportDocument()
            objRpt.PrintToPrinter(1, False, 0, 0)
            'frmReportViewer.Show()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
End Class