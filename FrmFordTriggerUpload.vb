Imports System.IO
Imports System.Data.SqlClient
Imports ExcelAlias = Microsoft.Office.Interop.Excel

'*********************************************************************************************************************
'Copyright(c)       - MIND
'Name of Module     - Ford Trigger Upload
'Name of Form       - FrmFordTriggerUpload  , Ford Trigger Upload
'Created by         - Ashish sharma
'Created Date       - 11 Dec 2017
'description        - To upload Ford trigger manually (New Development)
'*********************************************************************************************************************

Public Class FrmFordTriggerUpload
#Region "Form Level Constant"
    Private Const HeaderRowIndex As Byte = 3
    Private Const DataRowIndex As Byte = 4
#End Region
#Region "Form Level Variable"
    Dim FordExcelColumnName As String() = {"Reader Date", "Reader Time", "Rotation Nbr", "Blend Number", _
                                          "Prod Sys", "VIN", "Vehicle Line", "Offline Date", "Event", _
                                          "Launch Pgm", "Build Level Date", "Req Date", "Line Feed", _
                                          "Part No", "Description", "Usage Qty", "Supplier", "CPSC", "Filter"}
    Dim dtUnits As DataTable
#End Region
#Region "Enum"
    Private Enum GridColumn
        Unit_Code = 0
        Seq_no
        VIN
        Supplier
        Part_No
    End Enum
    Private Enum FordExcelColumn
        Reader_Date = 1
        Reader_Time
        Rotation_Nbr
        Blend_Number
        Prod_Sys
        VIN
        Vehicle_Line
        Offline_Date
        Event_
        Launch_Pgm
        Build_Level_Date
        Req_Date
        Line_Feed
        Part_No
        Description
        Usage_Qty
        Supplier
        CPSC
        Filter
    End Enum
#End Region
#Region "Form Events"
    Private Sub FrmFordTriggerUpload_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        btnBrowse.Focus()
    End Sub

    Private Sub FrmFordTriggerUpload_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            dtUnits.Dispose()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FrmFordTriggerUpload_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Call FitToClient(Me, GrpMain, ctlHeader, GrpBoxButtons, 600)
            Me.MdiParent = mdifrmMain
            dtpTriggerDate.Value = GetServerDate()
            dtpTriggerDate.Enabled = False
            btnUpload.Enabled = False
            Call FillUnitCode()
            If lstUnitCode.Items.Count > 0 Then
                btnBrowse.Enabled = True
            Else
                btnBrowse.Enabled = False
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
#End Region
#Region "Private Methods"
    Private Sub FillUnitFromDB()
        Try
            dtUnits = New DataTable()
            dtUnits = SqlConnectionclass.GetDataTable("SELECT FIELD FROM DBO.SPLIT_STRING((SELECT DESCR FROM LISTS WHERE KEY1='FORD_TRG_UNIT' AND KEY2='UNIT' AND UNIT_CODE='" & gstrUnitId & "'),',')")
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub FillUnitCode()
        lstUnitCode.Items.Clear()
        If dtUnits Is Nothing OrElse dtUnits.Rows.Count = 0 Then
            FillUnitFromDB()
        End If
        lstUnitCode.DataSource = dtUnits
        lstUnitCode.DisplayMember = "FIELD"
        lstUnitCode.ValueMember = "FIELD"
    End Sub

    Private Sub ConfigureGrid()
        Try
            dgvUploadedRecords.Columns(GridColumn.Unit_Code).Width = 50
            dgvUploadedRecords.Columns(GridColumn.Seq_no).Width = 70
            dgvUploadedRecords.Columns(GridColumn.VIN).Width = 100
            dgvUploadedRecords.Columns(GridColumn.Supplier).Width = 60
            dgvUploadedRecords.Columns(GridColumn.Part_No).Width = 450

            dgvUploadedRecords.Columns(GridColumn.Unit_Code).HeaderCell.Style.Font = New Font(dgvUploadedRecords.Font, FontStyle.Bold)
            dgvUploadedRecords.Columns(GridColumn.Seq_no).HeaderCell.Style.Font = New Font(dgvUploadedRecords.Font, FontStyle.Bold)
            dgvUploadedRecords.Columns(GridColumn.VIN).HeaderCell.Style.Font = New Font(dgvUploadedRecords.Font, FontStyle.Bold)
            dgvUploadedRecords.Columns(GridColumn.Supplier).HeaderCell.Style.Font = New Font(dgvUploadedRecords.Font, FontStyle.Bold)
            dgvUploadedRecords.Columns(GridColumn.Part_No).HeaderCell.Style.Font = New Font(dgvUploadedRecords.Font, FontStyle.Bold)
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Function ValidateFordExcelFileColumn(ByVal fileData As Object(,)) As Boolean
        Dim result As Boolean = True
        Try
            Dim maxLengthFordExcelColumn As Integer = [Enum].GetValues(GetType(FordExcelColumn)).Length
            If fileData Is Nothing Then
                MsgBox("Browse File does not contain data.")
                result = False
                Return result
            End If
            If fileData.GetUpperBound(0) < DataRowIndex Then
                MsgBox("Browse File data is not in Ford Excel format, Please upload correct formatted file.")
                result = False
                Return result
            End If
            If fileData.GetUpperBound(1) < maxLengthFordExcelColumn Then
                MsgBox("Browse File data is not in Ford Excel format, Please upload correct formatted file.")
                result = False
                Return result
            End If
            For col As Integer = 1 To maxLengthFordExcelColumn
                If fileData(HeaderRowIndex, col).ToString.ToUpper.Trim <> FordExcelColumnName(col - 1).ToString.ToUpper.Trim Then
                    MsgBox("Browse File data is not in Ford Excel format, Please upload correct formatted file." & vbCrLf & "Excel file does not contain column : " & FordExcelColumnName(col - 1).ToString.ToUpper.Trim)
                    result = False
                    Return result
                End If
            Next
        Catch ex As Exception
            RaiseException(ex)
        End Try
        Return result
    End Function

    Private Function GetExcelFileDataIntoDatatable() As DataTable
        Dim dtFileData As New DataTable
        Dim xlApp As ExcelAlias.Application
        Dim xlWorkBook As ExcelAlias.Workbook
        Dim xlWorkSheet As ExcelAlias.Worksheet
        Try
            Dim maxLengthFordExcelColumn As Integer = [Enum].GetValues(GetType(FordExcelColumn)).Length

            xlApp = New ExcelAlias.ApplicationClass
            xlWorkBook = xlApp.Workbooks.Open(txtFilePath.Text)
            xlWorkSheet = xlWorkBook.ActiveSheet

            Dim data As Object(,) = DirectCast(xlWorkSheet.UsedRange.Value2, Object(,))

            If Not ValidateFordExcelFileColumn(data) Then

                txtFilePath.Text = String.Empty
                Return dtFileData
            End If

            For col As Integer = 1 To maxLengthFordExcelColumn
                If data(HeaderRowIndex, col).ToString.Trim.ToUpper = "READER DATE" Then
                    dtFileData.Columns.Add(data(HeaderRowIndex, col), GetType(System.DateTime))
                ElseIf data(HeaderRowIndex, col).ToString.Trim.ToUpper = "USAGE QTY" Then
                    dtFileData.Columns.Add(data(HeaderRowIndex, col), GetType(System.Decimal))
                Else
                    dtFileData.Columns.Add(data(HeaderRowIndex, col), GetType(System.String))
                End If
            Next

            For row As Integer = DataRowIndex To data.GetUpperBound(0)
                Dim newDataRow As DataRow = dtFileData.NewRow()
                For col As Integer = 1 To maxLengthFordExcelColumn
                    newDataRow(col - 1) = data(row, col)
                Next
                dtFileData.Rows.Add(newDataRow)
            Next
        Catch ex As Exception
            RaiseException(ex)
        Finally
            xlWorkBook.Close()
            xlApp.Quit()

            ReleaseComObject(xlApp)
            ReleaseComObject(xlWorkBook)
            ReleaseComObject(xlWorkSheet)
        End Try
        Return dtFileData
    End Function

    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Function GetUnitCodesFromListUnit() As String
        Try
            Dim unitCodes As String = String.Empty
            If dtUnits Is Nothing OrElse dtUnits.Rows.Count = 0 Then
                FillUnitFromDB()
            End If
            If dtUnits IsNot Nothing AndAlso dtUnits.Rows.Count > 0 Then
                For i As Integer = 0 To dtUnits.Rows.Count - 1
                    unitCodes += Convert.ToString(dtUnits.Rows(i)("FIELD")) & ";"
                Next
                If unitCodes.Length > 0 Then
                    unitCodes = unitCodes.TrimEnd(";")
                End If
            End If
            Return unitCodes
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function
#End Region
#Region "Form Controls Events"
    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        Try
            dgvUploadedRecords.DataSource = Nothing
            Dim fileExtension As String = String.Empty
            Dim validFileExtension As String = String.Empty
            Dim validSize As Integer = 0
            Dim byteOfFile As Long = 0
            Dim sql As String = String.Empty
            openFileDialog_1.InitialDirectory = "C:\"
            openFileDialog_1.CheckFileExists = True
            openFileDialog_1.CheckPathExists = True
            openFileDialog_1.DefaultExt = ".XLS"
            openFileDialog_1.Filter = "Files (*.XLS;*.XLSX)|*.XLS;*.XLSX"
            openFileDialog_1.FilterIndex = 1
            openFileDialog_1.RestoreDirectory = True
            openFileDialog_1.ReadOnlyChecked = True
            openFileDialog_1.ShowReadOnly = True
            If openFileDialog_1.ShowDialog() = DialogResult.OK Then
                txtFilePath.Text = openFileDialog_1.FileName
                fileExtension = Path.GetExtension(txtFilePath.Text)
                If String.IsNullOrEmpty(fileExtension) Then
                    MsgBox("Please select valid extension file.Valid file extensions are : .XLS;.XLSX")
                    txtFilePath.Text = String.Empty
                    Exit Sub
                End If
                If fileExtension.ToUpper() <> ".XLS" And fileExtension.ToUpper() <> ".XLSX" Then
                    MsgBox("Please select valid extension file.Valid file extensions are : .XLS;.XLSX")
                    txtFilePath.Text = String.Empty
                    Exit Sub
                End If
                If lstUnitCode.Items.Count > 0 Then
                    btnUpload.Enabled = True
                Else
                    btnUpload.Enabled = False
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        Dim dtFordFileData As New DataTable
        Dim dtUploadData As New DataTable
        Cursor = Cursors.WaitCursor
        Try
            If lstUnitCode.Items.Count = 0 Then
                MsgBox("Unit Code should be exist in list before upload data.")
                Exit Sub
            End If
            If String.IsNullOrEmpty(txtFilePath.Text) Then
                MsgBox("Please Browse Ford Trigger File.")
                Exit Sub
            End If
            dtpTriggerDate.Value = GetServerDate()
            dtFordFileData = GetExcelFileDataIntoDatatable()

            If dtFordFileData IsNot Nothing AndAlso dtFordFileData.Rows.Count > 0 Then
                Dim sqlCmd As New SqlCommand
                With sqlCmd
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 300 ' 5 Minute
                    .CommandText = "USP_FORD_TRIGGER_PREUPLOAD_MANUALLY"
                    .Parameters.Clear()
                    .Parameters.AddWithValue("@UDT_FORD_TRIGGER_PREUPLOAD", dtFordFileData)
                    .Parameters.AddWithValue("@TRG_DATE", dtpTriggerDate.Value)
                    .Parameters.AddWithValue("@UNIT_CODES", GetUnitCodesFromListUnit())
                    .Parameters.AddWithValue("@UPLOADED_FILE_NAME", Path.GetFileName(txtFilePath.Text))
                    .Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
                    .Parameters.AddWithValue("@USER_ID", mP_User)
                    .Parameters.Add("@MSG", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                    dtUploadData = SqlConnectionclass.GetDataTable(sqlCmd)
                    If Convert.ToString(.Parameters("@MSG").Value) <> "" Then
                        MsgBox(Convert.ToString(.Parameters("@MSG").Value), MsgBoxStyle.Exclamation, ResolveResString(100))
                        txtFilePath.Text = String.Empty
                        btnUpload.Enabled = False
                        btnBrowse.Focus()
                    Else
                        txtFilePath.Text = String.Empty
                        If dtUploadData IsNot Nothing AndAlso dtUploadData.Rows.Count > 0 Then
                            dgvUploadedRecords.DataSource = dtUploadData
                            ConfigureGrid()
                            btnUpload.Enabled = False
                        End If

                        MsgBox("Data Uploaded Successfully.")
                        btnBrowse.Focus()
                    End If
                End With
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            dtFordFileData.Dispose()
            Cursor = Cursors.Default
        End Try
    End Sub
#End Region
End Class