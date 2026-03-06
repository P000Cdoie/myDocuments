'**********************************************************************************************
'COPYRIGHT(C)   : MIND   
'FORM NAME      : FRMMKTTRN0141
'DESCRIPTION    : Upload Mahindra Daily Stock File
'CREATED BY     : TEK CHAND
'CREATED DATE   : 18 JAN 2026
'PURPOSE        : Uploading Mahindra Daily Stock File to showing real time data on dashboard
'**********************************************************************************************
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports VB = Microsoft.VisualBasic
Imports System.Collections.Generic
Imports ExcelDataReader


Public Class FRMMKTTRN0141
    Dim mintFormIndex As Integer
    Dim mConnString As String = ""
    Dim mFILE_NAME As String

    Private Sub FRMMKTTRN0141_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Me.MdiParent = mdifrmMain
            mintFormIndex = mdifrmMain.AddFormNameToWindowList(Me.ctlHeader.Tag)
            mConnString = "Server=" & gstrCONNECTIONSERVER & ";Database=" & gstrDatabaseName & ";user=" & gstrCONNECTIONUSER & ";password=" & gstrCONNECTIONPASSWORD & " "
        Catch ex As Exception

        End Try
    End Sub

    Private Sub FRMMKTTRN0141_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            Me.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub FRMMKTTRN0141_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress
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

    Private Sub cmdOpenFile_Click(sender As Object, e As EventArgs) Handles cmdOpenFile.Click
        Try
            Using ofd As New OpenFileDialog()
                ofd.Title = "Select Excel File"
                ofd.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls"
                ofd.Multiselect = False

                If ofd.ShowDialog() = DialogResult.OK Then
                    txtFilePath.Text = ofd.FileName
                End If
            End Using

        Catch ex As Exception

        End Try
    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Private Function ReadExcelToDataTable(filePath As String) As DataTable

        ' Required for ExcelDataReader
        'System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance)

        Dim dt As New DataTable()

        Using stream As FileStream = File.Open(filePath, FileMode.Open, FileAccess.Read)

            Dim reader As IExcelDataReader

            If Path.GetExtension(filePath).ToLower() = ".xls" Then
                reader = ExcelReaderFactory.CreateBinaryReader(stream)
            Else
                reader = ExcelReaderFactory.CreateOpenXmlReader(stream)
            End If

            Dim conf As New ExcelDataSetConfiguration With {
            .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration With {
                .UseHeaderRow = True
            }
        }

            Dim ds As DataSet = reader.AsDataSet(conf)
            reader.Close()

            ' Take first sheet
            dt = ds.Tables(0)
            If HasBlankCell(dt) Then
                MessageBox.Show("Excel contains blank cells.", "Error")
                Exit Function
            End If
            dt.Columns.Add("UploadID", GetType(Integer))
        End Using

        Return dt
    End Function

    Private Sub InsertIntoSql(dt As DataTable)


        Dim fileInfo As FileInfo = New FileInfo(txtFilePath.Text.Trim())
        Dim fileName As String = fileInfo.Name
        Dim modifiedDate As DateTime = fileInfo.LastWriteTime

        Using con As New SqlConnection(mConnString)
            con.Open()
            Dim tran As SqlTransaction = con.BeginTransaction()
            Try
                Dim cmdMaster As New SqlCommand("INSERT INTO MahindraStockFile_Log (FileName, UploadDate)
                                  VALUES (@FileName, @FileDateTime);
                                  SELECT CAST(SCOPE_IDENTITY() AS INT);", con, tran)

                cmdMaster.Parameters.AddWithValue("@FileName", fileName)
                cmdMaster.Parameters.AddWithValue("@FileDateTime", modifiedDate)
                Dim uploadId As Integer = CInt(cmdMaster.ExecuteScalar())

                'add by tek
                For Each row As DataRow In dt.Rows
                    row("UploadID") = uploadId   'value from master table
                Next
                'end

                'Using bulk As New SqlBulkCopy(con)
                Using bulk As New SqlBulkCopy(con, SqlBulkCopyOptions.Default, tran)
                    bulk.DestinationTableName = "MahindraStockFile"

                    ' Column mapping (Excel → SQL)
                    'For Each col As DataColumn In dt.Columns
                    '    bulk.ColumnMappings.Add(col.ColumnName, col.ColumnName)
                    'Next
                    bulk.ColumnMappings.Add("S.No", "Sr_No")
                    bulk.ColumnMappings.Add("Date", "Date")
                    bulk.ColumnMappings.Add("Model", "Model")
                    bulk.ColumnMappings.Add("Item Code", "Item_Code")
                    bulk.ColumnMappings.Add("Customer Drg no.", "CustDrgNo")
                    bulk.ColumnMappings.Add("Item Description", "Item_Description")
                    bulk.ColumnMappings.Add("Qty", "Qty")
                    bulk.ColumnMappings.Add("Unit Code", "Unit_Code")
                    bulk.ColumnMappings.Add("UploadID", "UploadID")
                    bulk.WriteToServer(dt)
                End Using
                tran.Commit()
            Catch ex As Exception
                tran.Rollback()
                MessageBox.Show(ex.Message)
                Return
            End Try
        End Using
    End Sub


    Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
        Try
            If txtFilePath.Text = "" Then
                MessageBox.Show("Please select Excel file")
                Exit Sub
            End If

            Dim dt As DataTable = ReadExcelToDataTable(txtFilePath.Text)



            InsertIntoSql(dt)

            MessageBox.Show("Excel data inserted successfully")
            txtFilePath.Text = ""
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Function HasBlankCell(dt As DataTable) As Boolean
        For Each row As DataRow In dt.Rows
            For Each col As DataColumn In dt.Columns
                If IsDBNull(row(col)) OrElse String.IsNullOrWhiteSpace(row(col).ToString()) Then
                    Return True
                End If
            Next
        Next
        Return False
    End Function
    Private Sub btnDownloadTemplate_Click(sender As Object, e As EventArgs) Handles btnDownloadTemplate.Click
        Dim templatePath As String = Application.StartupPath & "\Mahindra_Stock_file.xlsx"

        Dim newFilePath As String =
            Application.StartupPath & "\NewFile_" &
            DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx"

        If IO.File.Exists(templatePath) Then


            Using sfd As New SaveFileDialog()
                sfd.Filter = "Excel Files (*.xlsx)|*.xlsx"
                sfd.FileName = "ExcelTemplate_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx"

                If sfd.ShowDialog() = DialogResult.OK Then
                    IO.File.Copy(templatePath, sfd.FileName, True)
                    MessageBox.Show("Excel file downloaded successfully.", "Success",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        Else
            MessageBox.Show("Template not found!")
        End If
    End Sub
End Class