Imports System.Data.SqlClient
Imports Microsoft.Office.Interop
Imports System.IO
Imports Newtonsoft.Json
Public Class FRMMKTTRN0140
#Region "Constant & Variable"
    Private IntFormIndex As Integer
    Private StrTransType As String = "FRMMKTTRN0140"
    Private StrSuccessMsg As String
    Private DTExcel As System.Data.DataTable
    Private Enum EnumGrid
        ItemCode = 0
        Description
        CustPartNo
        InvoiceQty
        UploadedQty
    End Enum
    Private Enum EnumSunroofExcel
        ItemCode
        BoxID
        TracerID
    End Enum
    Private Enum EnumPFTExcel
        ItemCode
        TracerID
    End Enum
    Private Structure StructTransMode
        Friend Const SunRoof As String = "SunRoof"
        Friend Const PFT As String = "PFT"
        Friend Const Get_ManualIncoiceData As String = "GetManualInvoiceData"
        Friend Const SaveManualInoice As String = "SaveManualInvoice"
    End Structure
    Private Structure StructSunroofExcel
        Friend Const ItemCode As String = "Item Code"
        Friend Const BoxID As String = "Box ID"
        Friend Const TracerID As String = "Tracer ID"
    End Structure
    Private Structure StructPFTExcel
        Friend Const ItemCode As String = "Item Code"
        Friend Const TracerID As String = "Tracer ID"
    End Structure
    Private Structure StructExcelTemplate
        'Password of following excel template is ------> empro
        Friend Shared ReadOnly SunRoof As String = My.Application.Info.DirectoryPath & "\ManualInvoiceFRMMKTTRN0140_SunRoof.xltx"
        Friend Shared ReadOnly PFT As String = My.Application.Info.DirectoryPath & "\ManualInvoiceFRMMKTTRN0140_PFT.xltx"
    End Structure
#End Region
#Region "Controls"
    Private Sub FRMMKTTRN0140_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Me.Visible = False
            FitToClient(Me, pnlMain, ctlFormHeader, Panel2, 250)
            IntFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader.Tag)
            Me.MdiParent = mdifrmMain
            Call RefreshScreen()
            SetToolTip()
        Catch ex As Exception
            Dim StrMsg As String
            StrMsg = "FRMMKTTRN0140_Load" + vbCrLf + vbCrLf + ex.Message.ToString()
            MessageBox.Show(StrMsg, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Visible = True
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub FRMMKTTRN0140_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        txtCustomer_code.Focus()
    End Sub
    Private Sub txtCustomer_code_KeyDown(sender As Object, e As KeyEventArgs) Handles txtCustomer_code.KeyDown
        If e.KeyCode = Keys.F1 Then
            cmdCustomerHelp_Click(Nothing, Nothing)
        End If
    End Sub
    Private Sub cmdCustomerHelp_Click(sender As Object, e As EventArgs) Handles cmdCustomerHelp.Click
        Try
            Dim StrSql As String
            Dim strHelp() As String
            With ctlHelp
                .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                .ConnectAsUser = gstrCONNECTIONUSER
                .ConnectThroughDSN = gstrCONNECTIONDSN
                .ConnectWithPWD = gstrCONNECTIONPASSWORD
            End With

            StrSql = " SELECT Customer_Code , Description FROM dbo.UDF_CUSTOMER('" & gstrUNITID & "','','" & StrTransType & "','','','')"
            'dtpFromDate.Value.ToString("MM/dd/yyyy") & "','" & dtpToDate.Value.ToString("MM/dd/yyyy") & "')"
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, StrSql, "Help", 1)

            If UBound(strHelp) <> -1 Then
                If strHelp(0) <> "0" Then
                    Me.txtCustomer_code.Text = strHelp(0)
                    Me.txtCust_Desc.Text = strHelp(1)
                    txtInoiceNo.Text = ""
                    txtFilePath.Text = ""
                    GridPartDetail.DataSource = Nothing
                    txtInoiceNo.Focus()
                Else
                    MsgBox(" No record available", MsgBoxStyle.Critical, ResolveResString(100))
                    RefreshScreen()
                End If
            Else
                RefreshScreen()
            End If
        Catch ex As Exception
            Call RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen,, System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub txtInoiceNo_KeyDown(sender As Object, e As KeyEventArgs) Handles txtInoiceNo.KeyDown
        If e.KeyCode = Keys.F1 Then
            cmdInvoiceNoHelp_Click(Nothing, Nothing)
        End If
    End Sub
    Private Sub cmdInvoiceNoHelp_Click(sender As Object, e As EventArgs) Handles cmdInvoiceNoHelp.Click
        Try
            If txtCustomer_code.Text.Trim() = "" Then
                MsgBox("Please select Customer.", MsgBoxStyle.Critical, ResolveResString(100))
                txtCustomer_code.Focus()
                Return
            End If

            Dim StrSql As String
            Dim strHelp() As String
            With ctlHelp
                .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                .ConnectAsUser = gstrCONNECTIONUSER
                .ConnectThroughDSN = gstrCONNECTIONDSN
                .ConnectWithPWD = gstrCONNECTIONPASSWORD
            End With

            StrSql = " SELECT DOC_NO , DOC_DATE, CUSTOMER_CODE FROM dbo.[UFN_DOCNOHELP]('" & gstrUNITID & "','','" &
                      StrTransType & "','" & Trim(txtCustomer_code.Text) & "','','','','','','')"
            'dtpFromDate.Value.ToString("MM/dd/yyyy") & "','" & dtpToDate.Value.ToString("MM/dd/yyyy") & "')"
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, StrSql, "Help", 1)

            If UBound(strHelp) <> -1 Then
                If strHelp(0) <> "0" Then
                    Me.txtInoiceNo.Text = strHelp(0)
                    GridPartDetail.DataSource = Nothing
                    btnBrowseFile.Focus()
                Else
                    MsgBox(" No record available.", MsgBoxStyle.Critical, ResolveResString(100))
                    txtInoiceNo.Text = ""
                    GridPartDetail.DataSource = Nothing
                    txtInoiceNo.Focus()
                End If
            Else
                txtInoiceNo.Text = ""
                GridPartDetail.DataSource = Nothing
                txtInoiceNo.Focus()
            End If
        Catch ex As Exception
            Call RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen,, System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub btnExcelTemplateSunroof_Click(sender As Object, e As EventArgs) Handles btnExcelTemplateSunroof.Click
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen,, System.Windows.Forms.Cursors.WaitCursor)
            DownloadExcelTemplate(StructExcelTemplate.SunRoof)
        Catch ex As Exception
            Call RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen,, System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub btnExcelTemplatePFT_Click(sender As Object, e As EventArgs) Handles btnExcelTemplatePFT.Click
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen,, System.Windows.Forms.Cursors.WaitCursor)
            DownloadExcelTemplate(StructExcelTemplate.PFT)
        Catch ex As Exception
            Call RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen,, System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub btnBrowseFile_Click(sender As Object, e As EventArgs) Handles btnBrowseFile.Click
        Try
            OFDExcelFile.InitialDirectory = gstrLocalCDrive
            OFDExcelFile.Filter = "Microsoft Excel File  (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv"
            OFDExcelFile.FileName = ""
            OFDExcelFile.ShowDialog()
            txtFilePath.Text = OFDExcelFile.FileName
        Catch ex As Exception
            Call RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub btnUploadData_Click(sender As Object, e As EventArgs) Handles btnUploadData.Click
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen,, System.Windows.Forms.Cursors.WaitCursor)
            If ValidateOnUpload() Then
                DTExcel = ClsFlatFileDB.ReadFlatFileToDataTable(txtFilePath.Text.Trim(), "|"c, "", "")

                If ValidExcelTemplate() Then
                    Using DtGrid As DataTable = GetDataBase(StructTransMode.Get_ManualIncoiceData)
                        GridPartDetail.DataSource = DtGrid
                    End Using
                End If
            End If
        Catch ex As Exception
            Call RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen,, System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            If ValidationAll() Then
                If MsgBox("Are you sure?", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                    GetDataBase(StructTransMode.SaveManualInoice)
                    If StrSuccessMsg.Trim().Length > 0 Then
                        RefreshScreen()
                    End If
                End If
            End If
        Catch ex As Exception
            Call RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen,, System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        If MsgBox("Are you sure?", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
            RefreshScreen()
        Else
            txtCustomer_code.Focus()
        End If
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        If MsgBox("Are you sure?", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
            Me.Close()
            Me.Dispose()
        Else
            txtCustomer_code.Focus()
        End If
    End Sub
#End Region
#Region "Function & Method"
    Private Sub RefreshScreen()
        txtCustomer_code.Text = ""
        txtCust_Desc.Text = ""
        txtInoiceNo.Text = ""
        txtFilePath.Text = ""
        GridPartDetail.AutoGenerateColumns = False
        GridPartDetail.DataSource = Nothing
        txtCustomer_code.Focus()
    End Sub
    Private Sub SetToolTip()
        TTip.SetToolTip(btnBrowseFile, "File Should be closed.")
        TTip.SetToolTip(btnUploadData, "File Should be closed.")
    End Sub
    Private Sub PopulateGrid()

    End Sub
    Private Function ValidFile() As Boolean
        If txtFilePath.Text.Trim().Length < 1 Then
            MsgBox("Please select a Excel Data File.", MsgBoxStyle.Critical, ResolveResString(100))
            btnBrowseFile.Focus()
            Return False
        ElseIf Path.GetExtension(txtFilePath.Text.Trim()).ToLower() <> ".xlsx" Then
            MsgBox("Selected File should be in .xlsx format.", MsgBoxStyle.Critical, ResolveResString(100))
            btnBrowseFile.Focus()
            Return False
        End If

        Return True
    End Function
    Private Function ValidateOnUpload() As Boolean
        If txtCustomer_code.Text.Trim() = "" Then
            MsgBox("Customer Code should not be blank.", MsgBoxStyle.Critical, ResolveResString(100))
            txtCustomer_code.Focus()
            Return False
        ElseIf txtInoiceNo.Text.Trim() = "" Then
            MsgBox("Invoice No should not be blank.", MsgBoxStyle.Critical, ResolveResString(100))
            txtInoiceNo.Focus()
            Return False
        ElseIf txtFilePath.Text.Trim().Length < 1 Then
            MsgBox("Please select a Excel Data File.", MsgBoxStyle.Critical, ResolveResString(100))
            btnBrowseFile.Focus()
            Return False
        ElseIf Path.GetExtension(txtFilePath.Text.Trim()).ToLower() <> ".xlsx" Then
            MsgBox("Selected File should be in .xlsx format.", MsgBoxStyle.Critical, ResolveResString(100))
            btnBrowseFile.Focus()
            Return False
        End If

        Return True
    End Function
    Private Function ValidExcelTemplate() As Boolean
        If DTExcel Is Nothing Then
            MsgBox("Please Upload the selected File .", MsgBoxStyle.Critical, ResolveResString(100))
            btnUploadData.Focus()
            Return False
        ElseIf DTExcel.Rows.Count < 1 Then
            MsgBox("Selected File does not have data.", MsgBoxStyle.Critical, ResolveResString(100))
            btnBrowseFile.Focus()
            Return False
        ElseIf DTExcel.Columns.Count < 2 Or DTExcel.Columns.Count > 3 Then
            MsgBox("Selected File should have appropriate No of Columns(2 or 3).", MsgBoxStyle.Information, ResolveResString(100))
            btnBrowseFile.Focus()
            Return False
        ElseIf DTExcel.Columns.Count = 3 Then
            If DTExcel.Columns(EnumSunroofExcel.ItemCode).ColumnName <> StructSunroofExcel.ItemCode _
            Or DTExcel.Columns(EnumSunroofExcel.BoxID).ColumnName <> StructSunroofExcel.BoxID _
            Or DTExcel.Columns(EnumSunroofExcel.TracerID).ColumnName <> StructSunroofExcel.TracerID Then
                MsgBox("Selected File should be in appropriate template / format (columns heading).", MsgBoxStyle.Critical, ResolveResString(100))
                btnExcelTemplateSunroof.Focus()
                Return False
            End If
        ElseIf DTExcel.Columns.Count = 2 Then
            If DTExcel.Columns(EnumPFTExcel.ItemCode).ColumnName <> StructPFTExcel.ItemCode _
            Or DTExcel.Columns(EnumPFTExcel.TracerID).ColumnName <> StructPFTExcel.TracerID Then
                MsgBox("Selected File should be in appropriate template / format (columns heading).", MsgBoxStyle.Critical, ResolveResString(100))
                btnExcelTemplatePFT.Focus()
                Return False
            End If
        End If

        Return True
    End Function
    Private Function ValidationAll() As Boolean
        If ValidateOnUpload() = False Then
            Return False
        ElseIf ValidExcelTemplate() = False Then
            Return False
        End If
        Return True
    End Function
    Private Function GetDataBase(ByVal vStrTranMode As String) As DataTable
        Dim StrValidMessage As String
        Dim StrErrMessage As String
        Dim DtTmpReturn As DataTable

        Using oCmd As New SqlCommand() With {
            .Connection = SqlConnectionclass.GetConnection(),
            .CommandTimeout = 60,
            .CommandType = CommandType.StoredProcedure,
            .CommandText = "USP_YACHIYO_MANUAL_INVOICING"
        }
            With oCmd
                .Parameters.Clear()
                .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID.Trim())
                .Parameters.AddWithValue("@USERID", mP_User)
                .Parameters.AddWithValue("@TRAN_MODE", vStrTranMode)
                .Parameters.AddWithValue("@CUSTOMER_CODE", txtCustomer_code.Text.Trim())
                .Parameters.AddWithValue("@INVOICENO", txtInoiceNo.Text.Trim())
                .Parameters.AddWithValue("@JSON_STR", ConvertDataTableToJson(DTExcel))

                .Parameters.Add("@SUCCESS_MSG", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                .Parameters.Add("@VALIDATION_MSG", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                .Parameters.Add("@RET_ERROR_MSG", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output

                DtTmpReturn = SqlConnectionclass.GetDataTable(oCmd)

                StrErrMessage = Convert.ToString(.Parameters("@RET_ERROR_MSG").Value)
                StrValidMessage = Convert.ToString(.Parameters("@VALIDATION_MSG").Value)
                StrSuccessMsg = Convert.ToString(.Parameters("@SUCCESS_MSG").Value)
            End With
        End Using

        If StrErrMessage.Trim().Length() > 0 Then
            Throw New ArgumentException(StrErrMessage)
        ElseIf StrValidMessage.Length() > 0 Then
            MsgBox(StrValidMessage, MsgBoxStyle.Critical, ResolveResString(100))
        ElseIf StrSuccessMsg.Length() > 0 Then
            MsgBox(StrSuccessMsg, MsgBoxStyle.Information, ResolveResString(100))
        End If

        Return DtTmpReturn
    End Function
    Private Sub DownloadExcelTemplate(ByVal vStrFilePath As String)
        Dim ExcelApp As New Excel.Application
        Dim ExcelWorkBook As Excel.Workbook = ExcelApp.Workbooks.Open(vStrFilePath)

        ExcelApp.Visible = True

        ExcelApp = Nothing
        ExcelWorkBook = Nothing
    End Sub
    Private Function ConvertDataTableToJson(ByVal vDtJSON As DataTable) As String
        Dim DtJson As DataTable = vDtJSON

        If DtJson Is Nothing Then
            Return "[{}]".ToString()
        End If

        Dim lstJson As New List(Of Dictionary(Of String, String))
        For Each DtRow As DataRow In DtJson.Rows
            Dim DicJSON As New Dictionary(Of String, String)

            For Each DtCol As DataColumn In DtRow.Table.Columns
                Dim StrKeyName As String = DtCol.ColumnName
                Dim StrValue As String = DtRow(StrKeyName).ToString()

                DicJSON(StrKeyName) = StrValue
            Next

            lstJson.Add(DicJSON)
        Next

        Dim StrJson As String = JsonConvert.SerializeObject(lstJson, Formatting.Indented)
        Return StrJson
    End Function
#End Region
End Class