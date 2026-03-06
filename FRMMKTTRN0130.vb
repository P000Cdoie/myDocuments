
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports Mind
Imports System.ComponentModel
Imports System.Security.Cryptography
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports System.Net.Http
Imports System.Net.Http.Formatting
Imports Newtonsoft.Json
Imports System.Text.RegularExpressions
Imports System.Web
Imports System.Web.Http
Imports System.Net
Imports System.IO
Imports System.Configuration
Imports Mind.MailFormat
Imports System
Imports System.Security.Authentication




Public Class FRMMKTTRN0130
    Dim mintIndex As Short

    'Partial Public Class FRMMKTTRN0130
    '    Inherits Form


    Public Property SqlDataReaderClass() As Object
        Get
        End Get
        Set(ByVal value As Object)

        End Set
    End Property

    Private LogPath As String = ConfigurationManager.AppSettings("LogPath")


    Public Class ToDoItem
        Private _status_cd As String
        Public Property status_cd() As String
            Get
                Return _status_cd
            End Get
            Set(ByVal value As String)
                _status_cd = value
            End Set
        End Property


        Private _message As String
        Public Property message() As String
            Get
                Return _message
            End Get
            Set(ByVal value As String)
                _message = value
            End Set
        End Property

        Private _error_details As Object
        Public Property error_details() As Object
            Get
                Return _error_details
            End Get
            Set(ByVal value As Object)
                _error_details = value
            End Set
        End Property

        Private _error_cd As String
        Public Property error_cd() As String
            Get
                Return _error_cd
            End Get
            Set(ByVal value As String)
                _error_cd = value
            End Set
        End Property
    End Class


    Public Class ToDoItem2
        Private _gstin As String
        Public Property gstin() As String
            Get
                Return _gstin
            End Get
            Set(ByVal value As String)
                _gstin = value
            End Set
        End Property

        Private _unit_code As String
        Public Property unit_code() As String
            Get
                Return _unit_code
            End Get
            Set(ByVal value As String)
                _unit_code = value
            End Set
        End Property

        Private _error_message As String
        Public Property error_message() As String
            Get
                Return _error_message
            End Get
            Set(ByVal value As String)
                _error_message = value
            End Set
        End Property
    End Class

    Public Class IrnInvData
        Private _fileuploadcode As String
        Public Property fileuploadcode() As String
            Get
                Return _fileuploadcode
            End Get
            Set(ByVal value As String)
                _fileuploadcode = value
            End Set
        End Property

        Private _txn_id As String
        Public Property txn_id() As String
            Get
                Return _txn_id
            End Get
            Set(ByVal value As String)
                _txn_id = value
            End Set
        End Property

        Private _fileuploadstatus As String
        Public Property fileuploadstatus() As String
            Get
                Return _fileuploadstatus
            End Get
            Set(ByVal value As String)
                _fileuploadstatus = value
            End Set
        End Property
    End Class


    Dim dtlvar As String
    Dim dtlvar2 As String
    Dim col1 As String
    Dim col2 As String


    Private Sub FRMMKTTRN0130_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Call FitToClient(Me, grpMain, ctlFormHeader1, frmgrpbtn)
        Me.MdiParent = mdifrmMain
        Me.Icon = mdifrmMain.Icon
        mintIndex = mdifrmMain.AddFormNameToWindowList("Invoice Reprocess In Portal")

        Dim clientID As String
        Dim baseAddress As String
        Dim myUri As Uri
        Dim dtUnitAndClientIdDetails As DataTable
        dtUnitAndClientIdDetails = GetUnitAndclient_id()


        'For Each DataRow row In dtUnitAndClientIdDetails.Rows
        '    If row["baseAddress"].ToString().Trim() = "" Then
        '        Continue For
        '    Else

        '    End If
        'Next

    End Sub

    Private Sub BtnInvShow_Click_1(ByVal sender As Object, ByVal e As EventArgs) Handles btnInvShow.Click
        Try
            Dim ds As New DataSet
            Dim dt As New DataTable
            Dim dt2 As New DataTable
            dt = Nothing
            Dim sqlQuery As String
            Dim sqlQuery1 As String
            sqlQuery = Nothing
            sqlQuery1 = Nothing

            Dim difference As Integer = (Date.Today - dtFromDate.Value.Date).Days
            If difference > 1 Then
                MessageBox.Show("Please select Today or Yesterday as [From date ]")
                Return
            End If
           
            If rdobtnInv.Checked = True Then

                If dtFromDate.Value <> Nothing And dtToDate.Value <> Nothing Then


                    Dim fromDate As String = dtFromDate.Value.ToString("yyyy-MM-dd")
                    Dim toDate As String = dtToDate.Value.ToString("yyyy-MM-dd")
                    dtlvar = " GEN_EGSP_IRN "
                    col1 = "IRN_Invoice_ID"
                    dtlvar2 = " SALESCHALLAN_DTL_IRN "
                    col2 = "Doc_No"
                    ' sqlQuery = "SELECT CAST(0 AS BIT) AS [Select],IRN_INVOICE_ID AS [INVNO],* FROM GEN_EGSP_IRN WHERE IS_INV_DRCR = 'INV' AND CONVERT(DATE, IRN_CreatedOn , 103) >= '" & fromDate & "' AND CONVERT(DATE, IRN_CreatedOn , 103) <= '" & toDate & "' AND IRN_Invoice_Unit = '" & gstrUNITID & "'"


                    sqlQuery = "SELECT CAST(0 AS BIT) AS [Select], IRN_INVOICE_ID AS [INV NO], IRN_Invoice_Date AS [IRN INVOICE DATE], " & _
                        "IRN_GSTIN_ID AS [IRN GSTIN ID], IRN_JSON AS [IRN JSON], IRN_Status AS [IRN STATUS], " & _
                        "IRN_Remarks AS [IRN Remarks], IRN_TRAN_NO AS [IRN TRAN NO], IRN_CreatedOn AS [IRN CreatedOn], " & _
                        "FILEUPLOADCODE AS [FILE UPLOAD CODE] " & _
                        "FROM GEN_EGSP_IRN EGSP " & _
                        "WHERE IS_INV_DRCR = 'INV' " & _
                        "AND CONVERT(DATE, IRN_CreatedOn, 103) >= '" & fromDate & "' " & _
                        "AND CONVERT(DATE, IRN_CreatedOn, 103) <= '" & toDate & "' " & _
                        "AND IRN_Invoice_Unit = '" & gstrUNITID & "' " & _
                        "AND NOT EXISTS (select top 1 1 from FIRSTTIME_INVOICEPRINTING P where P.unit_code=EGSP.IRN_INVOICE_UNIT AND P.doc_no=EGSP.IRN_INVOICE_ID) Order by IRN_INVOICE_ID DESC "

                    'sqlQuery1 = "SELECT CAST(0 AS BIT) AS [Select],Doc_No AS [INVNO],* FROM SALESCHALLAN_DTL_IRN WHERE CONVERT(DATE, EWAY_UPD_DT  , 103) >= '" & fromDate & "' AND CONVERT(DATE, EWAY_UPD_DT  , 103) <= '" & toDate & "' AND Location_Code = '" & gstrUNITID & "'"      
                    sqlQuery1 = "SELECT CAST(0 AS BIT) AS [Select],Doc_No AS [INV NO],EWAY_BILL_NO AS [EWAY BILL NO], " & _
                          "EWAY_TRANSPORT_MODE AS [EWAY TRANSPORT MODE],EWAY_TRANSPORTER_ID AS [EWAY TRANSPORTER ID], " & _
                          "EWAY_TRANSPORTER_DOC_NO AS [EWAY TRANSPORTER DOC_NO],EWAY_TRANSPORTER_DOC_DATE AS [EWAY TRANSPORTER DOC DATE]," & _
                          "EWAY_VEHICLE_NO AS [EWAY VEHICLE NO], EWAY_UPD_DT AS [EWAY UPD DT], EWAY_UPD_USERID AS [EWAY_UPD_USERID], " & _
                          "EWAY_EXP_DT AS [ EWAY EXP DT],EWAY_UPD_DT_EMPRO AS [EWAY UPD_DT EMPRO], EWAY_UPD_DT_PARTB AS [EWAY UPD_DT PARTB]," & _
                          "IRN_NO AS [IRN NO],  IRN_DATE AS [IRN DATE],IRN_DATE_EMPRO AS [IRN DATE EMPRO]," & _
                          "IRN_DEACTIVATE AS [IRN DEACTIVATE],IRN_DEACTIVATE_DATE AS [IRN DEACTIVATE DATE]," & _
                          "ACK_NO AS [ACK NO],ACK_DT AS [ACK DT], CONEWBNO AS [CONEWB NO],CONEWBDT AS [CONEWB DT], " & _
                          "EINVOICE_STATUS AS[E-INVOICE STATUS],ERROR_MSG AS [ERROR_MSG]" & _
                          "FROM SALESCHALLAN_DTL_IRN EGSP " & _
                         " WHERE CONVERT(DATE, EWAY_UPD_DT, 103) >= '" & fromDate & "' " & _
                         "AND CONVERT(DATE, EWAY_UPD_DT, 103) <= '" & toDate & "' " & _
                         "AND Location_Code = '" & gstrUNITID & "' " & _
                         "AND NOT EXISTS (SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING P WHERE P.UNIT_CODE=EGSP.UNIT_CODE AND P.DOC_NO=EGSP.DOC_NO) ORDER BY EGSP.DOC_NO DESC "

                End If

            ElseIf rdobtnSplymntryInv.Checked = True Then

                If dtFromDate.Value <> Nothing And dtToDate.Value <> Nothing Then
                    Dim fromDate As String = dtFromDate.Value.ToString("yyyy-MM-dd")
                    Dim toDate As String = dtToDate.Value.ToString("yyyy-MM-dd")
                    dtlvar = " GEN_EGSP_IRN "
                    col1 = "IRN_Invoice_ID"
                    dtlvar2 = " Supplementary_IRN "
                    col2 = "Doc_No"


                    ' sqlQuery = "SELECT CAST(0 AS BIT) AS [Select],IRN_INVOICE_ID AS [INVNO],* FROM GEN_EGSP_IRN WHERE IS_INV_DRCR = 'SUP' AND CONVERT(DATE, IRN_CreatedOn , 103) >= '" & fromDate & "' AND CONVERT(DATE, IRN_CreatedOn , 103) <= '" & toDate & "' AND IRN_Invoice_Unit = '" & gstrUNITID & "'"
                    sqlQuery = "SELECT CAST(0 AS BIT) AS [Select], IRN_INVOICE_ID AS [INV NO], IRN_Invoice_Date AS [IRN INVOICE DATE], " & _
                       "IRN_GSTIN_ID AS [IRN GSTIN ID], IRN_JSON AS [IRN JSON], IRN_Status AS [IRN STATUS], " & _
                       "IRN_Remarks AS [IRN Remarks], IRN_TRAN_NO AS [IRN TRAN NO], IRN_CreatedOn AS [IRN CreatedOn], " & _
                       "FILEUPLOADCODE AS [FILE UPLOAD CODE] " & _
                       "FROM GEN_EGSP_IRN " & _
                       "WHERE IS_INV_DRCR = 'SUP' " & _
                       "AND CONVERT(DATE, IRN_CreatedOn, 103) >= '" & fromDate & "' " & _
                       "AND CONVERT(DATE, IRN_CreatedOn, 103) <= '" & toDate & "' " & _
                       "AND IRN_Invoice_Unit = '" & gstrUNITID & "'"


                    ' sqlQuery1 = "SELECT CAST(0 AS BIT) AS [Select],Doc_No AS [INVNO],*  FROM Supplementary_IRN WHERE CONVERT(DATE, EWAY_UPD_DT  , 103) >= '" & fromDate & "' AND CONVERT(DATE, EWAY_UPD_DT  , 103) <= '" & toDate & "' AND Location_Code = '" & gstrUNITID & "'"
                    sqlQuery1 = "SELECT CAST(0 AS BIT) AS [Select],VO_NO AS [INV NO],EWAY_BILL_NO AS [EWAY BILL NO], " & _
                           "EWAY_TRANSPORT_MODE AS [EWAY TRANSPORT MODE],EWAY_TRANSPORTER_ID AS [EWAY TRANSPORTER ID], " & _
                           "EWAY_TRANSPORTER_DOC_NO AS [EWAY TRANSPORTER DOC_NO],EWAY_TRANSPORTER_DOC_DATE AS [EWAY TRANSPORTER DOC DATE]," & _
                           "EWAY_VEHICLE_NO AS [EWAY VEHICLE NO], EWAY_UPD_DT AS [EWAY UPD DT], EWAY_UPD_USERID AS [EWAY_UPD_USERID], " & _
                           "EWAY_EXP_DT AS [ EWAY EXP DT],EWAY_UPD_DT_EMPRO AS [EWAY UPD_DT EMPRO], EWAY_UPD_DT_PARTB AS [EWAY UPD_DT PARTB]," & _
                           "IRN_NO AS [IRN NO],  IRN_DATE AS [IRN DATE],IRN_DATE_EMPRO AS [IRN DATE EMPRO]," & _
                           "IRN_DEACTIVATE AS [IRN DEACTIVATE],IRN_DEACTIVATE_DATE AS [IRN DEACTIVATE DATE]," & _
                           "ACK_NO AS [ACK NO],ACK_DT AS [ACK DT], CONEWBNO AS [CONEWB NO],CONEWBDT AS [CONEWB DT], " & _
                           "EINVOICE_STATUS AS[E-INVOICE STATUS],ERROR_MSG AS [ERROR_MSG]" & _
                           "FROM Supplementary_IRN " & _
                          " WHERE CONVERT(DATE, ENT_DT, 103) >= '" & fromDate & "' " & _
                          "AND CONVERT(DATE, ENT_DT, 103) <= '" & toDate & "' " & _
                          "AND Location_Code = '" & gstrUNITID & "'"


                End If
            ElseIf rdobtnCdt_Dbt.Checked = True Then

                If dtFromDate.Value <> Nothing And dtToDate.Value <> Nothing Then
                    Dim fromDate As String = dtFromDate.Value.ToString("yyyy-MM-dd")
                    Dim toDate As String = dtToDate.Value.ToString("yyyy-MM-dd")
                    dtlvar = " GEN_EGSP_IRN "
                    col1 = "IRN_Invoice_ID"
                    dtlvar2 = " ar_docMaster_IRN "
                    col2 = "VO_No"
                    'sqlQuery = "SELECT CAST(0 AS BIT) AS [Select],IRN_INVOICE_ID AS [INVNO],* FROM GEN_EGSP_IRN WHERE IS_INV_DRCR = 'DRCR' AND CONVERT(DATE, IRN_CreatedOn , 103) >= '" & fromDate & "' AND CONVERT(DATE, IRN_CreatedOn , 103) <= '" & toDate & "' AND IRN_Invoice_Unit = '" & gstrUNITID & "'"
                    sqlQuery = "SELECT CAST(0 AS BIT) AS [Select], IRN_INVOICE_ID AS [INV NO], IRN_Invoice_Date AS [IRN INVOICE DATE], " & _
                      "IRN_GSTIN_ID AS [IRN GSTIN ID], IRN_JSON AS [IRN JSON], IRN_Status AS [IRN STATUS], " & _
                      "IRN_Remarks AS [IRN Remarks], IRN_TRAN_NO AS [IRN TRAN NO], IRN_CreatedOn AS [IRN CreatedOn], " & _
                      "FILEUPLOADCODE AS [FILE UPLOAD CODE] " & _
                      "FROM GEN_EGSP_IRN " & _
                      "WHERE IS_INV_DRCR = 'DRCR' " & _
                      "AND CONVERT(DATE, IRN_CreatedOn, 103) >= '" & fromDate & "' " & _
                      "AND CONVERT(DATE, IRN_CreatedOn, 103) <= '" & toDate & "' " & _
                      "AND IRN_Invoice_Unit = '" & gstrUNITID & "'"


                    'sqlQuery1 = "SELECT CAST(0 AS BIT) AS [Select],VO_No AS [INVNO],* FROM ar_docMaster_IRN WHERE CONVERT(DATE, EWAY_UPD_DT  , 103) >= '" & fromDate & "' AND CONVERT(DATE, EWAY_UPD_DT  , 103) <= '" & toDate & "' AND UNIT_CODE = '" & gstrUNITID & "'"
                    sqlQuery1 = "SELECT CAST(0 AS BIT) AS [Select],VO_No AS [INV NO],EWAY_BILL_NO AS [EWAY BILL NO], " & _
                           "EWAY_TRANSPORT_MODE AS [EWAY TRANSPORT MODE],EWAY_TRANSPORTER_ID AS [EWAY TRANSPORTER ID], " & _
                           "EWAY_TRANSPORTER_DOC_NO AS [EWAY TRANSPORTER DOC_NO],EWAY_TRANSPORTER_DOC_DATE AS [EWAY TRANSPORTER DOC DATE]," & _
                           "EWAY_VEHICLE_NO AS [EWAY VEHICLE NO], EWAY_UPD_DT AS [EWAY UPD DT], EWAY_UPD_USERID AS [EWAY_UPD_USERID], " & _
                           "EWAY_EXP_DT AS [ EWAY EXP DT],EWAY_UPD_DT_EMPRO AS [EWAY UPD_DT EMPRO], EWAY_UPD_DT_PARTB AS [EWAY UPD_DT PARTB]," & _
                           "IRN_NO AS [IRN NO],IRN_DATE AS [IRN DATE],IRN_DATE_EMPRO AS [IRN DATE EMPRO]," & _
                           "IRN_DEACTIVATE AS [IRN DEACTIVATE],IRN_DEACTIVATE_DATE AS [IRN DEACTIVATE DATE]," & _
                           "ACK_NO AS [ACK NO],ACK_DT AS [ACK DT], CONEWBNO AS [CONEWB NO],CONEWBDT AS [CONEWB DT], " & _
                           "EINVOICE_STATUS AS[E-INVOICE STATUS],ERROR_MSG AS [ERROR_MSG]" & _
                           "FROM ar_docMaster_IRN " & _
                          " WHERE CONVERT(DATE, ENT_DT, 103) >= '" & fromDate & "' " & _
                          "AND CONVERT(DATE, ENT_DT, 103) <= '" & toDate & "' " & _
                          "AND UNIT_CODE = '" & gstrUNITID & "'"


                End If
            End If
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If (sqlQuery <> "") Then
                dt = SqlConnectionclass.GetDataTable(sqlQuery.ToString())
                GrdInv_DTL.DataSource = dt
                For Each column As DataGridViewColumn In GrdInv_DTL.Columns
                    If column.Name = "Select" Then
                        column.ReadOnly = False
                        column.Width = 30
                    Else
                        column.ReadOnly = True
                    End If
                Next

               
            End If
            If (sqlQuery1 <> "") Then
                dt2 = SqlConnectionclass.GetDataTable(sqlQuery1.ToString())
                Grd_IRN_DTL.DataSource = dt2
                For Each column As DataGridViewColumn In Grd_IRN_DTL.Columns
                    If column.Name = "Select" Then
                        column.ReadOnly = False
                        column.Width = 30
                    Else
                        column.ReadOnly = True
                    End If
                Next
               
            End If
            System.Windows.Forms.Cursor.Current = Cursors.Default

        Catch ex As Exception
            System.Windows.Forms.Cursor.Current = Cursors.Default
            MessageBox.Show(ex.Message)
        End Try


    End Sub

    Private Sub DeleteRows(ByVal dataGridView As DataGridView, ByVal dtlvar As String, ByVal dtlvar2 As String, ByVal col1 As String, ByVal col2 As String)


        Try
            Dim rowsToDelete1 As New List(Of DataGridViewRow)
            Dim rowsToDelete2 As New List(Of DataGridViewRow)

            If GrdInv_DTL.Columns.Contains("Select") AndAlso Grd_IRN_DTL.Columns.Contains("Select") Then

                For Each row As DataGridViewRow In GrdInv_DTL.Rows
                    If Convert.ToBoolean(row.Cells(0).Value) Then
                        rowsToDelete1.Add(row)
                    End If
                Next


                For Each row As DataGridViewRow In Grd_IRN_DTL.Rows
                    If Convert.ToBoolean(row.Cells(0).Value) Then
                        rowsToDelete2.Add(row)
                    End If
                Next


                For Each rowToDelete As DataGridViewRow In rowsToDelete1
                    Dim idToDelete As Integer = Convert.ToInt32(rowToDelete.Cells("Select").Value)
                    If idToDelete = 1 Then
                        Dim InvoiceToDelete As String = Convert.ToString(rowToDelete.Cells("INV NO").Value)
                        Dim Query As String = "INSERT INTO [GEN_EGSP_IRN_REPROCESS]  SELECT * FROM " & dtlvar & " WHERE " & col1 & " = '" & InvoiceToDelete & "' "
                        Query = Query & " DELETE FROM " & dtlvar & " WHERE " & col1 & " = '" & InvoiceToDelete & "'"
                        SqlConnectionclass.ExecuteNonQuery(Query.ToString())
                        GrdInv_DTL.Rows.Remove(rowToDelete)
                    End If
                Next

                For Each rowToDelete As DataGridViewRow In rowsToDelete2
                    Dim idToDelete As Integer = Convert.ToInt32(rowToDelete.Cells("Select").Value)
                    If idToDelete = 1 Then
                        Dim InvoiceToDelete As String = Convert.ToString(rowToDelete.Cells("INV NO").Value)
                        Dim deleteQuery As String = "DELETE FROM " & dtlvar2 & " WHERE " & col2 & " = '" & InvoiceToDelete & "'"
                        SqlConnectionclass.ExecuteNonQuery(deleteQuery.ToString())
                        Grd_IRN_DTL.Rows.Remove(rowToDelete)
                    End If
                Next
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub



    'Modified by sneha 

    Private Sub BtnReprocess_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnReprocess.Click
        Try
            Dim difference As Integer = (Date.Today - dtFromDate.Value.Date).Days

            If difference = 1 Or difference = 0 Then
                DeleteRows(GrdInv_DTL, dtlvar, dtlvar2, col1, col2)
                DeleteRows(Grd_IRN_DTL, dtlvar, dtlvar2, col1, col2)
                MessageBox.Show("Selected Invoices Have Been Reprocessed.")
            ElseIf difference >= 2 Then
                'DeleteRows(GrdInv_DTL, dtlvar, dtlvar2, col1, col2)
                'DeleteRows(Grd_IRN_DTL, dtlvar, dtlvar2, col1, col2)
                'Push_IRN_eWayBillDetail()
                MessageBox.Show("Invoices Can not be Reprocessed for more than 2 days backdate.")
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub



    Private Sub RefreshForm()
        Try
            GrdInv_DTL.DataSource = Nothing
            Grd_IRN_DTL.DataSource = Nothing
            Search_Field.Text = Nothing

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub DtFromDate_ValueChanged(ByVal sender As Object, ByVal e As EventArgs) Handles dtFromDate.ValueChanged
        If Not (rdobtnInv.Checked OrElse rdobtnSplymntryInv.Checked OrElse rdobtnCdt_Dbt.Checked) Then
            MessageBox.Show("Please select an invoice type before selecting a date.")
            'dtFromDate.Value = DateTime.Now 
        End If
    End Sub

    Private Sub DtToDate_ValueChanged(ByVal sender As Object, ByVal e As EventArgs) Handles dtToDate.ValueChanged
        If Not (rdobtnInv.Checked OrElse rdobtnSplymntryInv.Checked OrElse rdobtnCdt_Dbt.Checked) Then
            MessageBox.Show("Please select an invoice type before selecting a date.")
            'dtFromDate.Value = DateTime.Now
        End If

    End Sub

    Private Sub BtnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
        Try
            Me.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub RdobtnInv_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdobtnInv.CheckedChanged
        RefreshForm()

    End Sub

    Private Sub RdobtnSplymntryInv_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdobtnSplymntryInv.CheckedChanged
        RefreshForm()
    End Sub

    Private Sub RdobtnCdt_Dbt_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdobtnCdt_Dbt.CheckedChanged
        RefreshForm()

    End Sub



    Private Sub btnFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFilter.Click
        Try

            Dim ds1 As New DataSet
            Dim dt1 As New DataTable
            Dim strSQL3 As String = ""

            Dim ds2 As New DataSet
            Dim dt2 As New DataTable
            Dim strSQL4 As String = ""

            Dim searchValue As Integer
          

            Search_Field.Enabled = True


            If Search_Field.Text = " " Then
                Call MsgBox("Please select Field for Search.", vbOKOnly + vbInformation, "eMPro")
                Return
            End If



            '  If IsNothing(Search_Field.Text) OrElse Search_Field.Text.Trim() = "" Then
            If Search_Field.Text = " " Then
                MsgBox("Please enter a valid invoice Number for Search.", vbOKOnly + vbInformation, "eMPro")
                Return
            End If


            If rdobtnInv.Checked Then
                'strSQL3 = "SELECT CAST(0 AS BIT) AS [Select], IRN_INVOICE_ID AS [INVNO], * FROM GEN_EGSP_IRN WHERE IRN_INVOICE_ID = '" & Search_Field.Text & "'"

                strSQL3 = "SELECT CAST(0 AS BIT) AS [Select], IRN_INVOICE_ID AS [INV NO], IRN_Invoice_Date AS [IRN INVOICE DATE], " & _
                      "IRN_GSTIN_ID AS [IRN GSTIN ID], IRN_JSON AS [IRN JSON], IRN_Status AS [IRN STATUS], " & _
                      "IRN_Remarks AS [IRN Remarks], IRN_TRAN_NO AS [IRN TRAN NO], IRN_CreatedOn AS [IRN CreatedOn], " & _
                      "FILEUPLOADCODE AS [FILE UPLOAD CODE] " & _
                      "FROM GEN_EGSP_IRN " & _
                      "WHERE IRN_INVOICE_ID = '" & Search_Field.Text & "'"

                'strSQL4 = "SELECT CAST(0 AS BIT) AS [Select], Doc_No AS [INVNO], * FROM SALESCHALLAN_DTL_IRN WHERE Doc_No = '" & Search_Field.Text & "'"
                strSQL4 = "SELECT CAST(0 AS BIT) AS [Select], CAST(Doc_No AS VARCHAR) AS [INV NO],EWAY_BILL_NO AS [EWAY BILL NO], " & _
                          "EWAY_TRANSPORT_MODE AS [EWAY TRANSPORT MODE],EWAY_TRANSPORTER_ID AS [EWAY TRANSPORTER ID], " & _
                          "EWAY_TRANSPORTER_DOC_NO AS [EWAY TRANSPORTER DOC_NO],EWAY_TRANSPORTER_DOC_DATE AS [EWAY TRANSPORTER DOC DATE]," & _
                          "EWAY_VEHICLE_NO AS [EWAY VEHICLE NO], EWAY_UPD_DT AS [EWAY UPD DT], EWAY_UPD_USERID AS [EWAY_UPD_USERID], " & _
                          "EWAY_EXP_DT AS [ EWAY EXP DT],EWAY_UPD_DT_EMPRO AS [EWAY UPD_DT EMPRO], EWAY_UPD_DT_PARTB AS [EWAY UPD_DT PARTB]," & _
                          "IRN_NO AS [IRN NO],  IRN_DATE AS [IRN DATE],IRN_DATE_EMPRO AS [IRN DATE EMPRO]," & _
                          "IRN_DEACTIVATE AS [IRN DEACTIVATE],IRN_DEACTIVATE_DATE AS [IRN DEACTIVATE DATE]," & _
                          "ACK_NO AS [ACK NO],ACK_DT AS [ACK DT], CONEWBNO AS [CONEWB NO],CONEWBDT AS [CONEWB DT], " & _
                          "EINVOICE_STATUS AS[E-INVOICE STATUS],ERROR_MSG AS [ERROR_MSG]" & _
                          "FROM SALESCHALLAN_DTL_IRN " & _
                          "WHERE CAST(Doc_No AS VARCHAR) = '" & Search_Field.Text & "' "


            ElseIf rdobtnSplymntryInv.Checked Then

                ' strSQL3 = "SELECT CAST(0 AS BIT) AS [Select], IRN_INVOICE_ID AS [INVNO], * FROM GEN_EGSP_IRN WHERE IRN_INVOICE_ID = '" & Search_Field.Text & "'"
                strSQL3 = "SELECT CAST(0 AS BIT) AS [Select], IRN_INVOICE_ID AS [INV NO], IRN_Invoice_Date AS [IRN INVOICE DATE], " & _
                     "IRN_GSTIN_ID AS [IRN GSTIN ID], IRN_JSON AS [IRN JSON], IRN_Status AS [IRN STATUS], " & _
                     "IRN_Remarks AS [IRN Remarks], IRN_TRAN_NO AS [IRN TRAN NO], IRN_CreatedOn AS [IRN CreatedOn], " & _
                     "FILEUPLOADCODE AS [FILE UPLOAD CODE] " & _
                     "FROM GEN_EGSP_IRN " & _
                     "WHERE IRN_INVOICE_ID = '" & Search_Field.Text & "'"

                ' strSQL4 = "SELECT CAST(0 AS BIT) AS [Select], Doc_No AS [INVNO], * FROM Supplementary_IRN WHERE Doc_No = '" & Search_Field.Text & "'"
                strSQL4 = "SELECT CAST(0 AS BIT) AS [Select], VO_No  AS [INV NO],EWAY_BILL_NO AS [EWAY BILL NO], " & _
                        "EWAY_TRANSPORT_MODE AS [EWAY TRANSPORT MODE],EWAY_TRANSPORTER_ID AS [EWAY TRANSPORTER ID], " & _
                        "EWAY_TRANSPORTER_DOC_NO AS [EWAY TRANSPORTER DOC_NO],EWAY_TRANSPORTER_DOC_DATE AS [EWAY TRANSPORTER DOC DATE]," & _
                        "EWAY_VEHICLE_NO AS [EWAY VEHICLE NO], EWAY_UPD_DT AS [EWAY UPD DT], EWAY_UPD_USERID AS [EWAY_UPD_USERID], " & _
                        "EWAY_EXP_DT AS [ EWAY EXP DT],EWAY_UPD_DT_EMPRO AS [EWAY UPD_DT EMPRO], EWAY_UPD_DT_PARTB AS [EWAY UPD_DT PARTB]," & _
                        "IRN_NO AS [IRN NO], IRN_DATE AS [IRN DATE],IRN_DATE_EMPRO AS [IRN DATE EMPRO]," & _
                        "IRN_DEACTIVATE AS [IRN DEACTIVATE],IRN_DEACTIVATE_DATE AS [IRN DEACTIVATE DATE]," & _
                        "ACK_NO AS [ACK NO],ACK_DT AS [ACK DT], CONEWBNO AS [CONEWB NO],CONEWBDT AS [CONEWB DT], " & _
                        "EINVOICE_STATUS AS[E-INVOICE STATUS],ERROR_MSG AS [ERROR_MSG]" & _
                        "FROM Supplementary_IRN " & _
                        "WHERE VO_No = '" & Search_Field.Text & "' "


            ElseIf rdobtnCdt_Dbt.Checked Then

                'strSQL3 = "SELECT CAST(0 AS BIT) AS [Select], IRN_INVOICE_ID AS [INVNO], * FROM GEN_EGSP_IRN WHERE IRN_INVOICE_ID = '" & Search_Field.Text & "'"
                strSQL3 = "SELECT CAST(0 AS BIT) AS [Select], IRN_INVOICE_ID AS [INV NO], IRN_Invoice_Date AS [IRN INVOICE DATE], " & _
                     "IRN_GSTIN_ID AS [IRN GSTIN ID], IRN_JSON AS [IRN JSON], IRN_Status AS [IRN STATUS], " & _
                     "IRN_Remarks AS [IRN Remarks], IRN_TRAN_NO AS [IRN TRAN NO], IRN_CreatedOn AS [IRN CreatedOn], " & _
                     "FILEUPLOADCODE AS [FILE UPLOAD CODE] " & _
                     "FROM GEN_EGSP_IRN " & _
                     "WHERE IRN_INVOICE_ID = '" & Search_Field.Text & "'"


                ' strSQL4 = "SELECT CAST(0 AS BIT) AS [Select], VO_No AS [INVNO], * FROM ar_docMaster_IRN WHERE Doc_No = '" & Search_Field.Text & "'"
                strSQL4 = "SELECT CAST(0 AS BIT) AS [Select],VO_No AS [INV NO],EWAY_BILL_NO AS [EWAY BILL NO], " & _
                       "EWAY_TRANSPORT_MODE AS [EWAY TRANSPORT MODE],EWAY_TRANSPORTER_ID AS [EWAY TRANSPORTER ID], " & _
                       "EWAY_TRANSPORTER_DOC_NO AS [EWAY TRANSPORTER DOC_NO],EWAY_TRANSPORTER_DOC_DATE AS [EWAY TRANSPORTER DOC DATE]," & _
                       "EWAY_VEHICLE_NO AS [EWAY VEHICLE NO], EWAY_UPD_DT AS [EWAY UPD DT], EWAY_UPD_USERID AS [EWAY_UPD_USERID], " & _
                       "EWAY_EXP_DT AS [ EWAY EXP DT],EWAY_UPD_DT_EMPRO AS [EWAY UPD_DT EMPRO], EWAY_UPD_DT_PARTB AS [EWAY UPD_DT PARTB]," & _
                       "IRN_NO AS [IRN NO],  IRN_DATE AS [IRN DATE],IRN_DATE_EMPRO AS [IRN DATE EMPRO]," & _
                       "IRN_DEACTIVATE AS [IRN DEACTIVATE],IRN_DEACTIVATE_DATE AS [IRN DEACTIVATE DATE]," & _
                       "ACK_NO AS [ACK NO],ACK_DT AS [ACK DT], CONEWBNO AS [CONEWB NO],CONEWBDT AS [CONEWB DT], " & _
                       "EINVOICE_STATUS AS[E-INVOICE STATUS],ERROR_MSG AS [ERROR_MSG]" & _
                       "FROM ar_docMaster_IRN " & _
                       "WHERE VO_No = '" & Search_Field.Text & "' "

            Else
                MsgBox("Please select a valid option.", vbOKOnly + vbInformation, "eMPro")
                Return
            End If

            If Not String.IsNullOrEmpty(strSQL3) Then
                dt1 = SqlConnectionclass.GetDataTable(strSQL3)
                GrdInv_DTL.DataSource = dt1

                For Each column As DataGridViewColumn In GrdInv_DTL.Columns
                    If column.Name = "Select" Then
                        column.ReadOnly = False
                    Else
                        column.ReadOnly = True
                    End If
                Next

            End If

            If Not String.IsNullOrEmpty(strSQL4) Then
                dt2 = SqlConnectionclass.GetDataTable(strSQL4)
                Grd_IRN_DTL.DataSource = dt2

                For Each column As DataGridViewColumn In Grd_IRN_DTL.Columns
                    If column.Name = "Select" Then
                        column.ReadOnly = False
                    Else
                        column.ReadOnly = True
                    End If
                Next
            End If

        Catch ex As Exception
            ' Handle exceptions
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    'Changes by 07/03/2024 sneha 



    Private Function Push_IRN_eWayBillDetail() As String
        Try

            ' Using client = New HttpClient()

            Dim docType As String
            Dim Invtype As String
            Dim row1 As String = String.Empty
            Dim GUID As String = String.Empty
            Dim baseAddress_EWAY_IRN As String = String.Empty
            Dim token = ""
            Dim connetionString As String = Nothing
            Dim connection As SqlConnection
            Dim adapter As SqlDataAdapter
            Dim command As SqlCommand = New SqlCommand()
            Dim ds As DataSet = New DataSet()
            'connetionString = ConfigurationManager.AppSettings("ConnectionString")
            connection = glbSqlCon

            command.Connection = connection
            command.CommandType = CommandType.Text
            Dim FromDate As String = dtFromDate.Value.ToString("yyyy-MM-dd")
            'Dim FromDate As String = DateTime.Now.AddDays(Convert.ToDouble("-" & ConfigurationManager.AppSettings("BackDays"))).ToString("dd/MM/yyyy")
            Dim ToDate As String = DateTime.Now.ToString("dd/MM/yyyy")
            Dim Query As String = ""
            Query = "EXEC PROC_GET_IRN_DATA_TOJSON '" & gstrUNITID & "','" & FromDate & "','" & ToDate & "','SAVEDATA'"
            command.CommandText = Query
            command.CommandTimeout = 0
            'Mind.Log.WriteLog.WriteSteps(Query, "Data Fetch from SP PROC_GET_IRN_DATA_TOJSON ", LogPath)
            adapter = New SqlDataAdapter(command)
            adapter.Fill(ds)

            For Each row As DataRow In ds.Tables(0).Rows
                row1 = String.Empty
                row1 = row(2).ToString()
                GUID = String.Empty
                GUID = row(4).ToString()
                docType = row(0).ToString()
                Invtype = row(1).ToString()
                baseAddress_EWAY_IRN = row(3).ToString()
                'Mind.Log.WriteLog.WriteSteps(GUID, "GUID String", LogPath)

                If row1 <> "" Then
                    ServicePointManager.Expect100Continue = True
                    ServicePointManager.DefaultConnectionLimit = 9999
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls Or SecurityProtocolType.Ssl3

                    Dim jsonInputtoURI As String = JsonConvert.SerializeObject(row1)

                    If docType = "EWAY" Then
                        ''client.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "application/json")

                        Dim dt As DataTable = GetUnitAndclient_id()

                        If dt.Rows.Count > 0 Then
                            Dim url As String = dt.Rows(0)("baseAddress").ToString()
                            Dim myHttpWebRequest As HttpWebRequest = CType(HttpWebRequest.Create(url.ToString() & "/eInvoice/saveEInvoiceDetails"), HttpWebRequest)

                            myHttpWebRequest.Method = "POST"
                            Dim data As Byte() = Encoding.ASCII.GetBytes(jsonInputtoURI)
                            myHttpWebRequest.ContentType = "application/json"
                            myHttpWebRequest.ContentLength = data.Length
                            ''myHttpWebRequest.Headers.Add("Authorization", "Bearer " & access_token)
                            myHttpWebRequest.Headers.Add("x_api_key", "yBLbdDhnQlNpee6")
                            'Dim requestStream As Stream = myHttpWebRequest.GetRequestStream()
                            'requestStream.Write(data, 0, data.Length)
                            'requestStream.Close()
                            Dim myHttpWebResponse As HttpWebResponse = CType(myHttpWebRequest.GetResponse(), HttpWebResponse)
                            Dim responseStream As Stream = myHttpWebResponse.GetResponseStream()
                            Dim myStreamReader As StreamReader = New StreamReader(responseStream, Encoding.[Default])
                            Dim pageContent As String = myStreamReader.ReadToEnd()
                            myStreamReader.Close()
                            responseStream.Close()
                            myHttpWebResponse.Close()

                            'Dim theContent As StringContent = New StringContent(row1, System.Text.Encoding.UTF8, "application/json")
                            'Mind.Log.WriteLog.WriteSteps(DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss.fff"), "Service Call Start:", LogPath)

                            'Dim tokenResponse = client.PostAsync(baseAddress.ToString() & "/eInvoice/saveEInvoiceDetails", theContent).Result    'sneha
                            'Mind.Log.WriteLog.WriteSteps(DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss.fff"), "Service Call End:", LogPath)
                            'token = tokenResponse.Content.ReadAsStringAsync().Result
                            'Mind.Log.WriteLog.WriteSteps(token, "token", LogPath)
                            'token = "[" & token & "]"
                            Dim items As ToDoItem() = JsonConvert.DeserializeObject(Of ToDoItem())(pageContent)

                            If items(0).status_cd.ToString() = "0" Then


                                command.CommandType = CommandType.Text
                                Query = "update eMProDTLite_Utilities.dbo.GEN_EGSP_EWAYBILL set EWB_Status = '0', EWB_Remarks = '" & items(0).message.ToString() & " - Error Code : " & items(0).error_cd.ToString() & "' where EWB_GUID = '" & GUID & "' "
                                'Mind.Log.WriteLog.WriteSteps(Query, "Update eMProDTLite_Utilities.dbo.GEN_EGSP_EWAYBILL table with FAILURE", LogPath)
                                command.CommandText = Query
                                command.ExecuteNonQuery()
                                Dim Mailbody As String = "Please find the error for JSON mentioned in attached Log file.<br>" & items(0).message.ToString() & " - Error Code : " & items(0).error_cd.ToString()
                                ' Dim MailRet As String = eMail("EWAYBILL STATUS", Mailbody, Unit)
                            Else

                                If items(0).status_cd.ToString() = "1" Then
                                    command.CommandType = CommandType.Text
                                    Query = "update eMProDTLite_Utilities.dbo.GEN_EGSP_EWAYBILL set EWB_Status = '1', EWB_Remarks = 'SUCCESS' where EWB_GUID = '" & GUID & "' "
                                    'Mind.Log.WriteLog.WriteSteps(Query, "Update eMProDTLite_Utilities.dbo.GEN_EGSP_EWAYBILL table with SUCCESS", LogPath)
                                    command.CommandText = Query
                                    command.ExecuteNonQuery()
                                Else
                                    Dim items2 As ToDoItem2() = JsonConvert.DeserializeObject(Of ToDoItem2())(Convert.ToString(items(0).error_details))
                                    command.CommandType = CommandType.Text
                                    Query = "update eMProDTLite_Utilities.dbo.GEN_EGSP_EWAYBILL set EWB_Status = '2', EWB_Remarks = '" & items2(0).error_message.ToString() & " - Unit : " & items2(0).unit_code.ToString() & " - GSTIN : " & items2(0).gstin.ToString() & "' where EWB_GUID = '" & GUID & "' "
                                    'Mind.Log.WriteLog.WriteSteps(Query, "Update eMProDTLite_Utilities.dbo.GEN_EGSP_EWAYBILL table with FAILURE", LogPath)
                                    command.CommandText = Query
                                    command.ExecuteNonQuery()
                                    Dim Mailbody As String = "Please find the error for JSON mentioned in attached Log file.<br>" & items2(0).error_message.ToString()
                                    ' Dim MailRet As String = eMail("EWAYBILL STATUS for Unit :" & items2(0).unit_code.ToString(), Mailbody, Unit)
                                End If
                            End If
                        End If
                    End If


                    If docType = "IRN" AndAlso Invtype = "INV" Then
                        'client.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "application/json")

                        Dim dt As DataTable = GetUnitAndclient_id()
                        If dt.Rows.Count > 0 Then
                            Dim url As String = dt.Rows(0)("baseAddress").ToString()
                            Dim myHttpWebRequest As HttpWebRequest = CType(HttpWebRequest.Create(url.ToString() & "/eInvoice/saveEInvoiceDetails"), HttpWebRequest)

                            myHttpWebRequest.Method = "POST"
                            Dim data As Byte() = Encoding.ASCII.GetBytes(jsonInputtoURI)
                            myHttpWebRequest.ContentType = "application/json"
                            myHttpWebRequest.ContentLength = data.Length
                            ''myHttpWebRequest.Headers.Add("Authorization", "Bearer " & access_token)
                            myHttpWebRequest.Headers.Add("x_api_key", "yBLbdDhnQlNpee6")
                            'Dim requestStream As Stream = myHttpWebRequest.GetRequestStream()
                            'requestStream.Write(data, 0, data.Length)
                            'requestStream.Close()
                            Dim myHttpWebResponse As HttpWebResponse = CType(myHttpWebRequest.GetResponse(), HttpWebResponse)
                            Dim responseStream As Stream = myHttpWebResponse.GetResponseStream()
                            Dim myStreamReader As StreamReader = New StreamReader(responseStream, Encoding.[Default])
                            Dim pageContent As String = myStreamReader.ReadToEnd()
                            myStreamReader.Close()
                            responseStream.Close()
                            myHttpWebResponse.Close()

                            Dim items As IrnInvData() = JsonConvert.DeserializeObject(Of IrnInvData())(pageContent)

                            If items(0).fileuploadcode.ToString() = "90" Then
                                command.CommandType = CommandType.Text
                                Query = "update GEN_EGSP_IRN set IRN_Status = '1',FILEUPLOADCODE='" & items(0).fileuploadcode.ToString() & "',IRN_TRAN_NO='" & items(0).txn_id.ToString() & "', IRN_Remarks = 'SUCCESS' where IRN_GUID = '" & GUID & "' "
                                'Mind.Log.WriteLog.WriteSteps(Query, "Update GEN_EGSP_IRN table with SUCCESS", LogPath)
                                command.CommandText = Query
                                command.ExecuteNonQuery()
                            Else
                                command.CommandType = CommandType.Text
                                Query = "update GEN_EGSP_IRN set IRN_Status = '0',FILEUPLOADCODE='" & items(0).fileuploadcode.ToString() & "',IRN_TRAN_NO='" & items(0).txn_id.ToString() & "', IRN_Remarks = '" & items(0).fileuploadstatus.ToString() & " - Error Code : " & items(0).txn_id.ToString() & "' where IRN_GUID = '" & GUID & "' "
                                'Mind.Log.WriteLog.WriteSteps(Query, "Update GEN_EGSP_IRN table with FAILURE", LogPath)
                                command.CommandText = Query
                                command.ExecuteNonQuery()
                            End If
                        End If
                    End If

                    If docType = "IRN" AndAlso Invtype = "SUP" Then
                        'client.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "application/json")

                        Dim dt As DataTable = GetUnitAndclient_id()
                        If dt.Rows.Count > 0 Then
                            Dim url As String = dt.Rows(0)("baseAddress").ToString()
                            Dim myHttpWebRequest As HttpWebRequest = CType(HttpWebRequest.Create(url.ToString() & "/eInvoice/saveEInvoiceDetails"), HttpWebRequest)

                            myHttpWebRequest.Method = "POST"
                            Dim data As Byte() = Encoding.ASCII.GetBytes(jsonInputtoURI)
                            myHttpWebRequest.ContentType = "application/json"
                            myHttpWebRequest.ContentLength = data.Length
                            ''myHttpWebRequest.Headers.Add("Authorization", "Bearer " & access_token)
                            myHttpWebRequest.Headers.Add("x_api_key", "yBLbdDhnQlNpee6")
                            'Dim requestStream As Stream = myHttpWebRequest.GetRequestStream()
                            'requestStream.Write(data, 0, data.Length)
                            'requestStream.Close()
                            Dim myHttpWebResponse As HttpWebResponse = CType(myHttpWebRequest.GetResponse(), HttpWebResponse)
                            Dim responseStream As Stream = myHttpWebResponse.GetResponseStream()
                            Dim myStreamReader As StreamReader = New StreamReader(responseStream, Encoding.[Default])
                            Dim pageContent As String = myStreamReader.ReadToEnd()
                            myStreamReader.Close()
                            responseStream.Close()
                            myHttpWebResponse.Close()


                            Dim items As IrnInvData() = JsonConvert.DeserializeObject(Of IrnInvData())(pageContent)
                            If items(0).fileuploadcode.ToString() = "90" Then
                                command.CommandType = CommandType.Text
                                Query = "update GEN_EGSP_IRN set IRN_Status = '1',FILEUPLOADCODE='" & items(0).fileuploadcode.ToString() & "',IRN_TRAN_NO='" & items(0).txn_id.ToString() & "', IRN_Remarks = 'SUCCESS' where IRN_GUID = '" & GUID & "' "
                                'Mind.Log.WriteLog.WriteSteps(Query, "Update GEN_EGSP_IRN table with SUCCESS", LogPath)
                                command.CommandText = Query
                                command.ExecuteNonQuery()
                            Else
                                command.CommandType = CommandType.Text
                                Query = "update GEN_EGSP_IRN set IRN_Status = '0',FILEUPLOADCODE='" & items(0).fileuploadcode.ToString() & "',IRN_TRAN_NO='" & items(0).txn_id.ToString() & "', IRN_Remarks = '" & items(0).fileuploadstatus.ToString() & " - Error Code : " & items(0).txn_id.ToString() & "' where IRN_GUID = '" & GUID & "' "
                                'Mind.Log.WriteLog.WriteSteps(Query, "Update GEN_EGSP_IRN table with FAILURE", LogPath)
                                command.CommandText = Query
                                command.ExecuteNonQuery()
                            End If
                        End If
                    End If

                    If docType = "IRN" AndAlso Invtype = "DRCR" Then
                        ' client.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "application/json")

                        Dim dt As DataTable = GetUnitAndclient_id()
                        If dt.Rows.Count > 0 Then
                            Dim url As String = dt.Rows(0)("baseAddress").ToString()
                            Dim myHttpWebRequest As HttpWebRequest = CType(HttpWebRequest.Create(url.ToString() & "/eInvoice/saveEInvoiceDetails"), HttpWebRequest)

                            myHttpWebRequest.Method = "POST"
                            Dim data As Byte() = Encoding.ASCII.GetBytes(jsonInputtoURI)
                            myHttpWebRequest.ContentType = "application/json"
                            myHttpWebRequest.ContentLength = data.Length
                            ''myHttpWebRequest.Headers.Add("Authorization", "Bearer " & access_token)
                            myHttpWebRequest.Headers.Add("x_api_key", "yBLbdDhnQlNpee6")
                            'Dim requestStream As Stream = myHttpWebRequest.GetRequestStream()
                            'requestStream.Write(data, 0, data.Length)
                            'requestStream.Close()
                            Dim myHttpWebResponse As HttpWebResponse = CType(myHttpWebRequest.GetResponse(), HttpWebResponse)
                            Dim responseStream As Stream = myHttpWebResponse.GetResponseStream()
                            Dim myStreamReader As StreamReader = New StreamReader(responseStream, Encoding.[Default])
                            Dim pageContent As String = myStreamReader.ReadToEnd()
                            myStreamReader.Close()
                            responseStream.Close()
                            myHttpWebResponse.Close()


                            Dim items As IrnInvData() = JsonConvert.DeserializeObject(Of IrnInvData())(pageContent)

                            If items(0).fileuploadcode.ToString() = "90" Then
                                command.CommandType = CommandType.Text
                                Query = "update GEN_EGSP_IRN set IRN_Status = '1',FILEUPLOADCODE='" & items(0).fileuploadcode.ToString() & "',IRN_TRAN_NO='" & items(0).txn_id.ToString() & "', IRN_Remarks = 'SUCCESS' where IRN_GUID = '" & GUID & "' "
                                'Mind.Log.WriteLog.WriteSteps(Query, "Update GEN_EGSP_IRN table with SUCCESS", LogPath)
                                command.CommandText = Query
                                command.ExecuteNonQuery()
                            Else
                                command.CommandType = CommandType.Text
                                Query = "update GEN_EGSP_IRN set IRN_Status = '0',FILEUPLOADCODE='" & items(0).fileuploadcode.ToString() & "',IRN_TRAN_NO='" & items(0).txn_id.ToString() & "', IRN_Remarks = '" & items(0).fileuploadstatus.ToString() & " - Error Code : " & items(0).txn_id.ToString() & "' where IRN_GUID = '" & GUID & "' "
                                'Mind.Log.WriteLog.WriteSteps(Query, "Update GEN_EGSP_IRN table with FAILURE", LogPath)
                                command.CommandText = Query
                                command.ExecuteNonQuery()
                            End If
                        End If
                    End If
                End If

            Next

            connection.Close()
            Return token

        Catch exception As Exception
            'Mind.Log.WriteLog.WriteSteps(exception.ToString(), "WebAPI exception", LogPath)
            Return "exception"
        End Try

    End Function



    Private Function GetUnitAndclient_id() As DataTable

        Dim dt As DataTable
        dt = SqlConnectionclass.GetDataTable("select DISTINCT 'IRN' IRN_EWAY,unit_Code,RECEIVE_URL baseAddress,grant_type, client_id, client_secret,username, password from GEN_EGSP_IRN_CONFIG where IRN_REQUIRED=1 AND ACTIVE=1 And UNIT_CODE='" & gstrUNITID & "' ")

        Return dt
    End Function


    'Private Sub Get_Irn_From_Portal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Get_Irn_From_Portal.Click
    '    'System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
    '    'GeteWayBill_IRN_Detail()
    '    'System.Windows.Forms.Cursor.Current = Cursors.Default
    'End Sub




    Public Class DocDetail
        Private _conEwbNo As String
        Public Property conEwbNo() As String
            Get
                Return _conEwbNo
            End Get
            Set(ByVal value As String)
                _conEwbNo = value
            End Set
        End Property

        Private _conEwbDt As String
        Public Property conEwbDt() As String
            Get
                Return _conEwbDt
            End Get
            Set(ByVal value As String)
                _conEwbDt = value
            End Set
        End Property

        Private _isDeactivated As String
        Public Property isDeactivated() As String
            Get
                Return _isDeactivated
            End Get
            Set(ByVal value As String)
                _isDeactivated = value
            End Set
        End Property

        Private _doc_no As String
        Public Property doc_no() As String
            Get
                Return _doc_no
            End Get
            Set(ByVal value As String)
                _doc_no = value
            End Set
        End Property

        Private _irn As String
        Public Property irn() As String
            Get
                Return _irn
            End Get
            Set(ByVal value As String)
                _irn = value
            End Set
        End Property



        Private _ack_no As String
        Public Property ack_no() As String
            Get
                Return _ack_no
            End Get
            Set(ByVal value As String)
                _ack_no = value
            End Set
        End Property


        Private _ack_dt As String
        Public Property ack_dt() As String
            Get
                Return _ack_dt
            End Get
            Set(ByVal value As String)
                _ack_dt = value
            End Set
        End Property


        Private _qr_code As String
        Public Property qr_code() As String
            Get
                Return _qr_code
            End Get
            Set(ByVal value As String)
                _qr_code = value
            End Set
        End Property


        Private _einvoice_status As String
        Public Property einvoice_status() As String
            Get
                Return _einvoice_status
            End Get
            Set(ByVal value As String)
                _einvoice_status = value
            End Set
        End Property


        Private _ewb_no As String
        Public Property ewb_no() As String
            Get
                Return _ewb_no
            End Get
            Set(ByVal value As String)
                _ewb_no = value
            End Set
        End Property


        Private _ewb_Dt As String
        Public Property ewb_Dt() As String
            Get
                Return _ewb_Dt
            End Get
            Set(ByVal value As String)
                _ewb_Dt = value
            End Set
        End Property


        Private _ewb_status As String
        Public Property ewb_status() As String
            Get
                Return _ewb_status
            End Get
            Set(ByVal value As String)
                _ewb_status = value
            End Set
        End Property


        Private _ewb_valid_upto As String
        Public Property ewb_valid_upto() As String
            Get
                Return _ewb_valid_upto
            End Get
            Set(ByVal value As String)
                _ewb_valid_upto = value
            End Set
        End Property


        Private _vehicle_no As String
        Public Property vehicle_no() As String
            Get
                Return _vehicle_no
            End Get
            Set(ByVal value As String)
                _vehicle_no = value
            End Set
        End Property


        Private _transporter_id As String
        Public Property transporter_id() As String
            Get
                Return _transporter_id
            End Get
            Set(ByVal value As String)
                _transporter_id = value
            End Set
        End Property


        Private _transporter_name As String
        Public Property transporter_name() As String
            Get
                Return _transporter_name
            End Get
            Set(ByVal value As String)
                _transporter_name = value
            End Set
        End Property


        Private _distance As String
        Public Property distance() As String
            Get
                Return _distance
            End Get
            Set(ByVal value As String)
                _distance = value
            End Set
        End Property


        Private _transport_mode As String
        Public Property transport_mode() As String
            Get
                Return _transport_mode
            End Get
            Set(ByVal value As String)
                _transport_mode = value
            End Set
        End Property


        Private _transporter_doc_no As String
        Public Property transporter_doc_no() As String
            Get
                Return _transporter_doc_no
            End Get
            Set(ByVal value As String)
                _transporter_doc_no = value
            End Set
        End Property


        Private _consolidate_ewb_Date As String
        Public Property consolidate_ewb_Date() As String
            Get
                Return _consolidate_ewb_Date
            End Get
            Set(ByVal value As String)
                _consolidate_ewb_Date = value
            End Set
        End Property


        Private _transporter_doc_dt As String
        Public Property transporter_doc_dt() As String
            Get
                Return _transporter_doc_dt
            End Get
            Set(ByVal value As String)
                _transporter_doc_dt = value
            End Set
        End Property


        Private _error_message As String
        Public Property error_message() As String
            Get
                Return _error_message
            End Get
            Set(ByVal value As String)
                _error_message = value
            End Set
        End Property
    End Class



    Public Class UnitCode
        Private _unit_code As String
        Public Property unit_code() As String
            Get
                Return _unit_code
            End Get
            Set(ByVal value As String)
                _unit_code = value
            End Set
        End Property



        Private _doc_detail As List(Of DocDetail)
        Public Property Doc_Detail() As List(Of DocDetail)
            Get
                Return _doc_detail
            End Get
            Set(ByVal value As List(Of DocDetail))
                _doc_detail = value
            End Set
        End Property
    End Class


    Public Class Gstin
        Private _gstin As String
        Public Property gstin() As String
            Get
                Return _gstin
            End Get
            Set(ByVal value As String)
                _gstin = value
            End Set
        End Property

        Private _unit_code As List(Of UnitCode)
        Public Property unit_code() As List(Of UnitCode)
            Get
                Return _unit_code
            End Get
            Set(ByVal value As List(Of UnitCode))
                _unit_code = value
            End Set
        End Property
    End Class

    Public Class IRNType
        Private _type As String
        Public Property type() As String
            Get
                Return _type
            End Get
            Set(ByVal value As String)
                _type = value
            End Set
        End Property

        Private _gstin As List(Of Gstin)
        Public Property gstin() As List(Of Gstin)
            Get
                Return _gstin
            End Get
            Set(ByVal value As List(Of Gstin))
                _gstin = value
            End Set
        End Property
    End Class

    Public Class RootObject
        Private _status_cd As String
        Public Property status_cd() As String
            Get
                Return _status_cd
            End Get
            Set(ByVal value As String)
                _status_cd = value
            End Set
        End Property

        Private _type As List(Of IRNType)
        Public Property type() As List(Of IRNType)
            Get
                Return _type
            End Get
            Set(ByVal value As List(Of IRNType))
                _type = value
            End Set
        End Property
    End Class




    Private Function GeteWayBill_IRN_Detail() As String

        Try

            Dim FromDate As String = dtFromDate.Value.ToString("dd/MM/yyyy")
            Dim ToDate As String = DateTime.Now.ToString("dd/MM/yyyy")
            Dim Query As String = ""
            Query = "EXEC PROC_GET_IRN_DATA_TOJSON '" & gstrUNITID & "','" & FromDate & "','" & ToDate & "','GET_IRN_AND_EWAYBILL'"
           
            Dim DtTable As DataTable
            DtTable = SqlConnectionclass.GetDataTable(Query)
           
            Dim row1 As String = String.Empty
            row1 = DtTable.Rows(0)(0).ToString()
            Dim token = ""
            Dim isDRCR = ""


            For Each row As DataRow In DtTable.Rows
                row1 = String.Empty
                row1 = row(0).ToString()
                isDRCR = row(1).ToString()

                If row1 <> "" Then


                    ServicePointManager.Expect100Continue = True
                    ServicePointManager.DefaultConnectionLimit = 9999
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls Or SecurityProtocolType.Ssl3 Or CType(768, SecurityProtocolType) Or CType(3072, SecurityProtocolType)

                    'Const _Tls12 As SslProtocols = DirectCast(&HC00, SslProtocols)
                    'Const Tls12 As SecurityProtocolType = DirectCast(_Tls12, SecurityProtocolType)
                    'ServicePointManager.SecurityProtocol = Tls12

                    Dim intType As Integer
                    Dim jsonInputtoURI As String = JsonConvert.SerializeObject(row1)
                    Dim dt As DataTable = GetUnitAndclient_id()

                    Dim strGstnID As String = ""

                    If dt.Rows.Count > 0 Then
                        Dim url As String = dt.Rows(0)("baseAddress").ToString()
                        Dim myHttpWebRequest As HttpWebRequest = CType(HttpWebRequest.Create(url.ToString() & "/eInvoice/getEInvoiceDetails"), HttpWebRequest)

                        myHttpWebRequest.Method = "POST"
                        Dim data As Byte() = Encoding.ASCII.GetBytes(jsonInputtoURI)
                        myHttpWebRequest.ContentType = "application/json"
                        myHttpWebRequest.ContentLength = data.Length
                        myHttpWebRequest.Headers.Add("x_api_key", "yBLbdDhnQlNpee6")
                        ''myHttpWebRequest.KeepAlive = False
                        Dim requestStream As Stream = myHttpWebRequest.GetRequestStream()
                        requestStream.Write(data, 0, data.Length)
                        requestStream.Close()
                        Dim myHttpWebResponse As HttpWebResponse = CType(myHttpWebRequest.GetResponse(), HttpWebResponse)
                        Dim responseStream As Stream = myHttpWebResponse.GetResponseStream()
                        Dim myStreamReader As StreamReader = New StreamReader(responseStream, Encoding.[Default])
                        Dim pageContent As String = myStreamReader.ReadToEnd()
                        myStreamReader.Close()
                        responseStream.Close()
                        myHttpWebResponse.Close()

                        Dim items As RootObject() = JsonConvert.DeserializeObject(Of RootObject())(pageContent)
                        Dim strStatus As String = items(0).status_cd
                        Dim lstIRNType As List(Of IRNType) = items(0).type
                        Dim strIRNType As String = lstIRNType(0).type
                        Dim strIRNType2 As String = lstIRNType(1).type
                        Dim lstGstn As List(Of Gstin) = lstIRNType(0).gstin
                        Dim lstGstn2 As List(Of Gstin) = lstIRNType(1).gstin
                        strGstnID = lstGstn(0).gstin

                        Dim strGstnID2 As String = lstGstn2(0).gstin
                        Dim lstUnit As List(Of UnitCode) = lstGstn(0).unit_code
                        Dim lstUnit2 As List(Of UnitCode) = lstGstn2(0).unit_code
                        Dim strUnit_Code As String = lstUnit(0).unit_code
                        Dim strUnit_Code2 As String = lstUnit2(0).unit_code
                        Dim lstDoc_Details As List(Of DocDetail) = lstUnit(0).Doc_Detail
                        Dim lstDoc_Details2 As List(Of DocDetail) = lstUnit2(0).Doc_Detail
                        Query = ""


                        If isDRCR = "INV" Then

                            For intType = 0 To lstIRNType.Count - 1

                                If lstIRNType(intType).type = "1" Then

                                    If (lstDoc_Details IsNot Nothing) Then

                                        For i = 0 To lstDoc_Details.Count - 1
                                            Query = Query & " if exists (Select top 1 1 from SalesChallan_Dtl_IRN where DOC_No ='" & lstDoc_Details(i).doc_no & "' and UNIT_CODE='" & gstrUNITID & "') " & " begin " & " UPDATE SalesChallan_Dtl_IRN SET Location_Code='" & gstrUNITID & "',EWAY_BILL_NO='" & lstDoc_Details(i).ewb_no & "',EWAY_TRANSPORT_MODE='" & lstDoc_Details(i).transport_mode & "',EWAY_TRANSPORTER_ID='" & lstDoc_Details(i).transporter_id & "', " & " EWAY_TRANSPORTER_DOC_NO='" & lstDoc_Details(i).transporter_doc_no & "',EWAY_TRANSPORTER_DOC_DATE=Convert(Datetime,'" & lstDoc_Details(i).transporter_doc_dt & "',103),EWAY_VEHICLE_NO='" & lstDoc_Details(i).vehicle_no & "', " & " EWAY_UPD_DT=Convert(Datetime,'" & lstDoc_Details(i).ewb_Dt & "',103),EWAY_UPD_USERID='',EWAY_EXP_DT=Convert(Datetime,'" & lstDoc_Details(i).ewb_valid_upto & "',103),EWAY_UPD_DT_EMPRO='',EWAY_UPD_DT_PARTB= CASE WHEN EWAY_EXP_DT='1900-01-01 00:00:00.000' THEN '1900-01-01 00:00:00.000' ELSE EWAY_UPD_DT END ,IRN_NO='" & lstDoc_Details(i).irn & "', " & " IRN_DATE_EMPRO=getdate(),IRN_DEACTIVATE='" & lstDoc_Details(i).isDeactivated & "',ACK_NO='" & lstDoc_Details(i).ack_no & "',ACK_DT='" & lstDoc_Details(i).ack_dt & "',CONEWBNO='" & lstDoc_Details(i).conEwbNo & "',CONEWBDT='" & lstDoc_Details(i).conEwbDt & "',EINVOICE_STATUS='" & lstDoc_Details(i).einvoice_status & "',ERROR_MSG='" & lstDoc_Details(i).error_message & "' where DOC_No ='" & lstDoc_Details(i).doc_no & "' and UNIT_CODE='" & gstrUNITID & "' AND  ISNULL(IRN_DEACTIVATE,'0')='0' " & " end " & " else " & " begin " & " insert into SalesChallan_Dtl_IRN(Location_Code,Doc_No,UNIT_CODE,EWAY_BILL_NO,EWAY_TRANSPORT_MODE,EWAY_TRANSPORTER_ID,EWAY_TRANSPORTER_DOC_NO,EWAY_TRANSPORTER_DOC_DATE,EWAY_VEHICLE_NO, " & " EWAY_UPD_DT,EWAY_UPD_USERID,EWAY_EXP_DT,EWAY_UPD_DT_EMPRO,EWAY_UPD_DT_PARTB,IRN_NO,IRN_DATE,IRN_DATE_EMPRO,IRN_DEACTIVATE,ACK_NO,ACK_DT,CONEWBNO,CONEWBDT,EINVOICE_STATUS,ERROR_MSG) " & " Select '" & gstrUNITID & "','" & lstDoc_Details(i).doc_no & "','" & gstrUNITID & "','" & lstDoc_Details(i).ewb_no & "','" & lstDoc_Details(i).transport_mode & "','" & lstDoc_Details(i).transporter_id & "','" & lstDoc_Details(i).transporter_doc_no & "', " & " Convert(Datetime,'" & lstDoc_Details(i).transporter_doc_dt & "',103),'" & lstDoc_Details(i).vehicle_no & "',Convert(Datetime,'" & lstDoc_Details(i).ewb_Dt & "',103),'',Convert(Datetime,'" & lstDoc_Details(i).ewb_valid_upto & "',103),'','','" & lstDoc_Details(i).irn & "',getdate(),'','" & lstDoc_Details(i).isDeactivated & "','" & lstDoc_Details(i).ack_no & "','" & lstDoc_Details(i).ack_dt & "','" & lstDoc_Details(i).conEwbNo & "','" & lstDoc_Details(i).conEwbDt & "','" & lstDoc_Details(i).einvoice_status & "','" & lstDoc_Details(i).error_message & "' " & " end "
                                            Query = Query & " if exists (Select top 1 1 from SALESCHALLAN_DTL_IRN_BARCODE where DOC_No ='" & lstDoc_Details(i).doc_no & "' and UNIT_CODE='" & gstrUNITID & "') " & " begin " & " UPDATE SALESCHALLAN_DTL_IRN_BARCODE SET Location_Code='" & gstrUNITID & "',BARCODE_DATA='" & lstDoc_Details(i).qr_code & "' where DOC_No ='" & lstDoc_Details(i).doc_no & "' and UNIT_CODE='" & gstrUNITID & "'  " & " end " & " else " & " begin " & " insert into SALESCHALLAN_DTL_IRN_BARCODE(Location_Code,Doc_No,UNIT_CODE,BARCODE_DATA,barcodeimage) " & " Select '" & gstrUNITID & "','" & lstDoc_Details(i).doc_no & "','" & gstrUNITID & "','" & lstDoc_Details(i).qr_code & "',NULL " & " end "
                                        Next
                                    End If
                                End If


                                If lstIRNType(intType).type = "2" Then

                                    If (lstDoc_Details2 IsNot Nothing) Then

                                        For i = 0 To lstDoc_Details2.Count - 1

                                            If lstDoc_Details2(i).ewb_no = "" Then
                                                Query = Query & " update eMProDTLite_Utilities.dbo.GEN_EGSP_EWAYBILL set EWB_NO_STATUS = '" & strStatus & "', EWB_NO_Remarks = '" & lstDoc_Details2(i).error_message & "' where EWB_INVOICE_ID='" & lstDoc_Details2(i).doc_no & "' and (isnull(EWB_NO,'')='' OR (isnull(EWB_NO,'')<>'' AND EWB_NO_Exp='1900-01-01 00:00:00.000')) AND EWB_Invoice_Unit = '" & gstrUNITID & "' AND EWB_Invoice_Date BETWEEN Convert(Datetime,'" & FromDate & "',103) AND Convert(Datetime,'" & ToDate & "',103); "
                                            Else

                                                If Not String.IsNullOrEmpty(lstDoc_Details2(i).ewb_valid_upto) AndAlso Not String.IsNullOrEmpty(lstDoc_Details2(i).vehicle_no) Then
                                                    Query = Query & " update eMProDTLite_Utilities.dbo.GEN_EGSP_EWAYBILL set EWB_NO_STATUS = '" & strStatus & "', EWB_NO_Remarks = '" & lstDoc_Details2(i).error_message & "',EWB_NO = '" & lstDoc_Details2(i).ewb_no & "',EWB_NO_Exp = Convert(Datetime,'" & lstDoc_Details2(i).ewb_valid_upto & "',103), " & " EWB_NO_CreatedOn= Convert(Datetime,'" & lstDoc_Details2(i).ewb_Dt & "',103),vehicle_no = '" & lstDoc_Details2(i).vehicle_no & "',EWN_NO_PARTB_CreatedOn=getdate()  " & " where EWB_INVOICE_ID='" & lstDoc_Details2(i).doc_no & "' and (isnull(EWB_NO,'')='' OR (isnull(EWB_NO,'')<>'' AND EWB_NO_Exp='1900-01-01 00:00:00.000')) " & " AND EWB_Invoice_Unit = '" & gstrUNITID & "' AND EWB_Invoice_Date BETWEEN Convert(Datetime,'" & FromDate & "',103) AND Convert(Datetime,'" & ToDate & "',103); "
                                                Else
                                                    Query = Query & " update eMProDTLite_Utilities.dbo.GEN_EGSP_EWAYBILL set EWB_NO_STATUS = '" & strStatus & "', EWB_NO_Remarks = '" & lstDoc_Details2(i).error_message & "',EWB_NO = '" & lstDoc_Details2(i).ewb_no & "',EWB_NO_Exp = Convert(Datetime,'" & lstDoc_Details2(i).ewb_valid_upto & "',103), " & " EWB_NO_CreatedOn= Convert(Datetime,'" & lstDoc_Details2(i).ewb_Dt & "',103),vehicle_no = '" & lstDoc_Details2(i).vehicle_no & "'  " & " where EWB_INVOICE_ID='" & lstDoc_Details2(i).doc_no & "' and (isnull(EWB_NO,'')='' OR (isnull(EWB_NO,'')<>'' AND EWB_NO_Exp='1900-01-01 00:00:00.000')) " & " AND EWB_Invoice_Unit = '" & gstrUNITID & "' AND EWB_Invoice_Date BETWEEN Convert(Datetime,'" & FromDate & "',103) AND Convert(Datetime,'" & ToDate & "',103); "
                                                End If
                                            End If
                                        Next
                                    End If
                                End If
                            Next
                        End If


                        If isDRCR = "SUP" AndAlso (lstDoc_Details IsNot Nothing) Then

                            For intType = 0 To lstIRNType.Count - 1

                                If lstIRNType(intType).type = "1" Then

                                    For i = 0 To lstDoc_Details.Count - 1
                                        Query = Query & " if exists (Select top 1 1 from Supplementary_IRN where VO_NO ='" & lstDoc_Details(i).doc_no & "' and UNIT_CODE='" & gstrUNITID & "') " & " begin " & " UPDATE Supplementary_IRN SET DOC_NO=0,Location_Code='" & gstrUNITID & "',EWAY_BILL_NO='" & lstDoc_Details(i).ewb_no & "',EWAY_TRANSPORT_MODE='" & lstDoc_Details(i).transport_mode & "',EWAY_TRANSPORTER_ID='" & lstDoc_Details(i).transporter_id & "', " & " EWAY_TRANSPORTER_DOC_NO='" & lstDoc_Details(i).transporter_doc_no & "',EWAY_TRANSPORTER_DOC_DATE=Convert(Datetime,'" & lstDoc_Details(i).transporter_doc_dt & "',103),EWAY_VEHICLE_NO='" & lstDoc_Details(i).vehicle_no & "', " & " EWAY_UPD_DT=Convert(Datetime,'" & lstDoc_Details(i).ewb_Dt & "',103),EWAY_UPD_USERID='',EWAY_EXP_DT=Convert(Datetime,'" & lstDoc_Details(i).ewb_valid_upto & "',103),EWAY_UPD_DT_EMPRO='',EWAY_UPD_DT_PARTB= CASE WHEN EWAY_EXP_DT='1900-01-01 00:00:00.000' THEN '1900-01-01 00:00:00.000' ELSE EWAY_UPD_DT END ,IRN_NO='" & lstDoc_Details(i).irn & "', " & " IRN_DATE_EMPRO=getdate(),IRN_DEACTIVATE='" & lstDoc_Details(i).isDeactivated & "',ACK_NO='" & lstDoc_Details(i).ack_no & "',ACK_DT='" & lstDoc_Details(i).ack_dt & "',CONEWBNO='" & lstDoc_Details(i).conEwbNo & "',CONEWBDT='" & lstDoc_Details(i).conEwbDt & "',EINVOICE_STATUS='" & lstDoc_Details(i).einvoice_status & "',ERROR_MSG='" & lstDoc_Details(i).error_message & "' where VO_NO ='" & lstDoc_Details(i).doc_no & "' and UNIT_CODE='" & gstrUNITID & "' AND  ISNULL(IRN_DEACTIVATE,'0')='0' " & " end " & " else " & " begin " & " insert into Supplementary_IRN(DOC_NO,Location_Code,VO_NO,UNIT_CODE,EWAY_BILL_NO,EWAY_TRANSPORT_MODE,EWAY_TRANSPORTER_ID,EWAY_TRANSPORTER_DOC_NO,EWAY_TRANSPORTER_DOC_DATE,EWAY_VEHICLE_NO, " & " EWAY_UPD_DT,EWAY_UPD_USERID,EWAY_EXP_DT,EWAY_UPD_DT_EMPRO,EWAY_UPD_DT_PARTB,IRN_NO,IRN_DATE,IRN_DATE_EMPRO,IRN_DEACTIVATE,ACK_NO,ACK_DT,CONEWBNO,CONEWBDT,EINVOICE_STATUS,ERROR_MSG) " & " Select 0,'" & gstrUNITID & "','" & lstDoc_Details(i).doc_no & "','" & gstrUNITID & "','" & lstDoc_Details(i).ewb_no & "','" & lstDoc_Details(i).transport_mode & "','" & lstDoc_Details(i).transporter_id & "','" & lstDoc_Details(i).transporter_doc_no & "', " & " Convert(Datetime,'" & lstDoc_Details(i).transporter_doc_dt & "',103),'" & lstDoc_Details(i).vehicle_no & "',Convert(Datetime,'" & lstDoc_Details(i).ewb_Dt & "',103),'',Convert(Datetime,'" & lstDoc_Details(i).ewb_valid_upto & "',103),'','','" & lstDoc_Details(i).irn & "',getdate(),'','" & lstDoc_Details(i).isDeactivated & "','" & lstDoc_Details(i).ack_no & "','" & lstDoc_Details(i).ack_dt & "','" & lstDoc_Details(i).conEwbNo & "','" & lstDoc_Details(i).conEwbDt & "','" & lstDoc_Details(i).einvoice_status & "','" & lstDoc_Details(i).error_message & "' " & " end "
                                        Query = Query & " if exists (Select top 1 1 from Supplementary_IRN_BARCODE where VO_NO ='" & lstDoc_Details(i).doc_no & "' and UNIT_CODE='" & gstrUNITID & "') " & " begin " & " UPDATE Supplementary_IRN_BARCODE SET DOC_NO=0,Location_Code='" & gstrUNITID & "',BARCODE_DATA='" & lstDoc_Details(i).qr_code & "' where VO_NO ='" & lstDoc_Details(i).doc_no & "' and UNIT_CODE='" & gstrUNITID & "'  " & " end " & " else " & " begin " & " insert into Supplementary_IRN_BARCODE(DOC_NO,Location_Code,VO_NO,UNIT_CODE,BARCODE_DATA,barcodeimage) " & " Select 0,'" & gstrUNITID & "','" & lstDoc_Details(i).doc_no & "','" & gstrUNITID & "','" & lstDoc_Details(i).qr_code & "',NULL " & " end "
                                    Next
                                End If

                                If lstIRNType(intType).type = "2" Then
                                End If
                            Next
                        End If

                        If isDRCR = "DRCR" AndAlso (lstDoc_Details IsNot Nothing) Then

                            For intType = 0 To lstIRNType.Count - 1

                                If lstIRNType(intType).type = "1" Then

                                    For i = 0 To lstDoc_Details.Count - 1
                                        Query = Query & " if exists (Select top 1 1 from ar_docMaster_IRN where VO_NO ='" & lstDoc_Details(i).doc_no & "' and UNIT_CODE='" & gstrUNITID & "') " & " begin " & " UPDATE ar_docMaster_IRN SET EWAY_BILL_NO='" & lstDoc_Details(i).ewb_no & "',EWAY_TRANSPORT_MODE='" & lstDoc_Details(i).transport_mode & "',EWAY_TRANSPORTER_ID='" & lstDoc_Details(i).transporter_id & "', " & " EWAY_TRANSPORTER_DOC_NO='" & lstDoc_Details(i).transporter_doc_no & "',EWAY_TRANSPORTER_DOC_DATE=Convert(Datetime,'" & lstDoc_Details(i).transporter_doc_dt & "',103),EWAY_VEHICLE_NO='" & lstDoc_Details(i).vehicle_no & "', " & " EWAY_UPD_DT=Convert(Datetime,'" & lstDoc_Details(i).ewb_Dt & "',103),EWAY_UPD_USERID='',EWAY_EXP_DT=Convert(Datetime,'" & lstDoc_Details(i).ewb_valid_upto & "',103),EWAY_UPD_DT_EMPRO='',EWAY_UPD_DT_PARTB= CASE WHEN EWAY_EXP_DT='1900-01-01 00:00:00.000' THEN '1900-01-01 00:00:00.000' ELSE EWAY_UPD_DT END ,IRN_NO='" & lstDoc_Details(i).irn & "', " & " IRN_DATE_EMPRO=getdate(),IRN_DEACTIVATE='" & lstDoc_Details(i).isDeactivated & "',ACK_NO='" & lstDoc_Details(i).ack_no & "',ACK_DT='" & lstDoc_Details(i).ack_dt & "',CONEWBNO='" & lstDoc_Details(i).conEwbNo & "',CONEWBDT='" & lstDoc_Details(i).conEwbDt & "',EINVOICE_STATUS='" & lstDoc_Details(i).einvoice_status & "',ERROR_MSG='" & lstDoc_Details(i).error_message & "' where VO_NO ='" & lstDoc_Details(i).doc_no & "' and UNIT_CODE='" & gstrUNITID & "' AND  ISNULL(IRN_DEACTIVATE,'0')='0' " & " end " & " else " & " begin " & " insert into ar_docMaster_IRN(VO_NO,UNIT_CODE,EWAY_BILL_NO,EWAY_TRANSPORT_MODE,EWAY_TRANSPORTER_ID,EWAY_TRANSPORTER_DOC_NO,EWAY_TRANSPORTER_DOC_DATE,EWAY_VEHICLE_NO, " & " EWAY_UPD_DT,EWAY_UPD_USERID,EWAY_EXP_DT,EWAY_UPD_DT_EMPRO,EWAY_UPD_DT_PARTB,IRN_NO,IRN_DATE,IRN_DATE_EMPRO,IRN_DEACTIVATE,ACK_NO,ACK_DT,CONEWBNO,CONEWBDT,EINVOICE_STATUS,ERROR_MSG) " & " Select '" & lstDoc_Details(i).doc_no & "','" & gstrUNITID & "','" & lstDoc_Details(i).ewb_no & "','" & lstDoc_Details(i).transport_mode & "','" & lstDoc_Details(i).transporter_id & "','" & lstDoc_Details(i).transporter_doc_no & "', " & " Convert(Datetime,'" & lstDoc_Details(i).transporter_doc_dt & "',103),'" & lstDoc_Details(i).vehicle_no & "',Convert(Datetime,'" & lstDoc_Details(i).ewb_Dt & "',103),'',Convert(Datetime,'" & lstDoc_Details(i).ewb_valid_upto & "',103),'','','" & lstDoc_Details(i).irn & "',getdate(),'','" & lstDoc_Details(i).isDeactivated & "','" & lstDoc_Details(i).ack_no & "','" & lstDoc_Details(i).ack_dt & "','" & lstDoc_Details(i).conEwbNo & "','" & lstDoc_Details(i).conEwbDt & "','" & lstDoc_Details(i).einvoice_status & "','" & lstDoc_Details(i).error_message & "' " & " end "
                                        Query = Query & " if exists (Select top 1 1 from ar_docMaster_IRN_BARCODE where VO_NO ='" & lstDoc_Details(i).doc_no & "' and UNIT_CODE='" & gstrUNITID & "') " & " begin " & " UPDATE ar_docMaster_IRN_BARCODE SET BARCODE_DATA='" & lstDoc_Details(i).qr_code & "' where VO_NO ='" & lstDoc_Details(i).doc_no & "' and UNIT_CODE='" & gstrUNITID & "'  " & " end " & " else " & " begin " & " insert into ar_docMaster_IRN_BARCODE(VO_NO,UNIT_CODE,BARCODE_DATA,barcodeimage) " & " Select '" & lstDoc_Details(i).doc_no & "','" & gstrUNITID & "','" & lstDoc_Details(i).qr_code & "',NULL " & " end "
                                    Next
                                End If

                                If lstIRNType(intType).type = "2" Then
                                End If
                            Next
                        End If
                    End If ' changes ###########
                    If Query <> "" Then
                        'Mind.Log.WriteLog.WriteSteps("Update Start", "IRN Update Start for Unit:-" & gstrUNITID & " and GSTN:- " & strGstnID, LogPath)
                        SqlConnectionclass.ExecuteNonQuery(Query)
                        If isDRCR = "INV" Then
                            Query = "EXEC PROC_GET_IRN_DATA_TOJSON '" & gstrUNITID & "','" & FromDate & "','" & ToDate & "','UPDATE_IRN_EMPRO'"
                            Query = Query & " insert into IRN_GET_JSON_LOG(FROMTAB,UNIT_CODE,JSON_DATA,ENT_DT) select 'INV','" & gstrUNITID & "','" & token & "',getdate()"
                           SqlConnectionclass.ExecuteNonQuery(Query)
                        End If

                        If isDRCR = "SUP" Then
                            Query = "EXEC PROC_GET_IRN_DATA_TOJSON '" & gstrUNITID & "','" & FromDate & "','" & ToDate & "','UPDATE_DOC_NO'"
                            Query = Query & " insert into IRN_GET_JSON_LOG(FROMTAB,UNIT_CODE,JSON_DATA,ENT_DT) select 'SUP','" & gstrUNITID & "','" & token & "',getdate()"
                           SqlConnectionclass.ExecuteNonQuery(Query)
                        End If

                        If isDRCR = "DRCR" Then
                            Query = "EXEC PROC_GET_IRN_DATA_TOJSON '" & gstrUNITID & "','" & FromDate & "','" & ToDate & "','UPDATE_IRN_DRCR'"
                            Query = Query & " insert into IRN_GET_JSON_LOG(FROMTAB,UNIT_CODE,JSON_DATA,ENT_DT) select 'DRCR','" & gstrUNITID & "','" & token & "',getdate()"
                           SqlConnectionclass.ExecuteNonQuery(Query)
                        End If

                        'Mind.Log.WriteLog.WriteSteps("Update END", "IRN Update End for Unit:-" & gstrUNITID & " and GSTN:- " & strGstnID, LogPath)
                    End If
                End If
            Next

            System.Windows.Forms.Cursor.Current = Cursors.Default
            Return "SUCCESS"
        Catch exception As Exception
            System.Windows.Forms.Cursor.Current = Cursors.Default
            Return exception.Message
        End Try
    End Function

End Class
















