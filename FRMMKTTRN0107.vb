Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.IO
Imports System.Security.AccessControl
Imports Microsoft.Office.Interop


Public Class FRMMKTTRN0107

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Try
            DGVw_UploadedData.DataSource = Nothing
            txtFileLocation.Text = String.Empty
        Catch ex As Exception
            MessageBox.Show(ex.Message, "eMPRO")
        End Try
    End Sub

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        Try
            DGVw_UploadedData.DataSource = Nothing
            txtFileLocation.Text = String.Empty
            ' If OpenFileDlg.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            ' Dim con.ConnectionString = String.Format("Provider={0};Data Source={1};Extended Properties=""Excel 12.0 XML;HDR=Yes;""", "Microsoft.ACE.OLEDB.12.0", )
            Dim excelPathName As String = String.Empty
            With OpenFileDlg
                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    excelPathName = (CType(.FileName, String))
                    If (excelPathName.Length) <> 0 Then
                        Me.txtFileLocation.Text = excelPathName
                    Else
                        MsgBox("File Does Not Exist")
                    End If
                End If
            End With


        Catch ex As Exception
            MessageBox.Show(ex.Message, "eMPRO")
        End Try
    End Sub

    Private Function ReadData() As Boolean
        Dim Obj_EX As Microsoft.Office.Interop.Excel.Application
        Dim xlWorkSheet As Excel.Worksheet
        Dim icols As Integer
        Dim Conn As OleDbConnection
        Dim Conn2 As SqlConnection
        Dim sqlCmd As SqlCommand
        Dim sqlAdapter As SqlDataAdapter
        Dim destData As New DataTable
        Dim Doc_no As String = String.Empty
        Dim sourceData As New DataTable
        '--ADDED BY VIPIN--
        Dim destinationData_To_DB As New DataTable
        Dim Manupulate_To_DB As New DataTable
        Dim sourceConnString As String

        Try

            Obj_EX = New Microsoft.Office.Interop.Excel.Application
            Obj_EX.Workbooks.Open(txtFileLocation.Text.ToString.Trim)

            xlWorkSheet = Obj_EX.Worksheets(Obj_EX.Sheets(1).name)
            '---- FILL DATA INTO ARRAY OF OBJECTS----
            Dim strarr(xlWorkSheet.UsedRange.Rows.Count, xlWorkSheet.UsedRange.Columns.Count) As Object
            strarr = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(xlWorkSheet.UsedRange.Rows.Count, xlWorkSheet.UsedRange.Columns.Count)).Value

            Obj_EX.Workbooks.Close()
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
            Call releaseObject(xlWorkSheet)
            Application.DoEvents()
            '--FILL DATATABLE WITH ARRAY OBJECT----
            sourceData = MultidimensionalArrayToDataSetMultidime(strarr).Tables(0)

            '--- SET SOURCE DATATABLE TO DESTINATION DATA TABLE
            sourceData.Rows(0).Delete()
            sourceData.Rows(1).Delete()
            sourceData.AcceptChanges()

            Dim dv As DataView = New DataView(sourceData)
            dv.RowFilter = "F15='481' OR  F15='M48'"
            destinationData_To_DB = dv.ToTable()

            '---- MANUPULATE DATA INTO ROW PART NUMBER FORMAT--
            Manupulate_To_DB = destinationData_To_DB.Clone

            For i As Integer = 0 To destinationData_To_DB.Rows.Count - 1
                '-- FOR PART-1
                If Not String.IsNullOrEmpty(destinationData_To_DB.Rows(i)("F8").ToString) Then
                    Manupulate_To_DB.Rows.Add(destinationData_To_DB.Rows(i)("F1").ToString, destinationData_To_DB.Rows(i)("F3").ToString, destinationData_To_DB.Rows(i)("F4").ToString, destinationData_To_DB.Rows(i)("F5").ToString, destinationData_To_DB.Rows(i)("F6"), destinationData_To_DB.Rows(i)("F8").ToString, destinationData_To_DB.Rows(i)("F15").ToString, "LH", "F", destinationData_To_DB.Rows(i)("F13").ToString)
                End If

                '-- FOR PART-2
                If Not String.IsNullOrEmpty(destinationData_To_DB.Rows(i)("F10").ToString) Then
                    Manupulate_To_DB.Rows.Add(destinationData_To_DB.Rows(i)("F1").ToString, destinationData_To_DB.Rows(i)("F3").ToString, destinationData_To_DB.Rows(i)("F4").ToString, destinationData_To_DB.Rows(i)("F5").ToString, destinationData_To_DB.Rows(i)("F6"), destinationData_To_DB.Rows(i)("F10").ToString, destinationData_To_DB.Rows(i)("F15").ToString, "RH", "F", destinationData_To_DB.Rows(i)("F13").ToString)
                End If

                '-- FOR PART-3
                If Not String.IsNullOrEmpty(destinationData_To_DB.Rows(i)("F12").ToString) Then
                    Manupulate_To_DB.Rows.Add(destinationData_To_DB.Rows(i)("F1").ToString, destinationData_To_DB.Rows(i)("F3").ToString, destinationData_To_DB.Rows(i)("F4").ToString, destinationData_To_DB.Rows(i)("F5").ToString, destinationData_To_DB.Rows(i)("F6"), destinationData_To_DB.Rows(i)("F12").ToString, destinationData_To_DB.Rows(i)("F15").ToString, "LH", "R", destinationData_To_DB.Rows(i)("F13").ToString)
                End If

                '--FOR PART-4
                If Not String.IsNullOrEmpty(destinationData_To_DB.Rows(i)("F16").ToString) Then
                    Manupulate_To_DB.Rows.Add(destinationData_To_DB.Rows(i)("F1").ToString, destinationData_To_DB.Rows(i)("F3").ToString, destinationData_To_DB.Rows(i)("F4").ToString, destinationData_To_DB.Rows(i)("F5").ToString, destinationData_To_DB.Rows(i)("F6"), destinationData_To_DB.Rows(i)("F16").ToString, destinationData_To_DB.Rows(i)("F15").ToString, "RH", "R", destinationData_To_DB.Rows(i)("F13").ToString)
                End If
            Next

            If Manupulate_To_DB.Rows.Count = 0 Then
                MessageBox.Show(Now + "  : " + "Fail To manupulate Excel data.", "eMPRO")
                Obj_EX.Workbooks.Close()
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If
                '----END Z
                Return False
                Exit Function
            End If
            '-- GET THE DOCUMENT NUMBER--
            Doc_no = GetDocNumber()
            strDocNo = Doc_no
            If Doc_no.ToString.Contains("Fail To Generate Document Number") Then
                MessageBox.Show(Now + "  : " + "Fail To Generate Document Number", "eMPRO")
                Obj_EX.Workbooks.Close()
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If
                '----END Z
                Return False
                Exit Function
            End If
            'oleCmd.Dispose()
            'oleCmd = Nothing
            '-- start writing into database--


            sqlCmd = New SqlCommand
            sqlCmd.Connection = SqlConnectionclass.GetConnection()
            sqlCmd.CommandText = " DELETE FROM TRIGGER_FILE_MARUTI WHERE UNIT_CODE='MTM' AND DOC_NO='" & Doc_no & "'"
            sqlCmd.ExecuteNonQuery()

            ' Dim Trans As SqlTransaction = Conn2.BeginTransaction
            Dim qExecuted As Boolean = True
            Dim QuryIns As String
            sqlCmd.Transaction = sqlCmd.Connection.BeginTransaction

            For i As Integer = 0 To Manupulate_To_DB.Rows.Count - 1
                QuryIns = String.Empty
                sqlCmd.CommandText = String.Empty
                sqlCmd.CommandType = CommandType.Text
                sqlCmd.CommandTimeout = 0

                QuryIns = " If Not Exists (" & _
                        "   Select Top 1 1 From TRIGGER_FILE_MARUTI " & _
                        "   Where PSN ='" & Manupulate_To_DB.Rows(i)("F1").ToString & "' AND CHASSIS='" & Manupulate_To_DB.Rows(i)("F2").ToString & "' AND Unit_Code='MTM' AND PART_NO='" & Manupulate_To_DB.Rows(i)("F6").ToString & "'  )" & _
                        "   INSERT INTO TRIGGER_FILE_MARUTI (Doc_No,PSN,CHASSIS,VENDOR_CODE,MODEL_CODE,MODEL_DESC,ACHV_DATE,PART_NO,Unit_Code,Seq_date,Entered_date,UPLOAD_SOURCE,Uploaded_FileName,Status,QTY,IP_ADDRESS,ENT_DT,Ent_By,Upd_DT,Upd_By,CUST_CODE,SEQ_NO,Mset,Front_RearTYPE,Mode )" & _
                        "   VALUES ( '" & Doc_no & "','" & Manupulate_To_DB.Rows(i)("F1").ToString & "' ,'" & Manupulate_To_DB.Rows(i)("F2").ToString & "','" & Manupulate_To_DB.Rows(i)("F7").ToString & "','" & Manupulate_To_DB.Rows(i)("F3").ToString & "','" & Manupulate_To_DB.Rows(i)("F4").ToString & "' " & _
                        "   ,'" & Convert.ToDateTime(Manupulate_To_DB.Rows(i)("F5")).ToString("MM/dd/yyyy hh:mmm:sss tt") & "','" & Manupulate_To_DB.Rows(i)("F6").ToString & "','" & gstrUNITID & "','" & Convert.ToDateTime(Manupulate_To_DB.Rows(i)("F5")).ToString("yyyy-MM-dd hh:mmm:sss") & "',GETDATE(),'MARUTI','" & txtFileLocation.Text.ToString.Trim & "','1',1,'" & gstrIpaddressWinSck & "',GETDATE(),'" & mP_User & "',GETDATE(),'" & mP_User & "','C0000037','" & Manupulate_To_DB.Rows(i)("F1").ToString & "','" & Manupulate_To_DB.Rows(i)("F8").ToString() & "','" & Manupulate_To_DB.Rows(i)("F9").ToString() & "','" & "61" + Manupulate_To_DB.Rows(i)("F10").ToString() & "' )"
                sqlCmd.CommandText = QuryIns.ToString
                If sqlCmd.ExecuteNonQuery() = 0 Then
                    qExecuted = False
                    Exit For
                Else
                    qExecuted = True
                End If
            Next
            If qExecuted Then

                '-----%%%%%%%%%%%%----  GENERATE PICKLIST---------&&&&&&&&&&&&&&&&-----
                Dim ErrorMsg As String = String.Empty
                sqlCmd.Parameters.Clear()
                sqlCmd.CommandText = String.Empty
                sqlCmd.CommandType = CommandType.StoredProcedure
                sqlCmd.CommandText = "USP_GETSEQUENCENO_YSD"
                sqlCmd.Parameters.Add("@Unit_Code", SqlDbType.VarChar, 10).Value = "MTM"
                sqlCmd.Parameters.Add("@Document_No", SqlDbType.VarChar, 20).Value = Doc_no.ToString
                sqlCmd.Parameters.Add("@Cust_Code", SqlDbType.VarChar, 20).Value = "C0000037"
                sqlCmd.Parameters.Add("@Para", SqlDbType.VarChar, 100).Value = "GENERATE_PICKLISTNO"
                sqlCmd.Parameters.Add("@ErrorMsg", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                sqlCmd.CommandTimeout = 0
                sqlCmd.ExecuteNonQuery()
                ErrorMsg = sqlCmd.Parameters("@ErrorMsg").Value.ToString
                If String.IsNullOrEmpty(ErrorMsg) Then
                    sqlCmd.Transaction.Commit()
                    sqlCmd.Dispose()
                    MessageBox.Show("File Uploaded Successfully and Picklist Generated !!", "eMPRO")
                    Return True
                Else
                    sqlCmd.Transaction.Rollback()
                    MessageBox.Show("Error Uploaddata:-" & ErrorMsg & " - YSD File Uploading Fail. File Number-" + txtFileLocation.Text, "eMPRO")

                    If Not sqlCmd Is Nothing Then
                        sqlCmd.Dispose()
                        sqlCmd = Nothing
                    End If
                    Return False
                End If
            Else
                sqlCmd.Transaction.Rollback()
                MessageBox.Show(Now + "  : " + "Fail To Upload Document-" + txtFileLocation.Text.ToString())
                Return False
            End If
            '--- VALIDATE ENTRIES WETHER EXIST ALREADY IN ARCHIVE TABLE OR NOT
            Dim RowCount As Object
            Dim Qselect As String = String.Empty
            Qselect = "SELECT COUNT(*) FROM TRIGGER_FILE_MARUTI_ARCH(NOLOCK) A " & _
                      " WHERE A.UNIT_CODE='MTM' AND  EXISTS" & _
                      "(" & _
                      " SELECT * FROM TRIGGER_FILE_MARUTI(NOLOCK) B WHERE A.PART_NO=B.PART_NO AND A.UNIT_CODE=B.UNIT_CODE AND A.PSN=B.PSN AND A.CHASSIS=B.CHASSIS AND A.CUST_CODE=B.CUST_CODE " & _
                      " )"
            If Not sqlCmd Is Nothing Then
                sqlCmd.Dispose()
                sqlCmd = Nothing
            End If

            sqlCmd = New SqlCommand
            sqlCmd.Connection = SqlConnectionclass.GetConnection()

            sqlCmd.CommandType = CommandType.Text
            sqlCmd.CommandText = Qselect

            RowCount = sqlCmd.ExecuteScalar()

            If Val(RowCount) > 0 Then
                MessageBox.Show("Error Uploaddata - Record already exist for. File Number-" + txtFileLocation.Text.ToString, "eMPRO")
                '------- Z
                Obj_EX.Workbooks.Close()
                If Not Obj_EX Is Nothing Then
                    KillExcelProcess(Obj_EX)
                    Obj_EX = Nothing
                End If
                '----END Z
                Return False
                Exit Function
            Else
                Return False
                MessageBox.Show("Nothing To Upload.", "eMPRO")
            End If


            Obj_EX.Workbooks.Close()


            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If

        Catch ex As Exception
            MessageBox.Show(+"  : " + ex.Message, "eMPRO")
            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If
            MessageBox.Show("  : " + ex.Message, "eMPRO")
        Finally

            If Not Obj_EX Is Nothing Then
                KillExcelProcess(Obj_EX)
                Obj_EX = Nothing
            End If

            'Obj_FSO = Nothing
            If Not sqlCmd Is Nothing Then
                sqlCmd.Dispose()
                sqlCmd = Nothing
            End If


        End Try
    End Function

    Private Function MultidimensionalArrayToDataSetMultidime(ByVal input As Object(,)) As DataSet
        Dim dataSet = New DataSet()
        Dim dataTable = dataSet.Tables.Add()
        Dim iFila = input.GetLongLength(0)
        Dim iCol = input.GetLongLength(1)


        'For f As Integer = 1 To iFila - 1
        For f As Integer = 1 To iFila
            Dim row = dataTable.Rows.Add()

            For c As Integer = 0 To iCol - 1
                If f = 1 Then dataTable.Columns.Add("F" & (c + 1))
                row(c) = input(f, c + 1)
            Next
        Next

        Return dataSet
    End Function
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Dim strDocNo As String = String.Empty
    Private Function GetDocNumber() As String
        Dim sqlCmd As SqlCommand
        Dim Doc_No As Object = String.Empty
        Dim Qselect As String = String.Empty
        Qselect = " DECLARE @DOC_NO VARCHAR(10)" & _
                    " DECLARE @LASTID VARCHAR(10)" & _
                    " AGAIN: SELECT  @LASTID= MAX(CURRENT_NO)+1" & _
                    " FROM DOCUMENTTYPE_MST " & _
                    " WHERE UNIT_CODE='MTM'  AND  DOC_TYPE=66 " & _
                    " AND FIN_START_DATE BETWEEN " & _
                    "(" & _
                    " SELECT TOP 1 FIN_START_DATE  FROM FINANCIAL_YEAR_TB WHERE UNIT_CODE='MTM' AND  Convert(varchar(12),GETDATE(),106) BETWEEN FIN_START_DATE AND FIN_END_DATE " & _
                    " ) AND " & _
                    " (" & _
                    " SELECT TOP 1 FIN_END_DATE   FROM FINANCIAL_YEAR_TB WHERE UNIT_CODE='MTM' AND Convert(varchar(12),GETDATE(),106) BETWEEN FIN_START_DATE AND FIN_END_DATE " & _
                    " )" & _
                    " SET @DOC_NO=(SELECT CAST(RIGHT(YEAR(GETDATE()),2) AS VARCHAR(2))+SUBSTRING(CONVERT(NVARCHAR(6),GETDATE(), 112),5,2)+" & _
                    " + SUBSTRING(CONVERT(NVARCHAR(8),GETDATE(), 112),7,2)" & _
                    " +'0000'+ CAST(@LASTID AS INT) )" & _
                    " IF (SELECT TOP 1 1 FROM TRIGGER_FILE_MARUTI WHERE DOC_NO=@DOC_NO AND Unit_Code='MTM')=1" & _
                    " BEGIN " & _
                    " UPDATE DOCUMENTTYPE_MST SET CURRENT_NO=CURRENT_NO+1" & _
                    " WHERE UNIT_CODE='MTM'  AND  DOC_TYPE=66 " & _
                    " AND FIN_START_DATE BETWEEN " & _
                    " (" & _
                    " SELECT TOP 1 FIN_START_DATE  FROM FINANCIAL_YEAR_TB WHERE UNIT_CODE='MTM' AND Convert(varchar(12),GETDATE(),106) BETWEEN FIN_START_DATE AND FIN_END_DATE " & _
                    " ) AND " & _
                    " ( " & _
                    " SELECT TOP 1 FIN_END_DATE   FROM FINANCIAL_YEAR_TB WHERE UNIT_CODE='MTM' AND Convert(varchar(12),GETDATE(),106) BETWEEN FIN_START_DATE AND FIN_END_DATE" & _
                    " ) " & _
                    " GoTo AGAIN " & _
                    " End ELSE SELECT @DOC_NO AS DOC_NO "

        sqlCmd = New SqlCommand
        sqlCmd.Connection = SqlConnectionclass.GetConnection()
        sqlCmd.CommandType = CommandType.Text
        sqlCmd.CommandText = Qselect
        sqlCmd.CommandTimeout = 0
        Doc_No = sqlCmd.ExecuteScalar
        If IsDBNull(Doc_No) Then
            Return "Fail To Generate Document Number"
        Else
            Return Doc_No
        End If
    End Function
    Public Function KillExcelProcess(ByVal objEx As Microsoft.Office.Interop.Excel.Application) As Boolean
        Dim iHandle As IntPtr
        Dim proc As System.Diagnostics.Process
        Dim intPID As Integer
        Dim intResult As Integer
        Dim strVer As String
        Try
            strVer = objEx.Version
            iHandle = IntPtr.Zero
            If CInt(strVer) > 9 Then
                iHandle = New IntPtr(CType(objEx.Parent.Hwnd, Integer))
            Else
                iHandle = FindWindow(Nothing, objEx.Caption)
            End If
            objEx.Workbooks.Close()
            objEx.Quit()

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objEx)

            'EventLog1.WriteEntry("Excel Killed Successffully")
            objEx = Nothing

            intResult = GetWindowThreadProcessId(iHandle, intPID)
            proc = System.Diagnostics.Process.GetProcessById(intPID)
            proc.Kill()
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "eMPRO")
        End Try
    End Function

    Dim mintFormIndex As Integer
    Private Sub frmPLNTRN0027_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            mdifrmMain.CheckFormName = mintFormIndex
            System.Windows.Forms.Application.DoEvents()
            frmModules.NodeFontBold(Me.Tag) = True
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub frmPLNTRN0027_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        Try
            frmModules.NodeFontBold(Me.Tag) = False
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub frmPLNTRN0027_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            Dim KeyCode As Short = e.KeyCode
            Dim Shift As Short = e.KeyData \ &H10000
            If Shift <> 0 Then Exit Sub
            If KeyCode = System.Windows.Forms.Keys.F4 Then Call ctlHeader_Click(ctlHeader, New System.EventArgs()) : Exit Sub
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FMPLNTRN0027_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Call FitToClient(Me, GbxMain, ctlHeader, GrpBoxButtons)
            Me.MdiParent = mdifrmMain

            OpenFileDlg.Filter = "Excel Files|*.xls;*.xlsx"

        Catch ex As Exception
            MessageBox.Show(ex.Message, "eMPRO")
        End Try
    End Sub

    Private Sub ctlHeader_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlHeader.Click
        Try
            ' Call ShowHelp("underconstruction.htm")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "eMPRO")
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "eMPRO")
        End Try
    End Sub

    Private Sub btnUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        Try
            If txtFileLocation.Text.Trim.Length = 0 Then
                MessageBox.Show("Browse File Before Uploading!", "eMPRO")
                Exit Sub
            End If
            If MessageBox.Show("Are you Sure To Upload File?", "Confirmation", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
            If ReadData() Then
                Dim StrSql As String = String.Empty
                StrSql = "SELECT PICKLIST_NO,DOC_NO,PSN,CHASSIS,VENDOR_CODE,CUST_CODE,MODEL_CODE,MODEL_DESC,ACHV_DATE,PART_NO,QTY,MSET,FRONT_REARTYPE,CASE WHEN LEFT(MODE,2)='61' THEN 'ONLINE' WHEN LEFT(MODE,2)='63' THEN 'OFFLINE' END AS MODE, UPD_DT AS UPLOADED_STAMP FROM TRIGGER_FILE_MARUTI(NOLOCK) WHERE Unit_Code='" & gstrUNITID & "' and Doc_No='" & strDocNo & "'  ORDER BY PSN"
                Dim da As SqlDataAdapter = New SqlDataAdapter(StrSql, SqlConnectionclass.GetConnection())
                Dim dt As New DataTable
                da.SelectCommand.CommandTimeout = 0
                da.Fill(dt)
                DGVw_UploadedData.DataSource = dt

            End If
            txtFileLocation.Text = String.Empty
        Catch ex As Exception
            MessageBox.Show(ex.Message, "eMPRO")
        End Try
    End Sub
End Class