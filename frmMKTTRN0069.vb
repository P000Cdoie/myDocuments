Imports System
Imports System.Data.SqlClient
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Globalization

'CREATED BY     :   SHUBHRA VERMA
'CREATED ON     :   APR 2011
'FORM NAME      :   FORD RSA - CX RELEASE UPLOAD
'-------------------------------------------------------------------------------------------------------------------
'REVISED BY     :   VINOD SINGH
'REVISION DATE  :   03/06/2011
'REASON         :   CHANGES DONE FOR NEW FILE NAMES
'ISSUE ID       :   10101895
'REVISED BY     :   Prashant Dhingra
'REVISION DATE  :   16/06/2011
'REASON         :   Changes done For RSA (Report Added for Duplicate and Not Defined Item Codes)
'ISSUE ID       :   10104491
'REVISED BY     :   Prashant Dhingra
'REVISED DATE   :   27/06/2011
'REASON         :   1. New Columns added in ASN Generation 2. Schedule to be uploaded except for missing drawing No.
'ISSUE ID       :   10108772 
'-------------------------------------------------------------------------------------------------------------------
'REVISED BY     :   Shubhra Verma
'REVISED DATE   :   22/11/2011
'REASON         :   after uploading all files for the specified customer, first move all files to backup folder and then open the report.
'               :   While taking backup of uploaded files give an indication that backup process is in progress.
'ISSUE ID       :   10162313 
'-------------------------------------------------------------------------------------------------------------------
'REVISED BY     :   Shubhra Verma
'REVISED DATE   :   02/12/2011
'REASON         :   object reference error
'ISSUE ID       :   10165341
'-------------------------------------------------------------------------------------------------------------------
'REVISED BY     :   Shubhra Verma
'REVISED DATE   :   15/12/2011
'REASON         :   HResult exception error in report
'ISSUE ID       :   10170824
'-------------------------------------------------------------------------------------------------------------------
'REVISED BY     :   Shubhra Verma
'REVISED DATE   :   02/12/2011
'REASON         :   File Should Not Contain Both Delimiters (~ and ,)
'ISSUE ID       :   10172988
'-------------------------------------------------------------------------------------------------------------------
'REVISED BY     :   Shubhra Verma
'REVISED DATE   :   27/04/2012
'REASON         :   Report should dispaly only Active Schedule.
'ISSUE ID       :   10217089
'-------------------------------------------------------------------------------------------------------------------
'REVISED BY     :   Shubhra Verma
'REVISED DATE   :   08/05/2012
'REASON         :   Addition of Schedule upload date in Doc No Help
'ISSUE ID       :   10222060
'-------------------------------------------------------------------------------------------------------------------
'REVISED BY     :   Shubhra Verma
'REVISED DATE   :   29/05/2012
'REASON         :   DCI files for CX RSA upload 
'ISSUE ID       :   10230015
'-------------------------------------------------------------------------------------------------------------------
'REVISED BY     :   Shubhra Verma
'REVISED DATE   :   04/06/2012
'REASON         :   All files not moving to backup location
'ISSUE ID       :   10230995
'-------------------------------------------------------------------------------------------------------------------
'REVISED BY     :   Shubhra Verma
'REVISED DATE   :   05/10/2012
'REASON         :   Object reference error is coming if forecast files are not in correct format
'ISSUE ID       :   1043649
'-------------------------------------------------------------------------------------------------------------------
'REVISED BY     :   Shubhra Verma
'REVISED DATE   :   03/10/2013
'REASON         :   In CX File uploading, Remove process of moving files to backup folder
'ISSUE ID       :   10462517
'-------------------------------------------------------------------------------------------------------------------
'REVISED BY     :   Shubhra Verma
'REVISED DATE   :   25/11/2013
'REASON         :   In CX File uploading, enable process of moving files to backup folder
'ISSUE ID       :   10491045
'-------------------------------------------------------------------------------------------------------------------
'REVISED BY     :   Shubhra Verma
'REVISED DATE   :   20/03/2014
'REASON         :   MSSL RSA: Ford Schedule Uploading Twice
'ISSUE ID       :   10558163  
'-------------------------------------------------------------------------------------------------------------------
'REVISED BY     :   SHALINI SINGH
'REVISED DATE   :   08/07/2014
'REASON         :   10574175  NISSAN RSA file uploading.
'ISSUE ID       :   10574175  
'-------------------------------------------------------------------------------------------------------------------
'REVISED BY     :   SHALINI SINGH
'REVISED DATE   :   19/09/2014
'REASON         :   10672164   VW RSA file uploading.
'ISSUE ID       :   10672164   
'-------------------------------------------------------------------------------------------------------------------

Public Class frmMKTTRN0069
    Private mintFormIndex As Integer
    Dim mblnfilemove As Boolean
    Dim mBkpLocation As String
    Dim mLocalLocation As String
    Dim mShippingDays As Integer
    Dim mblnReportopen As Boolean
    Dim strFile As String = Nothing
    Dim mstrIncorrectFileFormatMsg As String
    Dim strCustFileString As String = ""
    Dim strHelpDoc() As String = Nothing

    Private Sub cmdCustHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCustHelp.Click
        Dim strHelp As String
        Dim strSql As String = ""
        Dim e1 As System.ComponentModel.CancelEventArgs

        Try

            strSql = "(select distinct c1.account_code, c2.cust_name ,c1.Unit_Code " & _
                " from custitem_mst c1, customer_mst c2 " & _
                " where c1.Unit_Code = c2.Unit_code and c1.unit_code ='" & gstrUNITID & "' and c1.account_code = c2.Customer_Code and ((isnull(c2.deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),c2.deactive_date,106)))) a"

            strHelp = ShowList(1, 1000, , "account_code", "cust_name", strSql, , "Customer Help", , , , )
            If strHelp = "-1" Then
                MessageBox.Show("No Customer Code Defined", ResolveResString(100), MessageBoxButtons.OK)
            Else
                txtCustomer.Text = strHelp
                Call txtCustomer_Validating(txtCustomer, e1)
            End If
            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try

    End Sub

    Private Sub txtCustomer_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustomer.KeyPress
        Try
            Dim KeyAscii As Short = Asc(e.KeyChar)
            Select Case KeyAscii
                Case System.Windows.Forms.Keys.Return
                    Call txtCustomer_Validating(txtCustomer, New System.ComponentModel.CancelEventArgs(False))
                Case 39, 34, 96
                    KeyAscii = 0
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub txtCustomer_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomer.KeyUp
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000
        Try
            If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
                Call cmdCustHelp_Click(cmdCustHelp, New System.EventArgs())
            End If
            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub txtCustomer_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCustomer.Validating
        Try
            If txtCustomer.Text.Trim.Length = 0 Then Exit Sub

            Dim strsql As String = ""
            Dim oCmd As SqlCommand
            Dim oRdr As SqlDataReader

            lblMessage.Text = ""
            strsql = "select distinct c1.account_code, c2.cust_name " & _
                " from custitem_mst c1, customer_mst c2 " & _
                " where c1.unit_code = c2.unit_code and c1.account_code = c2.Customer_Code" & _
                " and c1.unit_code = '" & gstrUNITID & "' and C1.account_code = '" & txtCustomer.Text & "'"

            oCmd = New SqlCommand(strsql, SqlConnectionclass.GetConnection)
            oRdr = oCmd.ExecuteReader
            If oRdr.HasRows Then
                Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
                oRdr.Read()
                lblCustName.Text = oRdr("cust_name").ToString
            Else
                MessageBox.Show("Invalid Customer Code.", ResolveResString(100), MessageBoxButtons.OK)
                Exit Sub
            End If
            Exit Sub
        Catch ex As Exception
            Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Function GetLocation() As Boolean
        'REVISED BY     :   Prashant Dhingra
        'REVISION DATE  :   13/06/2011
        'REASON         :   Changes done For RSA (Report Added for Duplicate and Not Defined Item Codes)
        'ISSUE ID       :   10104491
        'REVISED BY     :   Prashant Dhingra
        'REVISED DATE   :   27/06/2011
        'REASON         :   1. New Columns added in ASN Generation 2. Schedule to be uploaded except for missing drawing No.
        'ISSUE ID       :   10108772 
        Dim rs As IO.StreamReader = Nothing
        Dim readLine As String = Nothing
        Dim strText As String = Nothing
        Dim upldFiles As Scripting.File
        Dim SQLCMD As SqlCommand
        Dim sqlCon As SqlConnection
        Dim SQLRDR As SqlDataReader = Nothing
        Dim STRSQL As String = ""
        Dim DOC_NO As String = ""
        Dim objFSO As Scripting.FileSystemObject = Nothing
        Dim sqlTran As SqlTransaction
        Dim isTrans As Boolean = False
        Dim i As Integer
        'Added by vinod

        mstrIncorrectFileFormatMsg = ""
        Try
            STRSQL = "select TOP 1 BackUpLocation, LOCALLOCATION, SHIPPINGDAYS from FTP_Parameter_Mst (NOLOCK) " & _
                " WHERE UNIT_CODE = '" & gstrUNITID & "' AND CUSTOMER_CODE = '" & txtCustomer.Text & "' ORDER BY UPD_DT DESC "

            sqlCon = SqlConnectionclass.GetConnection
            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()

            SQLCMD = New SqlCommand(STRSQL, sqlCon)
            sqlTran = SQLCMD.Connection.BeginTransaction
            SQLCMD.CommandTimeout = 0
            SQLCMD.Transaction = sqlTran
            isTrans = True
            SQLRDR = SQLCMD.ExecuteReader()
            'shalini 10574175
            If SQLRDR.HasRows Then
                SQLRDR.Read()
                mBkpLocation = SQLRDR("BackUpLocation").ToString
                mLocalLocation = SQLRDR("LOCALLOCATION").ToString
                mShippingDays = SQLRDR("SHIPPINGDAYS").ToString
                'mBkpLocation = "D:\FORD_RELEASES_BACKUP"
                'mLocalLocation = "D:\FORD_RELEASES"
                'mLocalLocation = "c:\NISSANRSA_Schedule"
                'mBkpLocation = "c:\NISSANRSA_Schedule_Backup"
                'mLocalLocation = "c:\VWRSA_Schedule"
                'mBkpLocation = "c:\VWRSA_Schedule_Backup"
            Else
                MessageBox.Show("Locations Not Defined in Parameter Master.", ResolveResString(100), MessageBoxButtons.OK)
                SQLRDR.Close()
                If isTrans = True Then
                    sqlTran.Rollback()
                    isTrans = False
                End If
                Return False
            End If
            SQLRDR.Close()

            objFSO = New Scripting.FileSystemObject
            If Not objFSO.FolderExists(mLocalLocation) Then
                MessageBox.Show("Location" + mLocalLocation + " does not exists.", "help", MessageBoxButtons.OK)
                Return False
            End If

            'ADDED BY VINOD ON 03/06/2011, Issue Id : 10101895
            strCustFileString = GetCustomerFileString()
            ' END OF ADDITION
            i = 0
            mblnReportopen = False

            If objFSO.GetFolder(mLocalLocation).Files.Count > 0 Then
                STRSQL = "select current_no + 1 as current_no from documenttype_mst where UNIT_CODE = '" & gstrUNITID & "'" &
                        " AND doc_type = 609 AND CONVERT(DATETIME,CONVERT(VARCHAR(12),GETDATE(),106)) BETWEEN Fin_Start_date" &
                        " AND Fin_end_date"

                SQLCMD.CommandType = CommandType.Text
                SQLCMD.CommandText = STRSQL
                SQLRDR = SQLCMD.ExecuteReader

                If SQLRDR.HasRows Then
                    SQLRDR.Read()
                    DOC_NO = SQLRDR("current_no").ToString
                Else
                    MessageBox.Show("Document Type Master Not Defined.", ResolveResString(100), MessageBoxButtons.OK)
                    SQLRDR.Close()
                    If isTrans = True Then
                        sqlTran.Rollback()
                        isTrans = False
                    End If
                    Return False
                End If
                SQLRDR.Close()

                mP_Connection.Execute("Truncate Table FORDRSADRAWINGNOCORRECTION", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)

                HoldQueryTable = TableHoldQueryWithLineNo()
                HoldQueryTable2 = TableHoldQueryWithLineNo()

                GetItemCode()
                GetFileFormat()

                Dim fileNo As Integer = 0
                For Each upldFiles In objFSO.GetFolder(mLocalLocation).Files
                    fileNo = fileNo + 1
                    'Added by Vinod on 03/06/2011, Issue Id : 10101895
                    'If upldFiles.Name.ToUpper.Contains("FORD") = False Then
                    '    If upldFiles.Name.ToUpper.Contains(strCustFileString.ToUpper) = False Then
                    '        Continue For
                    '    End If
                    'End If
                    'End of Addition

                    strFile = upldFiles.Path
                    lbl_FailedToUpload.Tag = strFile
                    rs = File.OpenText(strFile)
                    readLine = rs.ReadLine
                    strText = ""
                    While Not readLine Is Nothing
                        strText = strText + readLine
                        readLine = rs.ReadLine
                    End While
                    Label1.Text = strText
                    rs.Close()
                    rs.Dispose()
                    If GetValues(DOC_NO, fileNo) = False Then
                        i = i + 1
                    End If
                Next

                SpExecution(HoldQueryTable)

                SpExecution(HoldQueryTable2)

                Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)

                If isTrans = True Then
                    STRSQL = "UPDATE DOCUMENTTYPE_MST SET CURRENT_NO = " & DOC_NO & " " &
                        " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_TYPE = '609' AND CONVERT(DATETIME,CONVERT(VARCHAR(12),GETDATE(),106)) BETWEEN Fin_Start_date" &
                        " AND Fin_end_date"
                    SQLCMD.CommandText = STRSQL
                    SQLCMD.ExecuteNonQuery()

                    SQLCMD.CommandType = CommandType.Text
                    SQLCMD.CommandText = " EXEC SP_UPDATEDAILYMKTFORRSASCHEDULE '" & gstrUNITID & "'," & DOC_NO & ""

                    SQLCMD.ExecuteNonQuery()
                    ' sqlTran.Rollback()
                    sqlTran.Commit()
                    isTrans = False
                    txtdocNo.Text = DOC_NO
                    lbl_Date.Text = DateTime.Now.ToString(format:="dd MMM yyyy")
                    MessageBox.Show("Schedule uploaded successfully.", ResolveResString(100), MessageBoxButtons.OK)
                    Return True
                End If

            Else

                MessageBox.Show("No File found in " + mLocalLocation + ".", ResolveResString(100), MessageBoxButtons.OK)
                objFSO = Nothing
                sqlTran.Rollback()
                Return False
            End If
            Exit Function

        Catch ex As Exception
            If Not strFile = Nothing Then
                Kill(strFile)
                rs.Dispose()
            End If
            objFSO = Nothing
            If isTrans = True Then
                sqlTran.Rollback()
                isTrans = False
            End If
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        Finally
            If Not SQLCMD Is Nothing Then
                SQLCMD.Dispose()
                SQLCMD = Nothing
            End If
            If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
        End Try
    End Function

    Dim strFormat As String = String.Empty

    Private Sub GetFileFormat()
        Dim sqlcmd As SqlCommand
        Dim sqlRDR As SqlDataReader
        Dim strSQL As String
        Dim sqlCon As SqlConnection

        Try
            sqlCon = SqlConnectionclass.GetConnection
            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()

            sqlcmd = New SqlCommand
            sqlcmd.Connection = sqlCon
            strSQL = "SELECT ISNULL(FILEFORMAT,'') AS FILEFORMAT FROM FTP_PARAMETER_MST WHERE UNIT_CODE = '" & gstrUNITID & "' AND CUSTOMER_CODE = '" & txtCustomer.Text & "'"
            sqlcmd.CommandText = strSQL
            sqlRDR = sqlcmd.ExecuteReader
            If sqlRDR.HasRows Then
                sqlRDR.Read()
                strFormat = sqlRDR("FILEFORMAT").ToString
            End If
            sqlRDR.Close()
            sqlcmd.Dispose()
            sqlcmd = Nothing
        Catch ex As Exception
            If Not sqlcmd Is Nothing Then
                sqlcmd.Dispose()
                sqlcmd = Nothing
            End If
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        Finally
            If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
        End Try
    End Sub

    Dim ItemCodeTable As DataTable
    Dim HoldQueryTable As DataTable
    Dim HoldQueryTable2 As DataTable
    Private Sub GetItemCode()
        Try
            ItemCodeTable = New DataTable()
            ItemCodeTable = SqlConnectionclass.GetDataTable("select item_code,cust_drgno from custitem_mst where UNIT_CODE = '" & gstrUNITID & "' AND account_code = '" & txtCustomer.Text & "' and active = 1")
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        Finally
        End Try
    End Sub

    Private Sub SpExecution(ByVal table As DataTable)
        Dim sqlCon As SqlConnection
        Dim sqlcmd As SqlCommand

        Try
            sqlCon = SqlConnectionclass.GetConnection
            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()

            sqlcmd = New SqlCommand
            sqlcmd.Connection = sqlCon
            sqlcmd.CommandText = "dbo.FordRsaUpload_LineNo"
            sqlcmd.CommandType = CommandType.StoredProcedure
            sqlcmd.CommandTimeout = 0
            Dim param As SqlParameter = sqlcmd.Parameters.AddWithValue("@Data", table)
            param.SqlDbType = SqlDbType.Structured
            param.TypeName = "dbo.FordRsaTable_LineNo"

            sqlcmd.ExecuteNonQuery()
            sqlcmd.Dispose()
            sqlcmd = Nothing

        Catch ex As Exception
            If Not sqlcmd Is Nothing Then
                sqlcmd.Dispose()
                sqlcmd = Nothing
            End If
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        Finally
            If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
        End Try
    End Sub

    Private Function GetValues(ByVal Doc_No As String, Optional ByVal FileNo As Integer = 0) As Boolean

        'Dim sqlcmd As SqlCommand
        'Dim sqlRDR As SqlDataReader
        'Dim strSQL As String
        'Dim strFormat As String
        'Dim sqlCon As SqlConnection

        Try
            'sqlCon = SqlConnectionclass.GetConnection
            'If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()

            'sqlcmd = New SqlCommand
            'sqlcmd.Connection = sqlCon
            'strSQL = "SELECT ISNULL(FILEFORMAT,'') AS FILEFORMAT FROM FTP_PARAMETER_MST (Nolock) WHERE UNIT_CODE = '" & gstrUNITID & "'" & _
            '    " AND CUSTOMER_CODE = '" & txtCustomer.Text & "'"
            'sqlcmd.CommandText = strSQL
            'sqlRDR = sqlcmd.ExecuteReader
            'If sqlRDR.HasRows Then
            '    sqlRDR.Read()
            '    strFormat = sqlRDR("FILEFORMAT").ToString
            'End If
            'sqlRDR.Close()
            'sqlcmd.Dispose()
            'sqlcmd = Nothing

            If strFormat.ToUpper = "FORD" Then
                If InStr(strFile, strCustFileString + "R_", CompareMethod.Text) >= 1 Then
                    Return GetValues_FORD_forecast(Doc_No, FileNo)
                End If
                If InStr(strFile, strCustFileString + "D_", CompareMethod.Text) >= 1 Then
                    Return GetValues_FORD_firm(Doc_No)
                End If
            End If

            If strFormat.ToUpper = "BMW" Then
                Return GetValues_BMW(Doc_No)
            End If

            If strFormat.ToUpper = "NISSAN" Then
                'Shalini
                If Mid(Label1.Text, 5, 6) = "DELINS" Then
                    If gstrUNITID.ToUpper = "MGS" Then
                        Return GetValues_NISSAN_RSA_Forecast(Doc_No)
                    Else
                        Return GetValues_NISSAN_Forecast(Doc_No)
                    End If
                Else
                    If gstrUNITID.ToUpper = "MGS" Then
                        Return GetValues_NISSAN_RSA_Firm(Doc_No)
                    Else
                        Return GetValues_NISSAN_Firm(Doc_No)
                    End If
                End If
            End If

            If strFormat.ToUpper = "BENZ" Then
                Return GetValues_BENZ(Doc_No)
            End If

            If strFormat.ToUpper = "VW" Then
                'Shalini
                If Mid(Label1.Text, 5, 6) = "DELINS" Then
                    If gstrUNITID.ToUpper = "MGS" Then
                        Return GetValues_NISSAN_RSA_Forecast(Doc_No)
                    Else
                        Return GetValues_NISSAN_Forecast(Doc_No)
                    End If
                Else
                    If gstrUNITID.ToUpper = "MGS" Then
                        Return GetValues_NISSAN_RSA_Firm(Doc_No)
                    Else
                        Return GetValues_NISSAN_Firm(Doc_No)
                    End If
                End If
            End If


        Catch ex As Exception
            'If Not sqlcmd Is Nothing Then
            '    sqlcmd.Dispose()
            '    sqlcmd = Nothing
            'End If
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        Finally
            'If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
        End Try

    End Function

    Private Function GetValues_FORD_forecast(ByVal DOC_NO As String) As Boolean
        'REVISED BY     :   Prashant Dhingra
        'REVISION DATE  :   16/06/2011
        'REASON         :   Changes done For RSA (Report Added for Duplicate and Not Defined Item Codes)
        'ISSUE ID       :   10104491
        'REVISED BY     :   Prashant Dhingra
        'REVISED DATE   :   27/06/2011
        'REASON         :   1. New Columns added in ASN Generation 2. Schedule to be uploaded except for missing drawing No.
        'ISSUE ID       :   10108772 
        Dim intStart, intEnd As Integer
        Dim lngLength As Long
        Dim i As Integer = 0
        Dim itemcount As Integer = 0
        Dim sqlcmd As SqlCommand
        Dim sqlCon As SqlConnection
        Dim sqlRdr As SqlDataReader
        Dim strSql As String = ""
        Dim strHdr(48) As String
        'Dim strDtl(300) As String
        Dim strDtl(800) As String
        Dim lngPointer As Long
        Dim chr As Char = String.Empty
        Dim strItemCode As String
        Dim strCreateDate As Date
        Dim strSQLD As String = String.Empty
        Dim blnUpldFileStatus As Boolean
        Dim blnDrgStatus As Boolean
        Dim chrDupORNotDefined As Char
        Dim LineNo As Integer

        Dim k As Integer
        k = 0
        Dim j As Integer
        j = 0
        GetValues_FORD_forecast = True
        Try
            intStart = 0
            If InStr(intStart + 1, Label1.Text, "~") > 0 And InStr(intStart + 1, Label1.Text, ",") > 0 Then
                mstrIncorrectFileFormatMsg = mstrIncorrectFileFormatMsg + "File Name: " + strFile + "{File Contains Both Delimiters (~ and ,)}" + vbCrLf
                Return False
            End If
            chr = ""
            If InStr(intStart + 1, Label1.Text, "~") > 0 Then
                chr = "~"
            End If
            If InStr(intStart + 1, Label1.Text, ",") > 0 Then
                chr = ","
            End If

            If chr = String.Empty Or chr = Nothing Then
                mstrIncorrectFileFormatMsg = mstrIncorrectFileFormatMsg + "File Name: " + strFile + "{File Contains No Delimiter (~ OR ,)}" + vbCrLf
                Return False
            End If


            intEnd = InStr(1, Label1.Text, "DTL") - 1

            lngLength = intEnd
            lngPointer = 0

            While lngPointer <= lngLength
                intEnd = InStr(intStart + 1, Label1.Text, chr)
                If intEnd < intStart Then Exit While
                If LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString <> "HDR" Then
                    strHdr(i) = LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString
                    i = i + 1
                End If
                intStart = intEnd
                lngPointer = intEnd
            End While

            i = i - 1
            strHdr(i) = Mid(strHdr(i), 1, Len(strHdr(i)) - 3)

            lngLength = Label1.Text.Length
            i = 0
            While lngPointer <= lngLength
                intEnd = InStr(intStart + 1, Label1.Text, chr)
                If intEnd = 0 And intStart < lngLength Then intEnd = lngLength + 1
                strDtl(i) = LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString
                i = i + 1
                intStart = intEnd
                lngPointer = intEnd
            End While
            ReDim Preserve strDtl(UBound(strDtl) + 1)
            lngPointer = 0

            sqlCon = SqlConnectionclass.GetConnection
            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()
            sqlcmd = New SqlCommand
            sqlcmd.Connection = sqlCon

            blnUpldFileStatus = True
            blnDrgStatus = True

            'strSql = "select item_code from custitem_mst (Nolock) where UNIT_CODE = '" & gstrUNITID & "' AND account_code = '" & txtCustomer.Text & "'" & _
            '    " and cust_drgno = '" & strHdr(15) & "' and active = 1"
            'sqlcmd.CommandText = strSql
            'sqlRdr = sqlcmd.ExecuteReader

            'itemcount = 0
            'If sqlRdr.HasRows Then
            '    While sqlRdr.Read()
            '        itemcount = itemcount + 1
            '        strItemCode = sqlRdr("item_code").ToString
            '    End While
            '    If itemcount > 1 Then
            '        chrDupORNotDefined = "D"
            '        GetValues_FORD_forecast = False
            '    End If
            'Else
            '    chrDupORNotDefined = "N"
            '    GetValues_FORD_forecast = False
            'End If
            'sqlRdr.Close()

            itemcount = 0
            Dim datarow As DataRow() = ItemCodeTable.Select("cust_drgno ='" & strHdr(15) & "'")
            If datarow Is Nothing Or datarow.Count = 0 Then
                chrDupORNotDefined = "N"
                GetValues_FORD_forecast = False
            Else
                For Each row As DataRow In datarow
                    itemcount = itemcount + 1
                    strItemCode = row("item_code").ToString
                Next
                If (itemcount > 1) Then
                    chrDupORNotDefined = "D"
                    GetValues_FORD_forecast = False
                End If
            End If

            If GetValues_FORD_forecast = False Then
                GetValues_FORD_forecast = True
                mblnReportopen = True
                sqlcmd.Dispose()
                blnUpldFileStatus = False
                blnDrgStatus = False
            End If

            strCreateDate = Mid(strHdr(14), 5, 2) + " / " + Mid(strHdr(14), 3, 2) + " / " + Mid(strHdr(14), 1, 2)

            strSql = "set dateformat 'dmy' insert into RSASchedule_Hdr (Account_Code,Doc_No,Plant_Code,Unit_Code," &
                " CreateDate,Item_code,Cust_Drgno,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,DRG_STATUS,UPLDFILENAME," &
                " SCHEDULETYPE,FileUpld_Status,DupORNotDefined) " &
                " values('" & txtCustomer.Text & "'," & DOC_NO & ",'" & strHdr(5) & "','" & gstrUNITID & "'," &
                " '" & strCreateDate & "','" & strItemCode & "','" & strHdr(15) & "',getdate()," &
                " '" & mP_User & "',getdate(), '" & mP_User & "','" & blnDrgStatus & "','" & strFile & "','F','" & blnUpldFileStatus & "'," &
                " '" & chrDupORNotDefined & "')"

            HoldQueryTable.Rows.Add(1, "Insert", strHdr(15), strSql)

            'sqlcmd.CommandText = strSql
            'sqlcmd.ExecuteNonQuery()

            'If blnUpldFileStatus = False Then
            '    Return False
            '    If Not sqlcmd Is Nothing Then
            '        sqlcmd.Dispose()
            '        sqlcmd = Nothing
            '    End If
            'End If

            lngLength = i - 1
            i = 0
            While i <= lngLength
                LineNo = Convert.ToInt32(strDtl(i).ToString())
                If strDtl(i + 6) = "Nothing" Or strDtl(i + 6) = "" Then Exit While
                strCreateDate = Mid(strDtl(i + 6), 5, 2) + " / " + Mid(strDtl(i + 6), 3, 2) + " / " + Mid(strDtl(i + 6), 1, 2)
                strSql = "set dateformat 'dmy' UPDATE RSASchedule_DTL SET DRG_STATUS = 0, Upd_dt=getdate() WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & DOC_NO & "'" &
                         " AND CUST_DRGNO = '" & strHdr(15) & "' AND DRG_STATUS = 1 and scheduletype = 'F' AND TRANS_DATE = '" & strCreateDate & "'"

                HoldQueryTable2.Rows.Add(LineNo, "Update", strHdr(15), strSql)
                'sqlcmd.CommandText = strSql
                'sqlcmd.ExecuteNonQuery()

                strSql = "set dateformat 'dmy' insert into RSASchedule_Dtl(Account_Code,Unit_Code,Line_No,Trans_Date," &
                    " Item_code,Cust_Drgno,Quantity,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,Doc_No,DRG_STATUS,UPLDFILENAME,SCHEDULETYPE)" &
                    " values ('" & txtCustomer.Text & "','" & gstrUNITID & "', '" & strDtl(i) & "'," &
                    " '" & strCreateDate & "','" & strItemCode & "','" & strHdr(15) & "','" & strDtl(i + 2) & "'," &
                    " getdate(), '" & mP_User & "',getdate(), '" & mP_User & "'," & DOC_NO & ",1,'" & strFile & "','F')"

                HoldQueryTable2.Rows.Add(LineNo, "Insert", strHdr(15), strSql)
                'sqlcmd.CommandText = strSql
                'sqlcmd.ExecuteNonQuery()

                i = i + 7
            End While

            Exit Function

        Catch ex As Exception
            sqlcmd.Dispose()
            MessageBox.Show(ex.Message.ToString(), ResolveResString(100), MessageBoxButtons.OK)
            Return False
        Finally
            If Not sqlcmd Is Nothing Then
                sqlcmd.Dispose()
                sqlcmd = Nothing
                If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
            End If

        End Try

    End Function
    Private Function GetValues_FORD_forecast(ByVal DOC_NO As String, ByVal FileNo As Integer) As Boolean
        'REVISED BY     :   Prashant Dhingra
        'REVISION DATE  :   16/06/2011
        'REASON         :   Changes done For RSA (Report Added for Duplicate and Not Defined Item Codes)
        'ISSUE ID       :   10104491
        'REVISED BY     :   Prashant Dhingra
        'REVISED DATE   :   27/06/2011
        'REASON         :   1. New Columns added in ASN Generation 2. Schedule to be uploaded except for missing drawing No.
        'ISSUE ID       :   10108772 
        Dim intStart, intEnd As Integer
        Dim lngLength As Long
        Dim i As Integer = 0
        Dim itemcount As Integer = 0
        Dim sqlcmd As SqlCommand
        Dim sqlCon As SqlConnection
        Dim sqlRdr As SqlDataReader
        Dim strSql As String = ""
        Dim strHdr(48) As String
        'Dim strDtl(300) As String
        Dim strDtl(800) As String
        Dim lngPointer As Long
        Dim chr As Char = String.Empty
        Dim strItemCode As String
        Dim strCreateDate As Date
        Dim strSQLD As String = String.Empty
        Dim blnUpldFileStatus As Boolean
        Dim blnDrgStatus As Boolean
        Dim chrDupORNotDefined As Char
        Dim LineNo As Integer

        Dim k As Integer
        k = 0
        Dim j As Integer
        j = 0
        GetValues_FORD_forecast = True
        Try
            intStart = 0
            If InStr(intStart + 1, Label1.Text, "~") > 0 And InStr(intStart + 1, Label1.Text, ",") > 0 Then
                mstrIncorrectFileFormatMsg = mstrIncorrectFileFormatMsg + "File Name: " + strFile + "{File Contains Both Delimiters (~ and ,)}" + vbCrLf
                Return False
            End If
            chr = ""
            If InStr(intStart + 1, Label1.Text, "~") > 0 Then
                chr = "~"
            End If
            If InStr(intStart + 1, Label1.Text, ",") > 0 Then
                chr = ","
            End If

            If chr = String.Empty Or chr = Nothing Then
                mstrIncorrectFileFormatMsg = mstrIncorrectFileFormatMsg + "File Name: " + strFile + "{File Contains No Delimiter (~ OR ,)}" + vbCrLf
                Return False
            End If


            intEnd = InStr(1, Label1.Text, "DTL") - 1

            lngLength = intEnd
            lngPointer = 0

            While lngPointer <= lngLength
                intEnd = InStr(intStart + 1, Label1.Text, chr)
                If intEnd < intStart Then Exit While
                If LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString <> "HDR" Then
                    strHdr(i) = LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString
                    i = i + 1
                End If
                intStart = intEnd
                lngPointer = intEnd
            End While

            i = i - 1
            strHdr(i) = Mid(strHdr(i), 1, Len(strHdr(i)) - 3)

            lngLength = Label1.Text.Length
            i = 0
            While lngPointer <= lngLength
                intEnd = InStr(intStart + 1, Label1.Text, chr)
                If intEnd = 0 And intStart < lngLength Then intEnd = lngLength + 1
                strDtl(i) = LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString
                i = i + 1
                intStart = intEnd
                lngPointer = intEnd
            End While
            ReDim Preserve strDtl(UBound(strDtl) + 1)
            lngPointer = 0

            sqlCon = SqlConnectionclass.GetConnection
            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()
            sqlcmd = New SqlCommand
            sqlcmd.Connection = sqlCon

            blnUpldFileStatus = True
            blnDrgStatus = True

            'strSql = "select item_code from custitem_mst (Nolock) where UNIT_CODE = '" & gstrUNITID & "' AND account_code = '" & txtCustomer.Text & "'" & _
            '    " and cust_drgno = '" & strHdr(15) & "' and active = 1"
            'sqlcmd.CommandText = strSql
            'sqlRdr = sqlcmd.ExecuteReader

            'itemcount = 0
            'If sqlRdr.HasRows Then
            '    While sqlRdr.Read()
            '        itemcount = itemcount + 1
            '        strItemCode = sqlRdr("item_code").ToString
            '    End While
            '    If itemcount > 1 Then
            '        chrDupORNotDefined = "D"
            '        GetValues_FORD_forecast = False
            '    End If
            'Else
            '    chrDupORNotDefined = "N"
            '    GetValues_FORD_forecast = False
            'End If
            'sqlRdr.Close()

            itemcount = 0
            Dim datarow As DataRow() = ItemCodeTable.Select("cust_drgno ='" & strHdr(15) & "'")
            If datarow Is Nothing Or datarow.Count = 0 Then
                chrDupORNotDefined = "N"
                GetValues_FORD_forecast = False
            Else
                For Each row As DataRow In datarow
                    itemcount = itemcount + 1
                    strItemCode = row("item_code").ToString
                Next
                If (itemcount > 1) Then
                    chrDupORNotDefined = "D"
                    GetValues_FORD_forecast = False
                End If
            End If

            If GetValues_FORD_forecast = False Then
                GetValues_FORD_forecast = True
                mblnReportopen = True
                sqlcmd.Dispose()
                blnUpldFileStatus = False
                blnDrgStatus = False
            End If

            strCreateDate = Mid(strHdr(14), 5, 2) + " / " + Mid(strHdr(14), 3, 2) + " / " + Mid(strHdr(14), 1, 2)

            strSql = "set dateformat 'dmy' insert into RSASchedule_Hdr (Account_Code,Doc_No,Plant_Code,Unit_Code," &
                " CreateDate,Item_code,Cust_Drgno,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,DRG_STATUS,UPLDFILENAME," &
                " SCHEDULETYPE,FileUpld_Status,DupORNotDefined) " &
                " values('" & txtCustomer.Text & "'," & DOC_NO & ",'" & strHdr(5) & "','" & gstrUNITID & "'," &
                " '" & strCreateDate & "','" & strItemCode & "','" & strHdr(15) & "',getdate()," &
                " '" & mP_User & "',getdate(), '" & mP_User & "','" & blnDrgStatus & "','" & strFile & "','F','" & blnUpldFileStatus & "'," &
                " '" & chrDupORNotDefined & "')"

            HoldQueryTable.Rows.Add(FileNo, 1, "Insert", strHdr(15), strSql)

            'sqlcmd.CommandText = strSql
            'sqlcmd.ExecuteNonQuery()

            'If blnUpldFileStatus = False Then
            '    Return False
            '    If Not sqlcmd Is Nothing Then
            '        sqlcmd.Dispose()
            '        sqlcmd = Nothing
            '    End If
            'End If

            lngLength = i - 1
            i = 0
            While i <= lngLength
                LineNo = Convert.ToInt32(strDtl(i).ToString())
                If strDtl(i + 6) = "Nothing" Or strDtl(i + 6) = "" Then Exit While
                strCreateDate = Mid(strDtl(i + 6), 5, 2) + " / " + Mid(strDtl(i + 6), 3, 2) + " / " + Mid(strDtl(i + 6), 1, 2)
                strSql = "set dateformat 'dmy' UPDATE RSASchedule_DTL SET DRG_STATUS = 0, Upd_dt=getdate() WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & DOC_NO & "'" &
                         " AND CUST_DRGNO = '" & strHdr(15) & "' AND DRG_STATUS = 1 and scheduletype = 'F' AND TRANS_DATE = '" & strCreateDate & "'"

                HoldQueryTable2.Rows.Add(FileNo, LineNo, "Update", strHdr(15), strSql)
                'sqlcmd.CommandText = strSql
                'sqlcmd.ExecuteNonQuery()

                strSql = "set dateformat 'dmy' insert into RSASchedule_Dtl(Account_Code,Unit_Code,Line_No,Trans_Date," &
                    " Item_code,Cust_Drgno,Quantity,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,Doc_No,DRG_STATUS,UPLDFILENAME,SCHEDULETYPE)" &
                    " values ('" & txtCustomer.Text & "','" & gstrUNITID & "', '" & strDtl(i) & "'," &
                    " '" & strCreateDate & "','" & strItemCode & "','" & strHdr(15) & "','" & strDtl(i + 2) & "'," &
                    " getdate(), '" & mP_User & "',getdate(), '" & mP_User & "'," & DOC_NO & ",1,'" & strFile & "','F')"

                HoldQueryTable2.Rows.Add(FileNo, LineNo, "Insert", strHdr(15), strSql)
                'sqlcmd.CommandText = strSql
                'sqlcmd.ExecuteNonQuery()

                i = i + 7
            End While

            Exit Function

        Catch ex As Exception
            sqlcmd.Dispose()
            MessageBox.Show(ex.Message.ToString(), ResolveResString(100), MessageBoxButtons.OK)
            Return False
        Finally
            If Not sqlcmd Is Nothing Then
                sqlcmd.Dispose()
                sqlcmd = Nothing
                If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
            End If

        End Try

    End Function

    Private Function GetValues_FORD_firm(ByVal DOC_NO As String) As Boolean
        'REVISED BY     :   Prashant Dhingra
        'REVISION DATE  :   16/06/2011
        'REASON         :   Changes done For RSA (Report Added for Duplicate and Not Defined Item Codes)
        'ISSUE ID       :   10104491
        'REVISED BY     :   Prashant Dhingra
        'REVISED DATE   :   27/06/2011
        'REASON         :   1. New Columns added in ASN Generation 2. Schedule to be uploaded except for missing drawing No.
        'ISSUE ID       :   10108772 
        Dim intStart, intEnd As Integer
        Dim lngLength As Long
        Dim i As Integer = 0
        Dim itemcount As Integer = 0
        Dim sqlcmd As SqlCommand
        Dim sqlCon As SqlConnection
        Dim sqlRdr As SqlDataReader
        Dim strSql As String = ""
        Dim strHdr(48) As String
        Dim strDtl(6000) As String
        Dim lngPointer As Long
        Dim chr As Char = String.Empty
        Dim strItemCode As String
        Dim strCreateDate As Date
        Dim strSQLD As String = String.Empty
        Dim blnUpldFileStatus As Boolean
        Dim blnDrgStatus As Boolean
        Dim chrDupORNotDefined As Char
        Dim strCreateTime As String

        Dim k As Integer
        k = 0
        Dim j As Integer
        j = 0
        GetValues_FORD_firm = True
        Try
            intStart = 0

            If InStr(intStart + 1, Label1.Text, "|") > 0 Then
                chr = "|"
            End If

            If chr = String.Empty Or chr = Nothing Then
                mstrIncorrectFileFormatMsg = mstrIncorrectFileFormatMsg + "File Name: " + strFile + "{Delimiter | is missing}" + vbCrLf
                Exit Function
            End If

            intEnd = InStr(1, Label1.Text, "DTL") - 1

            lngLength = intEnd
            lngPointer = 0

            While lngPointer <= lngLength
                intEnd = InStr(intStart + 1, Label1.Text, chr)
                If intEnd < intStart Then Exit While
                If LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString <> "HDR" Then
                    strHdr(i) = LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString
                    i = i + 1
                End If
                intStart = intEnd
                lngPointer = intEnd
            End While

            i = i - 1
            strHdr(i) = Mid(strHdr(i), 1, Len(strHdr(i)) - 3)

            lngLength = Label1.Text.Length
            i = 0
            While lngPointer <= lngLength
                intEnd = InStr(intStart + 1, Label1.Text, chr)
                If intEnd = 0 And intStart < lngLength Then intEnd = lngLength + 1
                strDtl(i) = LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString
                i = i + 1
                intStart = intEnd
                lngPointer = intEnd
            End While
            ReDim Preserve strDtl(UBound(strDtl) + 1)
            lngPointer = 0

            sqlCon = SqlConnectionclass.GetConnection
            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()

            sqlcmd = New SqlCommand
            sqlcmd.Connection = sqlCon

            blnUpldFileStatus = True
            blnDrgStatus = True

            strSql = "select item_code from custitem_mst (Nolock) where unit_code = '" & gstrUNITID & "' and" &
                " account_code = '" & txtCustomer.Text & "'" &
                " and cust_drgno = '" & strHdr(15) & "' and active = 1"
            sqlcmd.CommandText = strSql
            sqlRdr = sqlcmd.ExecuteReader

            itemcount = 0
            If sqlRdr.HasRows Then
                While sqlRdr.Read()
                    itemcount = itemcount + 1
                    strItemCode = sqlRdr("item_code").ToString
                End While
                If itemcount > 1 Then
                    chrDupORNotDefined = "D"
                    GetValues_FORD_firm = False
                End If
            Else
                chrDupORNotDefined = "N"
                GetValues_FORD_firm = False
            End If
            sqlRdr.Close()
            If GetValues_FORD_firm = False Then
                GetValues_FORD_firm = True
                mblnReportopen = True
                sqlcmd.Dispose()
                blnUpldFileStatus = False
                blnDrgStatus = False
            End If
            strCreateDate = Mid(strHdr(14), 5, 2) + " / " + Mid(strHdr(14), 3, 2) + " / " + Mid(strHdr(14), 1, 2)

            'Added by Shubhra - Dock code and storage location added
            strSql = "set dateformat 'dmy' insert into RSASchedule_Hdr (Account_Code,Doc_No,Plant_Code,Unit_Code," &
                " CreateDate,Item_code,Cust_Drgno,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,DRG_STATUS," &
                " UPLDFILENAME,SCHEDULETYPE,FileUpld_Status,DupORNotDefined,Dock,StorageLocation) " &
                " values('" & txtCustomer.Text & "'," & DOC_NO & ",'" & strHdr(5) & "','" & gstrUNITID & "'," &
                " '" & strCreateDate & "','" & strItemCode & "','" & strHdr(15) & "',getdate()," &
                " '" & mP_User & "',getdate(), '" & mP_User & "','" & blnDrgStatus & "','" & strFile & "','D','" & blnUpldFileStatus & "'," &
                " '" & chrDupORNotDefined & "','" & strHdr(7) & "','" & strHdr(8) & "')"

            sqlcmd.CommandText = strSql
            sqlcmd.ExecuteNonQuery()

            If blnUpldFileStatus = False Then
                Return False
                If Not sqlcmd Is Nothing Then
                    sqlcmd.Dispose()
                    sqlcmd = Nothing
                End If
            End If

            lngLength = i - 1
            i = 0
            While i <= lngLength
                If strDtl(i + 6) = "Nothing" Or strDtl(i + 6) = "" Then Exit While
                strCreateDate = Mid(strDtl(i + 6), 5, 2) + " / " + Mid(strDtl(i + 6), 3, 2) + " / " + Mid(strDtl(i + 6), 1, 2)
                strCreateTime = Mid(strDtl(i + 6), 7, 2) + ":" + Mid(strDtl(i + 6), 9, 2)

                strSql = "UPDATE RSASchedule_DTL SET DRG_STATUS = 0 WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & DOC_NO & "'" &
                    " AND CUST_DRGNO = '" & strHdr(15) & "' AND DRG_STATUS = 1 and scheduletype = 'D' AND TRANS_DATE = '" & strCreateDate & "'"

                sqlcmd.CommandText = strSql
                sqlcmd.ExecuteNonQuery()

                strSql = "set dateformat 'dmy' insert into RSASchedule_Dtl(Account_Code,Unit_Code,Line_No,Trans_Date,Trans_Time," &
                    " Item_code,Cust_Drgno,Quantity,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,Doc_No,DRG_STATUS,UPLDFILENAME,SCHEDULETYPE)" &
                    " values ('" & txtCustomer.Text & "','" & gstrUNITID & "', '" & strDtl(i) & "'," &
                    " '" & strCreateDate & "','" & strCreateTime & "','" & strItemCode & "','" & strHdr(15) & "','" & strDtl(i + 2) & "'," &
                    " getdate(), '" & mP_User & "',getdate(), '" & mP_User & "'," & DOC_NO & ",1,'" & strFile & "','D')"
                sqlcmd.CommandText = strSql
                sqlcmd.ExecuteNonQuery()

                i = i + 7
            End While

            Exit Function

        Catch ex As Exception
            sqlcmd.Dispose()
            txt_FailedToUpload.Text = txt_FailedToUpload.Text & lbl_FailedToUpload.Tag.ToString() & Environment.NewLine
            textBoxColor(txt_ErrorToUpload, lbl_FailedToUpload.Tag.ToString().Substring(lbl_FailedToUpload.Tag.ToString().LastIndexOf("\") + 1), ex)
            'txt_ErrorToUpload.Text = txt_ErrorToUpload.Text & lbl_FailedToUpload.Tag.ToString().Substring(lbl_FailedToUpload.Tag.ToString().LastIndexOf("\") + 1) & Environment.NewLine & "Error:-" & ex.Message.ToString() & Environment.NewLine & Environment.NewLine
            'MessageBox.Show(ex.Message.ToString(), ResolveResString(100), MessageBoxButtons.OK)
            Return False
        Finally
            If Not sqlcmd Is Nothing Then
                sqlcmd.Dispose()
                sqlcmd = Nothing
                If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
            End If

        End Try

    End Function



    Private Function GetValues_BMW(ByVal Doc_No As String) As Boolean
        Dim intStart, intEnd As Integer
        Dim lngLength As Long
        Dim i As Integer = 0
        Dim itemcount As Integer = 0
        Dim sqlcmd As SqlCommand
        Dim sqlRdr As SqlDataReader
        Dim strSql As String = ""
        Dim strHdr(48) As String
        Dim strDtl(300, 6) As String
        Dim lngPointer As Long
        Dim chr As Char = String.Empty
        Dim strItemCode As String
        Dim strCreateDate As Date
        Dim strSQLD As String = String.Empty
        Dim blnUpldFileStatus As Boolean
        Dim blnDrgStatus As Boolean
        Dim k As Integer
        Dim chrDupORNotDefined As Char
        Dim sqlCon As SqlConnection

        k = 0
        Dim j As Integer
        j = 0
        GetValues_BMW = True
        Try
            intStart = 0
            If InStr(intStart + 1, Label1.Text, "~") > 0 And InStr(intStart + 1, Label1.Text, ",") > 0 Then
                'MessageBox.Show("Invalid Format: " + vbCrLf + "File Name: " + strFile + vbCrLf + "File Contains Both Delimiters (~ and ,)", ResolveResString(100), MessageBoxButtons.OK)
                mstrIncorrectFileFormatMsg = mstrIncorrectFileFormatMsg + "File Name: " + strFile + "{File Contains Both Delimiters (~ and ,)}" + vbCrLf
                Return False
            End If
            chr = ""
            If InStr(intStart + 1, Label1.Text, "~") > 0 Then
                chr = "~"
            End If
            If InStr(intStart + 1, Label1.Text, ",") > 0 Then
                chr = ","
            End If

            If chr = String.Empty Or chr = Nothing Then
                mstrIncorrectFileFormatMsg = mstrIncorrectFileFormatMsg + "File Name: " + strFile + "{File Contains No Delimiter (~ OR ,)}" + vbCrLf
                Exit Function
            End If


            intEnd = InStr(1, Label1.Text, "DTL") - 1

            lngLength = intEnd
            lngPointer = 0

            While lngPointer <= lngLength
                intEnd = InStr(intStart + 1, Label1.Text, chr)
                If intEnd < intStart Then Exit While
                If LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString <> "HDR" Then
                    strHdr(i) = LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString
                    i = i + 1
                End If
                intStart = intEnd
                lngPointer = intEnd
            End While

            i = i - 1
            strHdr(i) = Mid(strHdr(i), 1, Len(strHdr(i)) - 3)

            lngLength = Label1.Text.Length
            i = 0 : j = 0
            While lngPointer <= lngLength
                intEnd = InStr(intStart + 1, Label1.Text, chr)
                If intEnd = 0 And intStart < lngLength Then intEnd = lngLength + 1
                strDtl(i, j) = LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString
                If InStr(LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString, "DTL") Then
                    i = i + 1
                    j = -1
                End If
                j = j + 1
                intStart = intEnd
                lngPointer = intEnd
            End While
            '    ReDim Preserve strDtl(UBound(strDtl) + 1, 1)
            lngPointer = 0

            blnUpldFileStatus = True
            blnDrgStatus = True

            sqlCon = SqlConnectionclass.GetConnection
            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()
            sqlcmd = New SqlCommand
            sqlcmd.Connection = sqlCon

            strSql = "select item_code from custitem_mst where UNIT_CODE = '" & gstrUNITID & "' AND account_code = '" & txtCustomer.Text & "'" &
                " and cust_drgno = '" & strHdr(15) & "' and active = 1"
            sqlcmd.CommandText = strSql
            sqlRdr = sqlcmd.ExecuteReader

            itemcount = 0
            If sqlRdr.HasRows Then
                While sqlRdr.Read()
                    itemcount = itemcount + 1
                    strItemCode = sqlRdr("item_code").ToString
                End While
                If itemcount > 1 Then
                    chrDupORNotDefined = "D"
                    GetValues_BMW = False
                End If
            Else
                chrDupORNotDefined = "N"
                GetValues_BMW = False
            End If
            sqlRdr.Close()
            If GetValues_BMW = False Then
                GetValues_BMW = True
                mblnReportopen = True
                sqlcmd.Dispose()
                blnUpldFileStatus = False
                blnDrgStatus = False
            End If

            strCreateDate = Mid(strHdr(14), 7, 2) + " / " + Mid(strHdr(14), 5, 2) + " / " + Mid(strHdr(14), 1, 4)

            strSql = "SELECT * FROM RSASchedule_Hdr WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & Doc_No & "'" &
                " AND CUST_DRGNO = '" & strHdr(15) & "' AND DRG_STATUS = 1"
            sqlcmd.CommandText = strSql
            sqlRdr = sqlcmd.ExecuteReader

            If sqlRdr.HasRows Then
                sqlRdr.Close()
                strSql = "UPDATE RSASchedule_Hdr SET DRG_STATUS = 0 WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & Doc_No & "'" &
                " AND CUST_DRGNO = '" & strHdr(15) & "' AND DRG_STATUS = 1"

                sqlcmd.CommandText = strSql
                sqlcmd.ExecuteNonQuery()

                strSql = "UPDATE RSASchedule_DTL SET DRG_STATUS = 0 WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & Doc_No & "'" &
                " AND CUST_DRGNO = '" & strHdr(15) & "' AND DRG_STATUS = 1"

                sqlcmd.CommandText = strSql
                sqlcmd.ExecuteNonQuery()
            Else
                sqlRdr.Close()
            End If

            strSql = "set dateformat 'dmy' insert into RSASchedule_Hdr (Account_Code,Doc_No,Plant_Code,Unit_Code," &
                " CreateDate,Item_code,Cust_Drgno,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,DRG_STATUS,UPLDFILENAME,SCHEDULETYPE,FileUpld_Status,DupORNotDefined) " &
                " values('" & txtCustomer.Text & "'," & Doc_No & ",'" & strHdr(5) & "','" & gstrUNITID & "'," &
                " '" & strCreateDate & "','" & strItemCode & "','" & strHdr(15) & "',getdate()," &
                " '" & mP_User & "',getdate(), '" & mP_User & "','" & blnDrgStatus & "','" & strFile & "','D','" & blnUpldFileStatus & "','" & chrDupORNotDefined & "')"

            sqlcmd.CommandText = strSql
            sqlcmd.ExecuteNonQuery()

            If blnUpldFileStatus = False Then
                Return False
            End If

            lngLength = i - 1
            i = 0
            While i <= lngLength
                If strDtl(i, 1) = "Nothing" Or strDtl(i, 1) = "" Then Exit While
                strCreateDate = Mid(strDtl(i, 6), 7, 2) + " / " + Mid(strDtl(i, 6), 5, 2) + " / " + Mid(strDtl(i, 6), 1, 4)
                strSql = "set dateformat 'dmy' insert into RSASchedule_Dtl(Account_Code,Unit_Code,Line_No,Trans_Date," &
                    " Item_code,Cust_Drgno,Quantity,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,Doc_No,DRG_STATUS,UPLDFILENAME,SCHEDULETYPE)" &
                    " values ('" & txtCustomer.Text & "','" & gstrUNITID & "', '" & strDtl(i, 0) & "'," &
                    " '" & strCreateDate & "','" & strItemCode & "','" & strHdr(15) & "','" & strDtl(i, 2) & "'," &
                    " getdate(), '" & mP_User & "',getdate(), '" & mP_User & "'," & Doc_No & ",1,'" & strFile & "','D')"
                sqlcmd.CommandText = strSql
                sqlcmd.ExecuteNonQuery()

                i = i + 1
            End While


            Exit Function

        Catch ex As Exception
            sqlcmd.Dispose()
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
            Return False
        Finally
            If Not sqlcmd Is Nothing Then
                sqlcmd.Dispose()
                sqlcmd = Nothing
                If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
            End If

        End Try
    End Function

    Private Function GetValues_BENZ(ByVal Doc_No As String) As Boolean
        Dim intStart, intEnd As Integer
        Dim lngLength As Long
        Dim i As Integer = 0
        Dim itemcount As Integer = 0
        Dim sqlcmd As SqlCommand
        Dim sqlRdr As SqlDataReader
        Dim strSql As String = ""
        Dim strHdr(48) As String
        Dim strDtl(300, 6) As String
        Dim lngPointer As Long
        Dim chr As Char = String.Empty
        Dim strItemCode As String
        Dim strCreateDate As Date
        Dim strSQLD As String = String.Empty
        Dim blnUpldFileStatus As Boolean
        Dim blnDrgStatus As Boolean
        Dim k As Integer
        Dim chrDupORNotDefined As Char
        Dim sqlCon As SqlConnection

        k = 0
        Dim j As Integer
        j = 0
        GetValues_BENZ = True
        Try
            intStart = 0
            If InStr(intStart + 1, Label1.Text, "~") > 0 And InStr(intStart + 1, Label1.Text, ",") > 0 Then
                'MessageBox.Show("Invalid Format: " + vbCrLf + "File Name: " + strFile + vbCrLf + "File Contains Both Delimiters (~ and ,)", ResolveResString(100), MessageBoxButtons.OK)
                mstrIncorrectFileFormatMsg = mstrIncorrectFileFormatMsg + "File Name: " + strFile + "{File Contains Both Delimiters (~ and ,)}" + vbCrLf
                Return False
            End If
            chr = ""
            If InStr(intStart + 1, Label1.Text, "~") > 0 Then
                chr = "~"
            End If
            If InStr(intStart + 1, Label1.Text, ",") > 0 Then
                chr = ","
            End If

            If chr = String.Empty Or chr = Nothing Then
                mstrIncorrectFileFormatMsg = mstrIncorrectFileFormatMsg + "File Name: " + strFile + "{File Contains No Delimiter (~ OR ,)}" + vbCrLf
                Exit Function
            End If


            intEnd = InStr(1, Label1.Text, "DTL") - 1

            lngLength = intEnd
            lngPointer = 0

            While lngPointer <= lngLength
                intEnd = InStr(intStart + 1, Label1.Text, chr)
                If intEnd < intStart Then Exit While
                If LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString <> "HDR" Then
                    strHdr(i) = LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString
                    i = i + 1
                End If
                intStart = intEnd
                lngPointer = intEnd
            End While

            i = i - 1
            strHdr(i) = Mid(strHdr(i), 1, Len(strHdr(i)) - 3)

            lngLength = Label1.Text.Length
            i = 0 : j = 0
            While lngPointer <= lngLength
                intEnd = InStr(intStart + 1, Label1.Text, chr)
                If intEnd = 0 And intStart < lngLength Then intEnd = lngLength + 1
                strDtl(i, j) = LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString
                If InStr(LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString, "DTL") Then
                    i = i + 1
                    j = -1
                End If
                j = j + 1
                intStart = intEnd
                lngPointer = intEnd
            End While
            '    ReDim Preserve strDtl(UBound(strDtl) + 1, 1)
            lngPointer = 0

            blnUpldFileStatus = True
            blnDrgStatus = True

            sqlCon = SqlConnectionclass.GetConnection
            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()
            sqlcmd = New SqlCommand
            sqlcmd.Connection = sqlCon

            strSql = "select item_code from custitem_mst where UNIT_CODE = '" & gstrUNITID & "' AND account_code = '" & txtCustomer.Text & "'" &
                " and cust_drgno = '" & strHdr(15) & "' and active = 1"
            sqlcmd.CommandText = strSql
            sqlRdr = sqlcmd.ExecuteReader

            itemcount = 0
            If sqlRdr.HasRows Then
                While sqlRdr.Read()
                    itemcount = itemcount + 1
                    strItemCode = sqlRdr("item_code").ToString
                End While
                If itemcount > 1 Then
                    chrDupORNotDefined = "D"
                    GetValues_BENZ = False
                End If
            Else
                chrDupORNotDefined = "N"
                GetValues_BENZ = False
            End If
            sqlRdr.Close()
            If GetValues_BENZ = False Then
                GetValues_BENZ = True
                mblnReportopen = True
                sqlcmd.Dispose()
                blnUpldFileStatus = False
                blnDrgStatus = False
            End If

            strCreateDate = Mid(strHdr(14), 7, 2) + " / " + Mid(strHdr(14), 5, 2) + " / " + Mid(strHdr(14), 1, 4)

            strSql = "SELECT * FROM RSASchedule_Hdr WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & Doc_No & "'" &
                " AND CUST_DRGNO = '" & strHdr(15) & "' AND DRG_STATUS = 1"
            sqlcmd.CommandText = strSql
            sqlRdr = sqlcmd.ExecuteReader

            If sqlRdr.HasRows Then
                sqlRdr.Close()
                strSql = "UPDATE RSASchedule_Hdr SET DRG_STATUS = 0 WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & Doc_No & "'" &
                " AND CUST_DRGNO = '" & strHdr(15) & "' AND DRG_STATUS = 1"

                sqlcmd.CommandText = strSql
                sqlcmd.ExecuteNonQuery()

                strSql = "UPDATE RSASchedule_DTL SET DRG_STATUS = 0 WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & Doc_No & "'" &
                " AND CUST_DRGNO = '" & strHdr(15) & "' AND DRG_STATUS = 1"

                sqlcmd.CommandText = strSql
                sqlcmd.ExecuteNonQuery()
            Else
                sqlRdr.Close()
            End If

            strSql = "set dateformat 'dmy' insert into RSASchedule_Hdr (Account_Code,Doc_No,Plant_Code,Unit_Code," &
                " CreateDate,Item_code,Cust_Drgno,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,DRG_STATUS,UPLDFILENAME,SCHEDULETYPE,FileUpld_Status,DupORNotDefined) " &
                " values('" & txtCustomer.Text & "'," & Doc_No & ",'" & strHdr(5) & "','" & gstrUNITID & "'," &
                " '" & strCreateDate & "','" & strItemCode & "','" & strHdr(15) & "',getdate()," &
                " '" & mP_User & "',getdate(), '" & mP_User & "','" & blnDrgStatus & "','" & strFile & "','D','" & blnUpldFileStatus & "','" & chrDupORNotDefined & "')"

            sqlcmd.CommandText = strSql
            sqlcmd.ExecuteNonQuery()

            If blnUpldFileStatus = False Then
                Return False
            End If

            lngLength = i
            i = 0
            While i <= lngLength
                If strDtl(i, 1) = "Nothing" Or strDtl(i, 1) = "" Then Exit While
                strCreateDate = Mid(strDtl(i, 1), 7, 2) + " / " + Mid(strDtl(i, 1), 5, 2) + " / " + Mid(strDtl(i, 1), 1, 4)
                strSql = "set dateformat 'dmy' insert into RSASchedule_Dtl(Account_Code,Unit_Code,Line_No,Trans_Date," &
                    " Item_code,Cust_Drgno,Quantity,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,Doc_No,DRG_STATUS,UPLDFILENAME,SCHEDULETYPE)" &
                    " values ('" & txtCustomer.Text & "','" & gstrUNITID & "', '" & strDtl(i, 0) & "'," &
                    " '" & strCreateDate & "','" & strItemCode & "','" & strHdr(15) & "','" & strDtl(i, 2) & "'," &
                    " getdate(), '" & mP_User & "',getdate(), '" & mP_User & "'," & Doc_No & ",1,'" & strFile & "','D')"
                sqlcmd.CommandText = strSql
                sqlcmd.ExecuteNonQuery()

                i = i + 1
            End While


            Exit Function

        Catch ex As Exception
            sqlcmd.Dispose()
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
            Return False
        Finally
            If Not sqlcmd Is Nothing Then
                sqlcmd.Dispose()
                sqlcmd = Nothing
                If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
            End If

        End Try
    End Function

    Private Function GetValues_NISSAN_Forecast(ByVal DOC_NO As String) As Boolean
        Dim intStart, intEnd As Integer
        Dim lngLength As Long
        Dim i As Integer = 0
        Dim itemcount As Integer = 0
        Dim sqlcmd As SqlCommand
        Dim sqlRdr As SqlDataReader
        Dim strSql As String = ""
        Dim strHdr(48) As String
        Dim strDtl(,) As String
        Dim lngPointer As Long
        Dim chr As Char = String.Empty
        Dim strItemCode As String
        Dim strCustDrgNo As String
        Dim strCreateDate As Date
        Dim strSQLD As String = String.Empty
        Dim STR As String, INTPOS As Integer
        Dim k As Integer
        Dim blnUpldFileStatus As Boolean
        Dim blnDrgStatus As Boolean
        Dim chrDupORNotDefined As Char
        Dim sqlcon As SqlConnection

        k = 0
        Dim j As Integer
        j = 0
        GetValues_NISSAN_Forecast = True
        Try
            intStart = 0
            If InStr(intStart + 1, Label1.Text, "~") > 0 And InStr(intStart + 1, Label1.Text, ",") > 0 Then
                'MessageBox.Show("Invalid Format: " + vbCrLf + "File Name: " + strFile + vbCrLf + "File Contains Both Delimiters (~ and ,)", ResolveResString(100), MessageBoxButtons.OK)
                mstrIncorrectFileFormatMsg = mstrIncorrectFileFormatMsg + "File Name: " + strFile + "{File Contains Both Delimiters (~ and ,)}" + vbCrLf
                Return False
            End If
            chr = ""
            If InStr(intStart + 1, Label1.Text, "~") > 0 Then
                chr = "~"
            End If
            If InStr(intStart + 1, Label1.Text, ",") > 0 Then
                chr = ","
            End If

            If chr = String.Empty Or chr = Nothing Then
                mstrIncorrectFileFormatMsg = mstrIncorrectFileFormatMsg + "File Name: " + strFile + "{File Contains No Delimiter (~ OR ,)}" + vbCrLf
                Exit Function
            End If

            intEnd = InStr(1, Label1.Text, "DTL") - 1
            lngLength = intEnd
            lngPointer = 0

            i = 0
            STR = Label1.Text
            INTPOS = InStr(STR, "HDR")
            If INTPOS > 0 Then
                While INTPOS <= Label1.Text.Length And INTPOS > 0
                    INTPOS = InStr(STR, "HDR")
                    If INTPOS > 0 Then
                        STR = Mid(STR, INTPOS + 3)
                        i = i + 1
                    End If
                End While
            End If

            STR = Label1.Text
            INTPOS = InStr(STR, "DTL")
            If INTPOS > 0 Then
                While INTPOS <= Label1.Text.Length And INTPOS > 0
                    INTPOS = InStr(STR, "DTL")
                    If INTPOS > 0 Then
                        STR = Mid(STR, INTPOS + 3)
                        i = i + 1
                    End If
                End While
            End If

            ReDim strDtl(i, 50)

            lngLength = Label1.Text.Length
            i = -1 : j = 0
            While lngPointer <= lngLength
                intEnd = InStr(intStart + 1, Label1.Text, chr)
                If intEnd = 0 And intStart < lngLength Then intEnd = lngLength + 1
                If LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString <> "HDR" And LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString <> "DTL" Then
                    strDtl(i, j) = LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString
                End If

                If InStr(LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString, "HDR") Or InStr(LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString, "DTL") Then

                    i = i + 1
                    j = -1
                End If
                j = j + 1
                intStart = intEnd
                lngPointer = intEnd
            End While

            lngPointer = 0

            sqlcon = SqlConnectionclass.GetConnection
            If sqlcon.State = ConnectionState.Closed Then sqlcon.Open()
            sqlcmd = New SqlCommand
            sqlcmd.Connection = sqlcon

            lngLength = i
            i = 0

            blnUpldFileStatus = True
            blnDrgStatus = True

            While i <= lngLength
                If Mid(strDtl(i, 0), 1, 6) = "DELINS" Then
                    strSql = "select item_code from custitem_mst where UNIT_CODE = '" & gstrUNITID & "' AND" &
                        " account_code = '" & txtCustomer.Text & "'" &
                        " and cust_drgno = '" & strDtl(i, 15) & "' and active = 1"
                    sqlcmd.CommandText = strSql
                    sqlRdr = sqlcmd.ExecuteReader

                    strCustDrgNo = strDtl(i, 15)
                    itemcount = 0

                    If sqlRdr.HasRows Then
                        While sqlRdr.Read()
                            itemcount = itemcount + 1
                            strItemCode = sqlRdr("item_code").ToString
                        End While
                        If itemcount > 1 Then
                            chrDupORNotDefined = "D"
                            GetValues_NISSAN_Forecast = False
                        End If
                    Else
                        chrDupORNotDefined = "N"
                        GetValues_NISSAN_Forecast = False
                    End If
                    sqlRdr.Close()

                    If GetValues_NISSAN_Forecast = False Then
                        GetValues_NISSAN_Forecast = True
                        mblnReportopen = True
                        sqlcmd.Dispose()
                        blnUpldFileStatus = False
                        blnDrgStatus = False
                    End If

                    If strDtl(i, 14) <> "" Then
                        strCreateDate = Mid(strDtl(i, 14), 7, 2) + " / " + Mid(strDtl(i, 14), 5, 2) + " / " + Mid(strDtl(i, 14), 1, 4)
                    End If

                    strSql = "SELECT * FROM RSASchedule_Hdr WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & DOC_NO & "'" &
                        " AND CUST_DRGNO = '" & strCustDrgNo & "' AND DRG_STATUS = 1 AND SCHEDULETYPE = 'F'"
                    sqlcmd.CommandText = strSql
                    sqlRdr = sqlcmd.ExecuteReader

                    If sqlRdr.HasRows Then
                        sqlRdr.Close()
                        strSql = "UPDATE RSASchedule_Hdr SET DRG_STATUS = 0 WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & DOC_NO & "'" &
                        " AND CUST_DRGNO = '" & strCustDrgNo & "' AND DRG_STATUS = 1 AND SCHEDULETYPE = 'F'"

                        sqlcmd.CommandText = strSql
                        sqlcmd.ExecuteNonQuery()

                        strSql = "UPDATE RSASchedule_DTL SET DRG_STATUS = 0 WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & DOC_NO & "'" &
                        " AND CUST_DRGNO = '" & strCustDrgNo & "' AND DRG_STATUS = 1 AND SCHEDULETYPE = 'F'"

                        sqlcmd.CommandText = strSql
                        sqlcmd.ExecuteNonQuery()
                    Else
                        sqlRdr.Close()
                    End If

                    strSql = "set dateformat 'dmy' insert into RSASchedule_Hdr (Account_Code,Doc_No,Plant_Code,Unit_Code," &
                    " CreateDate,Item_code,Cust_Drgno,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,RAN_No,DRG_STATUS,UPLDFILENAME,SCHEDULETYPE,fileupld_status,DupORNotDefined) " &
                    " values('" & txtCustomer.Text & "'," & DOC_NO & ",'" & strDtl(i, 5) & "','" & gstrUNITID & "'," &
                    " '" & strCreateDate & "','" & strItemCode & "','" & strDtl(i, 15) & "',getdate()," &
                    " '" & mP_User & "',getdate(), '" & mP_User & "','','" & blnDrgStatus & "','" & strFile & "','F','" & blnUpldFileStatus & "','" & chrDupORNotDefined & "')"
                Else
                    If strDtl(i, 1) = "Nothing" Or strDtl(i, 1) = "" Then Exit While
                    strCreateDate = Mid(strDtl(i, 6), 9, 2) + " / " + Mid(strDtl(i, 6), 6, 2) + " / " + Mid(strDtl(i, 6), 1, 4)
                    strSql = "set dateformat 'dmy' insert into RSASchedule_Dtl(Account_Code,Unit_Code,Line_No,Trans_Date," &
                        " Item_code,Cust_Drgno,Quantity,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,Doc_No,DRG_STATUS,UPLDFILENAME,SCHEDULETYPE)" &
                        " values ('" & txtCustomer.Text & "','" & gstrUNITID & "', '" & strDtl(i, 0) & "'," &
                        " '" & strCreateDate & "','" & strItemCode & "','" & strCustDrgNo & "','" & strDtl(i, 2) & "'," &
                        " getdate(), '" & mP_User & "',getdate(), '" & mP_User & "'," & DOC_NO & ",1,'" & strFile & "','F')"
                End If

                If Mid(strDtl(i, 0), 1, 6) = "DELINS" Then
                    sqlcmd.CommandText = strSql
                    sqlcmd.ExecuteNonQuery()
                    If blnUpldFileStatus = False Then
                        Exit Function
                    End If
                Else
                    sqlcmd.CommandText = strSql
                    sqlcmd.ExecuteNonQuery()
                End If

                i = i + 1
            End While

            Exit Function

        Catch ex As Exception

            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
            Return False
        Finally
            If Not sqlcmd Is Nothing Then
                sqlcmd.Dispose()
                sqlcmd = Nothing
                If sqlcon.State = ConnectionState.Open Then sqlcon.Close()
            End If

        End Try
    End Function

    Private Function GetValues_NISSAN_Firm(ByVal DOC_NO As String) As Boolean
        Dim intStart, intEnd As Integer
        Dim lngLength As Long
        Dim i As Integer = 0
        Dim itemcount As Integer = 0
        Dim sqlcmd As SqlCommand
        Dim sqlRdr As SqlDataReader
        Dim strSql As String = ""
        Dim strHdr(48) As String
        Dim strDtl(300, 50) As String
        Dim lngPointer As Long
        Dim chr As Char = String.Empty
        Dim strItemCode As String
        Dim strCustDrgNo As String
        Dim strCreateDate As Date
        Dim strSQLD As String = String.Empty
        Dim blnUpldFileStatus As Boolean
        Dim blnDrgStatus As Boolean
        Dim chrDupORNotDefined As Char
        Dim sqlCon As SqlConnection

        Dim k As Integer

        k = 0
        Dim j As Integer
        j = 0

        GetValues_NISSAN_Firm = True
        Try
            intStart = 0
            If InStr(intStart + 1, Label1.Text, "~") > 0 And InStr(intStart + 1, Label1.Text, ",") > 0 Then
                mstrIncorrectFileFormatMsg = mstrIncorrectFileFormatMsg + "File Name: " + strFile + "{File Contains Both Delimiters (~ and ,)}" + vbCrLf
                Return False
            End If
            chr = ""
            If InStr(intStart + 1, Label1.Text, "~") > 0 Then
                chr = "~"
            End If
            If InStr(intStart + 1, Label1.Text, ",") > 0 Then
                chr = ","
            End If

            If chr = String.Empty Or chr = Nothing Then
                mstrIncorrectFileFormatMsg = mstrIncorrectFileFormatMsg + "File Name: " + strFile + "{File Contains No Delimiter (~ OR ,)}" + vbCrLf
                Exit Function
            End If

            intEnd = InStr(1, Label1.Text, "DTL") - 1
            lngLength = intEnd
            lngPointer = 0

            lngLength = Label1.Text.Length
            i = -1 : j = 0
            While lngPointer <= lngLength
                intEnd = InStr(intStart + 1, Label1.Text, chr)
                If intEnd = 0 And intStart < lngLength Then intEnd = lngLength + 1
                If LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString <> "HDR" And LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString <> "DTL" Then
                    strDtl(i, j) = LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString
                End If
                If InStr(LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString, "HDR") Or InStr(LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString, "DTL") Then
                    i = i + 1
                    j = -1
                End If
                j = j + 1
                intStart = intEnd
                lngPointer = intEnd
            End While
            'ReDim Preserve strDtl(UBound(strDtl) + 1, 1)
            lngPointer = 0

            sqlCon = SqlConnectionclass.GetConnection
            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()

            sqlcmd = New SqlCommand
            sqlcmd.Connection = sqlCon

            lngLength = i
            i = 0

            blnUpldFileStatus = True
            blnDrgStatus = True

            While i <= lngLength
                If Mid(strDtl(i, 0), 1, 1) = "R" Then
                    strSql = "select item_code from custitem_mst where UNIT_CODE = '" & gstrUNITID & "' AND account_code = '" & txtCustomer.Text & "'" &
                        " and cust_drgno = '" & strDtl(i, 15) & "' and active = 1"
                    sqlcmd.CommandText = strSql
                    sqlRdr = sqlcmd.ExecuteReader

                    strCustDrgNo = strDtl(i, 15)
                    itemcount = 0
                    If sqlRdr.HasRows Then
                        While sqlRdr.Read()
                            itemcount = itemcount + 1
                            strItemCode = sqlRdr("item_code").ToString
                        End While
                        If itemcount > 1 Then
                            chrDupORNotDefined = "D"
                            GetValues_NISSAN_Firm = False
                        End If
                    Else
                        chrDupORNotDefined = "N"
                        GetValues_NISSAN_Firm = False
                    End If
                    sqlRdr.Close()

                    If GetValues_NISSAN_Firm = False Then
                        GetValues_NISSAN_Firm = True
                        mblnReportopen = True
                        sqlcmd.Dispose()
                        blnUpldFileStatus = False
                        blnDrgStatus = False
                    End If

                    If strDtl(i, 14) <> "" Then
                        strCreateDate = Mid(strDtl(i, 14), 7, 2) + " / " + Mid(strDtl(i, 14), 5, 2) + " / " + Mid(strDtl(i, 14), 1, 4)
                    End If

                    strSql = "set dateformat 'dmy' insert into RSASchedule_Hdr (Account_Code,Doc_No,Plant_Code,Unit_Code," &
                    " CreateDate,Item_code,Cust_Drgno,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,RAN_No,DRG_STATUS,UPLDFILENAME," &
                    " SCHEDULETYPE,fileupld_status,DupORNotDefined) " &
                    " values('" & txtCustomer.Text & "'," & DOC_NO & ",'" & strDtl(i, 5) & "','" & gstrUNITID & "'," &
                    " '" & strCreateDate & "','" & strItemCode & "','" & strDtl(i, 15) & "',getdate()," &
                    " '" & mP_User & "',getdate(), '" & mP_User & "','','" & blnDrgStatus & "','" & strFile & "','D','" & blnUpldFileStatus & "','" & chrDupORNotDefined & "')"
                Else
                    If strDtl(i, 6) <> "" Then
                        strCreateDate = Mid(strDtl(i, 6), 9, 2) + " - " + Mid(strDtl(i, 6), 6, 2) + " - " + Mid(strDtl(i, 6), 1, 4)
                    End If

                    strSql = "SELECT * FROM RSASchedule_DTL WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & DOC_NO & "'" &
                        " AND CUST_DRGNO = '" & strCustDrgNo & "' AND TRANS_DATE = '" & strCreateDate & "' AND DRG_STATUS = 1 AND SCHEDULETYPE = 'D'"
                    sqlcmd.CommandText = strSql
                    sqlRdr = sqlcmd.ExecuteReader

                    If sqlRdr.HasRows Then
                        sqlRdr.Close()
                        strSql = "UPDATE RSASchedule_DTL SET DRG_STATUS = 0 WHERE UNIT_CODE = '" & gstrUNITID & "' AND  DOC_NO = '" & DOC_NO & "'" &
                        " AND CUST_DRGNO = '" & strCustDrgNo & "' AND DRG_STATUS = 1 AND SCHEDULETYPE = 'D'"

                        sqlcmd.CommandText = strSql
                        sqlcmd.ExecuteNonQuery()
                    Else
                        sqlRdr.Close()
                    End If

                    If blnUpldFileStatus = False Then
                        Return False
                    End If

                    If strDtl(i, 1) = "Nothing" Or strDtl(i, 1) = "" Then Exit While
                    strCreateDate = Mid(strDtl(i, 6), 9, 2) + " - " + Mid(strDtl(i, 6), 6, 2) + " - " + Mid(strDtl(i, 6), 1, 4)
                    strSql = "set dateformat 'dmy' insert into RSASchedule_Dtl(Account_Code,Unit_Code,Line_No,Trans_Date," &
                        " Item_code,Cust_Drgno,Quantity,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,Doc_No,DRG_STATUS,UPLDFILENAME,SCHEDULETYPE)" &
                        " values ('" & txtCustomer.Text & "','" & gstrUNITID & "', '" & strDtl(i, 0) & "'," &
                        " '" & strCreateDate & "','" & strItemCode & "','" & strCustDrgNo & "','" & strDtl(i, 2) & "'," &
                        " getdate(), '" & mP_User & "',getdate(), '" & mP_User & "'," & DOC_NO & ",1,'" & strFile & "','D')"
                End If

                sqlcmd.CommandText = strSql
                sqlcmd.ExecuteNonQuery()
                i = i + 1

            End While
            Exit Function

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
            Return False
        Finally
            If Not sqlcmd Is Nothing Then
                sqlcmd.Dispose()
                sqlcmd = Nothing
                If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
            End If

        End Try
    End Function

    Private Sub frmMKTTRN0069_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            mdifrmMain.CheckFormName = mintFormIndex
            Me.MdiParent = prjMPower.mdifrmMain
            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.ToString, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub frmMKTTRN0069_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        Try
            frmModules.NodeFontBold(Me.Tag) = False
            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.ToString, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        Try
            Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.ToString, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub frmMKTTRN0069_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Call FitToClient(Me, CObj(Panel2), ctlFormHeader1, CObj(Panel1), 250)
            txtCustomer.CausesValidation = False
        Catch ex As Exception
            MessageBox.Show(ex.ToString, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Function MoveFile() As Object

        Dim OBJ_FSO As New Scripting.FileSystemObject
        Dim file As String = Nothing
        Dim upldFiles As Scripting.File
        Dim filearray(0) As Object

        Try
            If mblnfilemove = True Then
                Return Nothing
                Exit Function
            End If

            Dim subFileName As String = ""
            mblnfilemove = False

            OBJ_FSO = Nothing
            OBJ_FSO = New Scripting.FileSystemObject
            OBJ_FSO.GetFolder(mLocalLocation).Attributes = Scripting.FileAttribute.Normal

            If Not OBJ_FSO.FolderExists(mBkpLocation) Then
                MsgBox("Backup Location Not Found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                Return Nothing
                Exit Function
            End If

            If mBkpLocation = mLocalLocation Then
                MsgBox("Source and Destination are Same", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                Return Nothing
                Exit Function
            End If

            lblFileMoveMsg.Text = "Files are moving to backup location"

            'Added by Vinod on 03/06/2011, Issue Id : 10101895
            strCustFileString = GetCustomerFileString()
            'End of Addition
            If OBJ_FSO.GetFolder(mLocalLocation).Files.Count > 0 Then
                For Each upldFiles In OBJ_FSO.GetFolder(mLocalLocation).Files
                    ReDim Preserve filearray(UBound(filearray) + 1)
                    filearray(UBound(filearray)) = Mid(upldFiles.Path, Len(OBJ_FSO.GetFolder(mLocalLocation).Path) + 2, Len(upldFiles.Path))

                    file = mBkpLocation & "\" & filearray(UBound(filearray))

                    If OBJ_FSO.FileExists(mLocalLocation & "\" & filearray(UBound(filearray))) = True Then
                        If OBJ_FSO.FileExists(file) = True Then
                            OBJ_FSO.DeleteFile(file, True)
                        End If
                        subFileName = filearray(UBound(filearray))
                        OBJ_FSO.MoveFile(mLocalLocation & "\" & subFileName, mBkpLocation & "\")
                    Else
                        MsgBox("Source path does not exist", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    End If
                Next upldFiles

                lblFileMoveMsg.Text = "Files Moved To " + mBkpLocation

            End If
            Return Nothing
            Exit Function

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Function

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Dispose()
    End Sub

    Private Function GetCustomerFileString() As String
        'CREATED BY VINOD ON 03/06/2011
        'ISSUE ID : 10101895
        Dim strCustFile As String = ""
        Dim strCompCode As String
        Dim strCustName() As String
        Try

            strCompCode = gstrUNITID

            strCustName = lblCustName.Text.Split(" ")
            strCustFile = strCustName(0) + "_" + strCompCode + "_" 'R_"
            Return strCustFile
        Catch ex As Exception
            Throw ex
        Finally

        End Try

    End Function

    Private Sub btnUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        Try
            Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
            '10558163 BEGIN
            If GetLocation() = True Then
                Call MoveFile()
            End If
            '10558163 END
            Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtdocNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtdocNo.KeyPress
        Try
            Dim KeyAscii As Short = Asc(e.KeyChar)
            Select Case KeyAscii
                Case System.Windows.Forms.Keys.Return
                    Call txtdocNo_Validating(txtCustomer, New System.ComponentModel.CancelEventArgs(False))
                Case 39, 34, 96
                    KeyAscii = 0
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub txtdocNo_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtdocNo.Validating
        Try
            If txtdocNo.Text.Trim.Length = 0 Then Exit Sub

            Dim strsql As String = ""
            Dim oCmd As SqlCommand
            Dim oRdr As SqlDataReader

            lblMessage.Text = ""
            strsql = "select distinct doc_no, account_code " &
                " from rsaschedule_hdr (Nolock) where UNIT_CODE = '" & gstrUNITID & "' AND doc_no = '" & txtdocNo.Text & "' "

            If txtCustomer.Text.Trim.Length > 0 Then
                strsql = strsql + " and account_code = '" & txtCustomer.Text & "'"
            End If

            oCmd = New SqlCommand(strsql, SqlConnectionclass.GetConnection)
            oRdr = oCmd.ExecuteReader
            If oRdr.HasRows Then
                Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
                oRdr.Read()

            Else
                MessageBox.Show("Invalid Doc No.", ResolveResString(100), MessageBoxButtons.OK)
                Exit Sub
            End If
            Exit Sub
        Catch ex As Exception
            Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub cmdDocNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDocNo.Click
        Dim strHelp() As String = Nothing
        Dim strSql As String = ""
        Dim e1 As System.ComponentModel.CancelEventArgs

        Try

            If txtCustomer.Text.Trim.Length > 0 Then
                strSql = "select distinct doc_no, cast(account_code as varchar(8)) account_code, convert(varchar(11),ent_dt,106) as Doc_Date" &
                    " from rsaschedule_hdr (Nolock) where UNIT_CODE = '" & gstrUNITID & "' AND account_code = '" & txtCustomer.Text & "'"
            Else
                strSql = "select distinct doc_no, cast(account_code as varchar(8)) account_code, convert(varchar(11),ent_dt,106) as Doc_Date" &
                    " from rsaschedule_hdr (Nolock)  WHERE UNIT_CODE = '" & gstrUNITID & "'"
            End If

            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "Doc No Help")
            If UBound(strHelp) <> -1 Then
                If strHelp(0) <> "0" Then
                    txtdocNo.Text = strHelp(0)
                    lbl_Date.Text = strHelp(2)
                    Call txtdocNo_Validating(txtdocNo, e1)
                Else
                    MessageBox.Show("No Customer Code Defined", ResolveResString(100), MessageBoxButtons.OK)
                End If
            End If
            Exit Sub
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub cmdPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdPrint.Click
        Try
            Dim intmaxsubreportloop As Integer
            Dim intsubreportloopcounter As Integer
            Dim lngTotalFiles As Long
            Dim lngUploadedFiles As Long
            Dim sqlcmd As SqlCommand
            Dim sqlrdr As SqlDataReader
            Dim strSql As String
            Dim RepDoc As ReportDocument
            Dim Frm As New eMProCrystalReportViewer

            If (String.IsNullOrEmpty(txtCustomer.Text)) Then
                MessageBox.Show("Customer Code field is required")
                Return
            End If

            If (String.IsNullOrEmpty(txtdocNo.Text)) Then
                MessageBox.Show("Document No field is required")
                Return
            End If

            RepDoc = Frm.GetReportDocument()
            Frm.ReportHeader = Me.ctlFormHeader1.HeaderString()

            sqlcmd = New SqlCommand
            sqlcmd.Connection = SqlConnectionclass.GetConnection

            strSql = "select count(upldfilename) count from rsaschedule_hdr (Nolock)  where UNIT_CODE = '" & gstrUNITID & "' AND doc_no = '" & txtdocNo.Text & "' AND convert(varchar(12),cast(ENT_DT AS DATE),106) = '" & lbl_Date.Text & "'"
            sqlcmd.CommandText = strSql
            sqlrdr = sqlcmd.ExecuteReader

            If sqlrdr.HasRows Then
                sqlrdr.Read()
                lngTotalFiles = sqlrdr("count").ToString
            End If

            sqlrdr.Close()
            strSql = "select count(distinct upldfilename) count from rsaschedule_dtl (Nolock)  where UNIT_CODE = '" & gstrUNITID & "' AND doc_no = '" & txtdocNo.Text & "' AND convert(varchar(12),cast(ENT_DT AS DATE),106) = '" & lbl_Date.Text & "'"
            sqlcmd.CommandText = strSql
            sqlrdr = sqlcmd.ExecuteReader

            If sqlrdr.HasRows Then
                sqlrdr.Read()
                lngUploadedFiles = sqlrdr("count").ToString
            End If

            sqlrdr.Close()
            sqlcmd.Dispose()
            sqlcmd = Nothing
            Dim SelectDate As String = DateTypeCast(lbl_Date.Text, "dd MMM yyyy", CultureInfo.InvariantCulture).ToString(format:="yyyy,MM,dd")
            'Dim StartDate As String = Convert.ToDateTime(lbl_Date.Text).ToString(format:="yyyy,MM,dd") + ",00,00,00"
            'Dim EndDate As String = Convert.ToDateTime(lbl_Date.Text).ToString(format:="yyyy,MM,dd") + ",23,59,59"
            Dim RecordSelectionFormulaString As String = "{RSASCHEDULE_HDR.DOC_NO} = " & txtdocNo.Text & " AND {RSASCHEDULE_HDR.UNIT_CODE} = '" & gstrUNITID & "' AND {RSASchedule_Hdr.Ent_dt} = Date(" & SelectDate & ")  AND {RSASchedule_Dtl.Ent_dt} = Date(" & SelectDate & ")"
            With RepDoc
                .Load(My.Application.Info.DirectoryPath & "\Reports\rptFordRSADrawingNoException.rpt")

                .DataDefinition.FormulaFields("Compname").Text = "'" & gstrCOMPANY & "'"
                .DataDefinition.FormulaFields("CompAdd1").Text = "'" & gstr_WRK_ADDRESS1 & "'"
                .DataDefinition.FormulaFields("CompAdd2").Text = "'" & gstr_WRK_ADDRESS2 & "'"
                .DataDefinition.FormulaFields("totFiles").Text = "'" & lngTotalFiles & "'"
                .DataDefinition.FormulaFields("UpldFiles").Text = "'" & lngUploadedFiles & "'"
                '.RecordSelectionFormula = "{RSASCHEDULE_HDR.DOC_NO} = " & txtdocNo.Text & " and {RSASchedule_Dtl.DRG_STATUS}=true  AND {RSASCHEDULE_HDR.UNIT_CODE} = '" & gstrUNITID & "'"
                .RecordSelectionFormula = RecordSelectionFormulaString

                Frm.Show()

                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)

            End With
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)


        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub

    Private Function GetValues_NISSAN_RSA_Firm(ByVal DOC_NO As String) As Boolean
        '10574175
        Dim intStart, intEnd As Integer
        Dim lngLength As Long
        Dim i As Integer = 0
        Dim itemcount As Integer = 0
        Dim sqlcmd As SqlCommand
        Dim sqlRdr As SqlDataReader
        Dim strSql As String = ""
        Dim strHdr(48) As String
        Dim strDtl(300, 50) As String
        Dim lngPointer As Long
        Dim chr As Char = String.Empty
        Dim strItemCode As String
        Dim strCustDrgNo As String
        Dim strCreateDate As Date
        Dim strSQLD As String = String.Empty
        Dim blnUpldFileStatus As Boolean
        Dim blnDrgStatus As Boolean
        Dim chrDupORNotDefined As Char
        Dim sqlCon As SqlConnection
        Dim k As Integer
        Dim j As Integer
        Dim strDate As String = String.Empty

        k = 0
        j = 0

        GetValues_NISSAN_RSA_Firm = True

        Try
            intStart = 0

            'If InStr(intStart + 1, Label1.Text, "|") > 0 And InStr(intStart + 1, Label1.Text, ",") > 0 Then
            '    mstrIncorrectFileFormatMsg = mstrIncorrectFileFormatMsg + "File Name: " + strFile + "{File Contains Both Delimiters (| and ,)}" + vbCrLf
            '    Return False
            'End If

            chr = ""
            If InStr(intStart + 1, Label1.Text, "|") > 0 Then
                chr = "|"
            End If

            If chr = String.Empty Or chr = Nothing Then
                mstrIncorrectFileFormatMsg = mstrIncorrectFileFormatMsg + "File Name: " + strFile + "{File Contains No Delimiter (~ OR ,)}" + vbCrLf
                MessageBox.Show(mstrIncorrectFileFormatMsg, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Function
            End If

            intEnd = InStr(1, Label1.Text, "DTL") - 1
            lngLength = intEnd
            lngPointer = 0
            lngLength = Label1.Text.Length

            i = -1 : j = 0

            While lngPointer <= lngLength
                intEnd = InStr(intStart + 1, Label1.Text, chr)
                If intEnd = 0 And intStart < lngLength Then intEnd = lngLength + 1
                If LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString <> "HDR" And LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString <> "DTL" Then
                    strDtl(i, j) = LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString
                End If
                If InStr(LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString, "HDR") Or InStr(LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString, "DTL") Then
                    i = i + 1
                    j = -1
                End If
                j = j + 1
                intStart = intEnd
                lngPointer = intEnd
            End While

            lngPointer = 0

            sqlCon = SqlConnectionclass.GetConnection
            If sqlCon.State = ConnectionState.Closed Then sqlCon.Open()

            sqlcmd = New SqlCommand
            sqlcmd.Connection = sqlCon

            lngLength = i
            i = 0

            blnUpldFileStatus = True
            blnDrgStatus = True

            While i <= lngLength
                If Mid(strDtl(i, 0), 1, 1) = "O" Then
                    strSql = "SELECT ITEM_CODE FROM CUSTITEM_MST WHERE UNIT_CODE = '" & gstrUNITID & "' AND ACCOUNT_CODE = '" & txtCustomer.Text & "'" &
                        " AND CUST_DRGNO = '" & strDtl(i, 15) & "' AND ACTIVE = 1"
                    sqlcmd.CommandText = strSql
                    sqlRdr = sqlcmd.ExecuteReader

                    strCustDrgNo = strDtl(i, 15)
                    itemcount = 0
                    If sqlRdr.HasRows Then
                        While sqlRdr.Read()
                            itemcount = itemcount + 1
                            strItemCode = sqlRdr("ITEM_CODE").ToString
                        End While
                        If itemcount > 1 Then
                            'IF duplicate
                            chrDupORNotDefined = "D"
                            GetValues_NISSAN_RSA_Firm = False
                        End If
                    Else
                        'IF not defined
                        chrDupORNotDefined = "N"
                        GetValues_NISSAN_RSA_Firm = False
                    End If
                    sqlRdr.Close()

                    If GetValues_NISSAN_RSA_Firm = False Then
                        GetValues_NISSAN_RSA_Firm = True
                        mblnReportopen = True
                        sqlcmd.Dispose()
                        blnUpldFileStatus = False
                        blnDrgStatus = False
                    End If

                    'If blnUpldFileStatus = False Then
                    '    Return False
                    'End If

                    If strDtl(i, 14) <> "" Then
                        strCreateDate = Mid(strDtl(i, 14), 5, 2) + " / " + Mid(strDtl(i, 14), 3, 2) + " / " + Mid(strDtl(i, 14), 1, 2)
                    End If

                    strSql = "SET DATEFORMAT 'DMY' INSERT INTO RSASCHEDULE_HDR (ACCOUNT_CODE,DOC_NO,PLANT_CODE,UNIT_CODE," &
                    " CREATEDATE,ITEM_CODE,CUST_DRGNO,ENT_DT,ENT_USERID,UPD_DT,UPD_USERID,RAN_NO,DRG_STATUS,UPLDFILENAME," &
                    " SCHEDULETYPE,FILEUPLD_STATUS,DUPORNOTDEFINED) " &
                    " VALUES('" & txtCustomer.Text & "'," & DOC_NO & ",'" & strDtl(i, 6) & "','" & gstrUNITID & "'," &
                    " '" & strCreateDate & "','" & strItemCode & "','" & strDtl(i, 15) & "',GETDATE()," &
                    " '" & mP_User & "',GETDATE(), '" & mP_User & "','','" & blnDrgStatus & "','" & strFile & "','D','" & blnUpldFileStatus & "','" & chrDupORNotDefined & "')"
                Else
                    If strDtl(i, 6) <> "" Then
                        If InStr(1, strDtl(i, 6), "DTL") > 0 Or InStr(1, strDtl(i, 6), "HDR") > 0 Then
                            strDate = strDtl(i, 6).Substring(0, Len(strDtl(i, 6)) - 3)
                        Else
                            strDate = strDtl(i, 6)
                        End If

                        If Len(strDate) > 7 Then
                            strCreateDate = Mid(strDtl(i, 6), 7, 2) + " / " + Mid(strDtl(i, 6), 5, 2) + " / " + Mid(strDtl(i, 6), 1, 4)
                        Else
                            strCreateDate = Mid(strDtl(i, 6), 5, 2) + " / " + Mid(strDtl(i, 6), 3, 2) + " / " + Mid(strDtl(i, 6), 1, 2)
                        End If
                    End If

                    strSql = "SELECT * FROM RSASCHEDULE_DTL WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & DOC_NO & "'" &
                        " AND CUST_DRGNO = '" & strCustDrgNo & "' AND TRANS_DATE = '" & strCreateDate & "' AND DRG_STATUS = 1 AND SCHEDULETYPE = 'D'"
                    sqlcmd.CommandText = strSql
                    sqlRdr = sqlcmd.ExecuteReader

                    If sqlRdr.HasRows Then
                        sqlRdr.Close()
                        strSql = "UPDATE RSASCHEDULE_DTL SET DRG_STATUS = 0 WHERE UNIT_CODE = '" & gstrUNITID & "' AND  DOC_NO = '" & DOC_NO & "'" &
                        " AND CUST_DRGNO = '" & strCustDrgNo & "' AND DRG_STATUS = 1 AND SCHEDULETYPE = 'D'"

                        sqlcmd.CommandText = strSql
                        sqlcmd.ExecuteNonQuery()
                    Else
                        sqlRdr.Close()
                    End If

                    If blnUpldFileStatus = False Then
                        Return False
                    End If

                    If strDtl(i, 1) = "Nothing" Or strDtl(i, 1) = "" Then Exit While

                    If InStr(1, strDtl(i, 6), "DTL") > 0 Or InStr(1, strDtl(i, 6), "HDR") > 0 Then
                        strDate = strDtl(i, 6).Substring(0, Len(strDtl(i, 6)) - 3)
                    Else
                        strDate = strDtl(i, 6)
                    End If

                    If Len(strDate) > 7 Then
                        strCreateDate = Mid(strDtl(i, 6), 7, 2) + " / " + Mid(strDtl(i, 6), 5, 2) + " / " + Mid(strDtl(i, 6), 1, 4)
                    Else
                        strCreateDate = Mid(strDtl(i, 6), 5, 2) + " / " + Mid(strDtl(i, 6), 3, 2) + " / " + Mid(strDtl(i, 6), 1, 2)
                    End If

                    strSql = "SET DATEFORMAT 'DMY' INSERT INTO RSASCHEDULE_DTL(ACCOUNT_CODE,UNIT_CODE,LINE_NO,TRANS_DATE," &
                        " ITEM_CODE,CUST_DRGNO,QUANTITY,ENT_DT,ENT_USERID,UPD_DT,UPD_USERID,DOC_NO,DRG_STATUS,UPLDFILENAME,SCHEDULETYPE)" &
                        " VALUES ('" & txtCustomer.Text & "','" & gstrUNITID & "', '" & strDtl(i, 0) & "'," &
                        " '" & strCreateDate & "','" & strItemCode & "','" & strCustDrgNo & "','" & strDtl(i, 2) & "'," &
                        " GETDATE(), '" & mP_User & "',GETDATE(), '" & mP_User & "'," & DOC_NO & ",1,'" & strFile & "','D')"
                End If

                sqlcmd.CommandText = strSql
                sqlcmd.ExecuteNonQuery()
                i = i + 1

            End While
            Exit Function

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
            Return False
        Finally
            If Not sqlcmd Is Nothing Then
                sqlcmd.Dispose()
                sqlcmd = Nothing
                If sqlCon.State = ConnectionState.Open Then sqlCon.Close()
            End If

        End Try
    End Function

    Private Function GetValues_NISSAN_RSA_Forecast(ByVal DOC_NO As String) As Boolean
        '10574175
        Dim intStart, intEnd As Integer
        Dim lngLength As Long
        Dim i As Integer = 0
        Dim itemcount As Integer = 0
        Dim sqlcmd As SqlCommand
        Dim sqlRdr As SqlDataReader
        Dim strSql As String = ""
        Dim strHdr(48) As String
        Dim strDtl(,) As String
        Dim lngPointer As Long
        Dim chr As Char = String.Empty
        Dim strItemCode As String
        Dim strCustDrgNo As String
        Dim strCreateDate As Date
        Dim strSQLD As String = String.Empty
        Dim STR As String, INTPOS As Integer
        Dim k As Integer
        Dim blnUpldFileStatus As Boolean
        Dim blnDrgStatus As Boolean
        Dim chrDupORNotDefined As Char = ""
        Dim sqlcon As SqlConnection
        Dim j As Integer

        k = 0
        j = 0
        GetValues_NISSAN_RSA_Forecast = True

        Try
            intStart = 0
            'If InStr(intStart + 1, Label1.Text, "|") > 0 And InStr(intStart + 1, Label1.Text, ",") > 0 Then
            '    mstrIncorrectFileFormatMsg = mstrIncorrectFileFormatMsg + "File Name: " + strFile + "{File Contains Both Delimiters (| and ,)}" + vbCrLf
            '    Return False
            'End If

            chr = ""
            If InStr(intStart + 1, Label1.Text, "|") > 0 Then
                chr = "|"
            End If

            If InStr(intStart + 1, Label1.Text, "|") > 0 Then
                chr = "|"
            End If

            If chr = String.Empty Or chr = Nothing Then
                mstrIncorrectFileFormatMsg = mstrIncorrectFileFormatMsg + "File Name: " + strFile + "{File Contains No Delimiter (| OR ,)}" + vbCrLf
                MessageBox.Show(mstrIncorrectFileFormatMsg, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Exit Function
            End If

            intEnd = InStr(1, Label1.Text, "DTL") - 1
            lngLength = intEnd
            lngPointer = 0

            i = 0
            STR = Label1.Text
            INTPOS = InStr(STR, "HDR")
            If INTPOS > 0 Then
                While INTPOS <= Label1.Text.Length And INTPOS > 0
                    INTPOS = InStr(STR, "HDR")
                    If INTPOS > 0 Then
                        STR = Mid(STR, INTPOS + 3)
                        i = i + 1
                    End If
                End While
            End If

            STR = Label1.Text
            INTPOS = InStr(STR, "DTL")
            If INTPOS > 0 Then
                While INTPOS <= Label1.Text.Length And INTPOS > 0
                    INTPOS = InStr(STR, "DTL")
                    If INTPOS > 0 Then
                        STR = Mid(STR, INTPOS + 3)
                        i = i + 1
                    End If
                End While
            End If

            ReDim strDtl(i, 50)

            lngLength = Label1.Text.Length
            i = -1 : j = 0
            While lngPointer <= lngLength
                intEnd = InStr(intStart + 1, Label1.Text, chr)
                If intEnd = 0 And intStart < lngLength Then intEnd = lngLength + 1
                If LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString <> "HDR" And LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString <> "DTL" Then
                    strDtl(i, j) = LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString
                End If

                If InStr(LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString, "HDR") Or InStr(LTrim(RTrim(Mid(Label1.Text, intStart + 1, intEnd - intStart - 1))).ToString, "DTL") Then
                    i = i + 1
                    j = -1
                End If
                j = j + 1

                intStart = intEnd
                lngPointer = intEnd
            End While

            lngPointer = 0

            sqlcon = SqlConnectionclass.GetConnection
            If sqlcon.State = ConnectionState.Closed Then sqlcon.Open()
            sqlcmd = New SqlCommand
            sqlcmd.Connection = sqlcon

            lngLength = i
            i = 0

            blnUpldFileStatus = True
            blnDrgStatus = True

            While i <= lngLength
                If Mid(strDtl(i, 0), 1, 6) = "DELINS" Then
                    strSql = "select item_code from custitem_mst where UNIT_CODE = '" & gstrUNITID & "' AND" &
                        " account_code = '" & txtCustomer.Text & "'" &
                        " and cust_drgno = '" & strDtl(i, 15) & "' and active = 1"
                    sqlcmd.CommandText = strSql
                    sqlRdr = sqlcmd.ExecuteReader

                    strCustDrgNo = strDtl(i, 15)
                    itemcount = 0

                    If sqlRdr.HasRows Then
                        While sqlRdr.Read()
                            itemcount = itemcount + 1
                            strItemCode = sqlRdr("item_code").ToString
                        End While
                        If itemcount > 1 Then
                            chrDupORNotDefined = "D"
                            GetValues_NISSAN_RSA_Forecast = False
                        End If
                    Else
                        strItemCode = ""
                        chrDupORNotDefined = "N"
                        GetValues_NISSAN_RSA_Forecast = False
                    End If
                    sqlRdr.Close()

                    If GetValues_NISSAN_RSA_Forecast = False Then
                        GetValues_NISSAN_RSA_Forecast = True
                        mblnReportopen = True
                        sqlcmd.Dispose()
                        blnUpldFileStatus = False
                        blnDrgStatus = False
                    End If

                    ''If blnUpldFileStatus = False Then
                    ''    Return False
                    ''End If

                    If strDtl(i, 13) <> "" Then
                        strCreateDate = Mid(strDtl(i, 13), 5, 2) + " / " + Mid(strDtl(i, 13), 3, 2) + " / " + Mid(strDtl(i, 13), 1, 2)
                    End If

                    strSql = "SELECT * FROM RSASchedule_Hdr WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & DOC_NO & "'" &
                        " AND CUST_DRGNO = '" & strCustDrgNo & "' AND DRG_STATUS = 1 AND SCHEDULETYPE = 'F'"
                    sqlcmd.CommandText = strSql
                    sqlRdr = sqlcmd.ExecuteReader

                    If sqlRdr.HasRows Then
                        sqlRdr.Close()
                        strSql = "UPDATE RSASchedule_Hdr SET DRG_STATUS = 0 WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & DOC_NO & "'" &
                        " AND CUST_DRGNO = '" & strCustDrgNo & "' AND DRG_STATUS = 1 AND SCHEDULETYPE = 'F'"

                        sqlcmd.CommandText = strSql
                        sqlcmd.ExecuteNonQuery()

                        strSql = "UPDATE RSASchedule_DTL SET DRG_STATUS = 0 WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = '" & DOC_NO & "'" &
                        " AND CUST_DRGNO = '" & strCustDrgNo & "' AND DRG_STATUS = 1 AND SCHEDULETYPE = 'F'"

                        sqlcmd.CommandText = strSql
                        sqlcmd.ExecuteNonQuery()
                    Else
                        sqlRdr.Close()
                    End If

                    strSql = "set dateformat 'dmy' insert into RSASchedule_Hdr (Account_Code,Doc_No,Plant_Code,Unit_Code,"
                    strSql = strSql + " CreateDate,Item_code,Cust_Drgno,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,RAN_No,DRG_STATUS,UPLDFILENAME,SCHEDULETYPE,fileupld_status,DupORNotDefined)"
                    strSql = strSql + " values('" & txtCustomer.Text & "'," & DOC_NO & ",'" & strDtl(i, 5) & "','" & gstrUNITID & "',"
                    strSql = strSql + " '" & strCreateDate & "','" & strItemCode & "','" & strDtl(i, 15) & "',getdate(),"
                    strSql = strSql + "'" & mP_User & "',getdate(), '" & mP_User & "','','" & blnDrgStatus & "','" & strFile & "','F','" & blnUpldFileStatus & "',"
                    strSql = strSql + " '" & chrDupORNotDefined & "'"
                    strSql = strSql + ")"

                    'strSql = "set dateformat 'dmy' insert into RSASchedule_Hdr (Account_Code,Doc_No,Plant_Code,Unit_Code," & _
                    '" CreateDate,Item_code,Cust_Drgno,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,RAN_No,DRG_STATUS,UPLDFILENAME,SCHEDULETYPE,fileupld_status,DupORNotDefined) " & _
                    '" values('" & txtCustomer.Text & "'," & DOC_NO & ",'" & strDtl(i, 5) & "','" & gstrUNITID & "'," & _
                    '" '" & strCreateDate & "','" & strItemCode & "','" & strDtl(i, 15) & "',getdate()," & _
                    '" '" & mP_User & "',getdate(), '" & mP_User & "','','" & blnDrgStatus & "','" & strFile & "','F','" & blnUpldFileStatus & "','" & CStr(chrDupORNotDefined) & "')"
                Else
                    If strDtl(i, 1) = "Nothing" Or strDtl(i, 1) = "" Then Exit While
                    Dim strDate As String = String.Empty
                    'strCreateDate = Mid(strDtl(i, 6), 9, 2) + " / " + Mid(strDl(i, 6), 6, 2) + " / " + Mid(strDtl(i, 6), 1, 4)
                    If strDtl(i, 6) <> "" Then
                        If InStr(1, strDtl(i, 6), "DTL") > 0 Or InStr(1, strDtl(i, 6), "HDR") > 0 Then
                            strDate = strDtl(i, 6).Substring(0, Len(strDtl(i, 6)) - 3)
                        Else
                            strDate = strDtl(i, 6)
                        End If

                        If Len(strDate) > 7 Then
                            strCreateDate = Mid(strDtl(i, 6), 7, 2) + " / " + Mid(strDtl(i, 6), 5, 2) + " / " + Mid(strDtl(i, 6), 1, 4)
                        Else
                            strCreateDate = Mid(strDtl(i, 6), 5, 2) + " / " + Mid(strDtl(i, 6), 3, 2) + " / " + Mid(strDtl(i, 6), 1, 2)
                        End If
                    End If

                    strSql = "set dateformat 'dmy' insert into RSASchedule_Dtl(Account_Code,Unit_Code,Line_No,Trans_Date," &
                        " Item_code,Cust_Drgno,Quantity,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,Doc_No,DRG_STATUS,UPLDFILENAME,SCHEDULETYPE)" &
                        " values ('" & txtCustomer.Text & "','" & gstrUNITID & "', '" & strDtl(i, 0) & "'," &
                        " '" & strCreateDate & "','" & strItemCode & "','" & strCustDrgNo & "','" & strDtl(i, 2) & "'," &
                        " getdate(), '" & mP_User & "',getdate(), '" & mP_User & "'," & DOC_NO & ",1,'" & strFile & "','F')"
                End If

                If Mid(strDtl(i, 0), 1, 6) = "DELINS" Then
                    sqlcmd.CommandText = strSql
                    sqlcmd.ExecuteNonQuery()
                    'If blnUpldFileStatus = False Then
                    '    Exit Function
                    'End If
                Else
                    sqlcmd.CommandText = strSql
                    sqlcmd.ExecuteNonQuery()
                End If

                i = i + 1
            End While

            Exit Function

        Catch ex As Exception

            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
            Return False
        Finally
            If Not sqlcmd Is Nothing Then
                sqlcmd.Dispose()
                sqlcmd = Nothing
                If sqlcon.State = ConnectionState.Open Then sqlcon.Close()
            End If

        End Try
    End Function


    Private Function TableHoldQuery() As DataTable
        ' Create a new DataTable
        Dim dataTable As New DataTable()

        ' Define columns
        dataTable.Columns.Add("Qtype", GetType(String))
        dataTable.Columns.Add("cust_drgno", GetType(String))
        dataTable.Columns.Add("Query", GetType(String))

        Return dataTable
    End Function

    Private Function TableHoldQueryWithLineNo() As DataTable
        ' Create a new DataTable
        Dim dataTable As New DataTable()

        ' Define columns
        dataTable.Columns.Add("FileNo", GetType(Integer))
        dataTable.Columns.Add("LineNo", GetType(Integer))
        dataTable.Columns.Add("Qtype", GetType(String))
        dataTable.Columns.Add("cust_drgno", GetType(String))
        dataTable.Columns.Add("Query", GetType(String))

        Return dataTable
    End Function

    Function DateTypeCast(ByVal dateString As String, ByVal format As String, ByVal provider As CultureInfo) As DateTime
        Dim result As DateTime
        Try
            result = DateTime.ParseExact(dateString, format, provider)
        Catch e As FormatException
            result = DateTime.Now
        End Try
        Return result
    End Function

    Sub textBoxColor(ByRef txtbox As RichTextBox, ByVal filename As String, ByVal err As Exception)

        txtbox.SelectionStart = txtbox.TextLength
        txtbox.SelectionLength = 0
        txtbox.SelectionColor = Color.Black
        txtbox.AppendText(filename & Environment.NewLine)

        txtbox.SelectionStart = txtbox.TextLength
        txtbox.SelectionLength = 0
        txtbox.SelectionColor = Color.Red
        txtbox.AppendText("Error:-")

        txtbox.SelectionStart = txtbox.TextLength
        txtbox.SelectionLength = 0
        txtbox.SelectionFont = New Font(txtbox.Font, FontStyle.Bold)
        txtbox.SelectionColor = Color.Blue
        txtbox.AppendText(err.Message.ToString() & Environment.NewLine & Environment.NewLine)

        ' Reset the the default
        txtbox.SelectionColor = txtbox.ForeColor
        txtbox.SelectionFont = New Font(txtbox.Font, FontStyle.Regular)
    End Sub

End Class