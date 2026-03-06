'***************************************************************************************
'COPYRIGHT(C)   : MOTHERSON SUMI INFOTECH & DESIGN LTD.
'FORM NAME      : frmMKTTRN0078
'DESCRIPTION    : FTP FILE REPROCESSING OF INVOICES
'CREATED BY     : VINOD SINGH 
'CREATED DATE   : 14 FEB 2013 (10277476)
'***************************************************************************************
Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Public Class frmMKTTRN0078
    Dim mintFormIndex As Integer

#Region "Form Level Events"

    Private Sub frmMKTTRN0078_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            mdifrmMain.CheckFormName = mintFormIndex
            frmModules.NodeFontBold(Tag) = True
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End Try
    End Sub

    Private Sub frmMKTTRN0078_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        Try
            frmModules.NodeFontBold(Tag) = False
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End Try
    End Sub

    Private Sub frmMKTTRN0078_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            Me.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub frmMKTTRN0078_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Call FitToClient(Me, grpMain, ctlHeader, Me.cmdGrp, )
            Me.MdiParent = mdifrmMain
            mintFormIndex = mdifrmMain.AddFormNameToWindowList(Me.ctlHeader.Tag)
            cmdGrp.Caption(0) = "Create"
            SetListViewHeaders()
            optAllCust.Checked = True
            optAllInv.Checked = True
            dtFrom.Format = DateTimePickerFormat.Custom
            dtFrom.CustomFormat = gstrDateFormat
            dtFrom.Value = GetServerDate()
            dtTo.Format = DateTimePickerFormat.Custom
            dtTo.CustomFormat = gstrDateFormat
            dtTo.Value = GetServerDate()
            pgBar.Maximum = 0

        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End Try
    End Sub

#End Region

#Region "General Routines"

    Private Sub SetListViewHeaders()
        Try
            With lvwCustomer
                .Items.Clear()
                .Columns.Clear()
                .View = View.Details
                .CheckBoxes = True
                .FullRowSelect = True
                .GridLines = True
                .MultiSelect = False
                .Columns.Insert(0, "Customer Code", 180, HorizontalAlignment.Center)
                .Columns.Insert(1, "Customer Name", 265, HorizontalAlignment.Left)
            End With

            With lvwInvoices
                .Items.Clear()
                .Columns.Clear()
                .View = View.Details
                .CheckBoxes = True
                .FullRowSelect = True
                .GridLines = True
                .MultiSelect = False
                .Columns.Insert(0, "Invoice No.", 100, HorizontalAlignment.Center)
                .Columns.Insert(1, "Invoice Date", 145, HorizontalAlignment.Left)
                .Columns.Insert(2, "Customer Code", 200, HorizontalAlignment.Left)
            End With

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub FillCustomer()
        Dim strSQL As String = String.Empty
        Dim dtCust As DataTable
        Dim lvwItem As ListViewItem
        Try
            strSQL = "SELECT CUSTOMER_CODE,CUST_NAME FROM CUSTOMER_MST WHERE SCHEDULECODE IS NOT NULL AND UNIT_CODE ='" & gstrUNITID & "'"
            dtCust = SqlConnectionclass.GetDataTable(strSQL)
            lvwCustomer.Items.Clear()
            If dtCust.Rows.Count > 0 Then
                For Each Row As DataRow In dtCust.Rows
                    lvwItem = New ListViewItem(Convert.ToString(Row("CUSTOMER_CODE")))
                    lvwItem.SubItems.Add(Convert.ToString(Row("CUST_NAME")))
                    lvwCustomer.Items.Add(lvwItem)
                Next
            Else
                MsgBox("No Record Found", MsgBoxStyle.Information, ResolveResString(100))
                Return
            End If
        Catch ex As Exception
            Throw ex
        Finally
            If IsNothing(dtCust) = False Then dtCust.Dispose()
            dtCust = Nothing
            lvwItem = Nothing
        End Try
    End Sub

    Private Sub FillInvoices()
        Dim strSQL As String = String.Empty
        Dim dtInv As DataTable
        Dim lvwItem As ListViewItem
        Dim strCustomer As String = String.Empty
        Try
            If optAllCust.Checked Then

                strSQL = "SELECT DOC_NO,INVOICE_DATE,ACCOUNT_CODE AS CUSTOMER_CODE FROM SALESCHALLAN_DTL A WHERE BILL_FLAG=1 AND CANCEL_FLAG=0 AND UNIT_CODE='" & gstrUNITID & "' " & _
                         " AND INVOICE_DATE BETWEEN '" & getDateForDB(dtFrom.Value) & "'  AND '" & getDateForDB(dtTo.Value) & "' " & _
                         " AND EXISTS(SELECT TOP 1 1 FROM CUSTOMER_MST WHERE UNIT_CODE =A.UNIT_CODE AND CUSTOMER_CODE = A.ACCOUNT_CODE AND SCHEDULECODE IS NOT NULL) " & _
                         " ORDER BY INVOICE_DATE "

            Else
                strCustomer = ""
                If lvwCustomer.Items.Count > 0 Then
                    For Each lvwItem In lvwCustomer.Items
                        If lvwItem.Checked Then
                            strCustomer += "'" + lvwItem.Text + "',"
                        End If
                    Next
                Else
                    Return
                End If
                If strCustomer.Trim.Length > 0 Then
                    strCustomer = "(" + strCustomer.Substring(0, strCustomer.Trim.Length - 1) + ")"
                End If
                strSQL = "SELECT DOC_NO,CONVERT(VARCHAR(12),INVOICE_DATE,103) AS INVOICE_DATE,ACCOUNT_CODE AS CUSTOMER_CODE FROM SALESCHALLAN_DTL WHERE BILL_FLAG=1 AND CANCEL_FLAG=0 AND UNIT_CODE='" & gstrUNITID & "' " & _
                         " AND INVOICE_DATE BETWEEN '" & getDateForDB(dtFrom.Value) & "'  AND '" & getDateForDB(dtTo.Value) & "' " & _
                         " AND ACCOUNT_CODE IN " + strCustomer
            End If
            dtInv = SqlConnectionclass.GetDataTable(strSQL)
            lvwInvoices.Items.Clear()
            If dtInv.Rows.Count > 0 Then
                For Each Row As DataRow In dtInv.Rows
                    lvwItem = New ListViewItem(Convert.ToString(Row("DOC_NO")))
                    lvwItem.SubItems.Add(Convert.ToString(Row("INVOICE_DATE")))
                    lvwItem.SubItems.Add(Convert.ToString(Row("CUSTOMER_CODE")))
                    lvwInvoices.Items.Add(lvwItem)
                Next
            Else
                MsgBox("No Record Found", MsgBoxStyle.Information, ResolveResString(100))
                Return
            End If
        Catch ex As Exception
            Throw ex
        Finally
            If IsNothing(dtInv) = False Then dtInv.Dispose()
            dtInv = Nothing
            lvwItem = Nothing
        End Try
    End Sub

    Public Function CheckInvoices() As Boolean

        Dim minv As Object


        Dim strSBUFolder As String = String.Empty
        Dim strHdr, strSql, mSuff As Object
        Dim strRec, strDtl, mfile As Object
        Dim sFolderName, VCode, strUserMsg As String
        Dim rs_hdr As ClsResultSetDB
        Dim rs_dtl As ClsResultSetDB
        Dim nInvWrite, intInvoice, intDSWrite As Short
        Dim mValue As Double
        Dim invnotopost As String = String.Empty
        Dim objGetDSData As ClsResultSetDB
        Dim varDSFile As Object
        Dim strDSFileData As String = String.Empty
        Dim objFSO As New Scripting.FileSystemObject
        Dim strInvList As String = String.Empty
        Dim strEDIFolder As String = String.Empty
        Dim minvEDIFile As String = String.Empty
        Dim mdsEDIFile As String = String.Empty
        Dim intdsEDIFile As Short
        Dim intinvEDIFile As Short
        Dim strLogMsg As String = String.Empty
        Dim pstrTempPath As String = String.Empty
        Dim pstrLocalDSPath As String = String.Empty
        Dim pstrTempPathForEDI As String = String.Empty
        Dim strFTPPath As String = String.Empty
        Dim strFTPEDIPATH As String = String.Empty
        Dim strFTPPATHforlog As String = String.Empty
        Dim blnFTPwithEDI As Boolean
        Dim blnNewData As Boolean
        Dim strBuffer(14) As String
        Dim nInFile As Short 'File Handle of the arguments file
        Dim strInvoiceNo As String, strAccountCode As String
        Dim blnRecordExists As Boolean = False

        Try
            If gstrUNITID = "SML" Then
                VCode = "S073"
            ElseIf gstrUNITID = "M3E" Then
                VCode = "M582"
            ElseIf gstrUNITID = "MST" Then
                VCode = "M581"
            End If

            nInFile = FreeFile()


            strFTPPath = ReadValueFromINI(Application.StartupPath & "\MultipleFTpArgument.cfg", "FTPPATH-" & gstrUNITID, "FilepathFTP")
            strFTPEDIPATH = ReadValueFromINI(Application.StartupPath & "\MultipleFTpArgument.cfg", "FTPPATH-" & gstrUNITID, "FilepathforFTPEDI")
            strFTPPATHforlog = ReadValueFromINI(Application.StartupPath & "\MultipleFTpArgument.cfg", "FTPPATH-" & gstrUNITID, "FilepathforFTPlog")

            'Check whether folder c:\temp exist or not for EDI
            If Not Directory.Exists(strFTPPath & "\Invoices") Then
                Directory.CreateDirectory(strFTPPath & "\Invoices")
            End If
            'If Dir(pstrTempPathForEDI & "\Invoices\", FileAttribute.Directory) = "" Then MkDir(pstrTempPathForEDI & "\Invoices\")
            If Not Directory.Exists(strFTPEDIPATH & "\Invoices") Then
                Directory.CreateDirectory(strFTPEDIPATH & "\Invoices")
            End If

            Dim rsGetAmount As ClsResultSetDB
            Dim stramt As String

            pgBar.Maximum = 0
            If optSelectedInvoices.Checked Then
                If lvwInvoices.Items.Count = 0 Then
                    Return False
                End If
                pgBar.Maximum = lvwInvoices.Items.Count
                For Each lvwItem As ListViewItem In lvwInvoices.Items

                    If lvwItem.Checked = False Then
                        pgBar.PerformStep()
                        Continue For
                    End If

                    blnRecordExists = True
                    strInvoiceNo = ""
                    strAccountCode = ""
                    strInvoiceNo = lvwItem.Text
                    strAccountCode = lvwItem.SubItems(2).Text  'Account code'
                    strSql = "select * from dbo.FN_GET_FTP_FILEDATA(" & strInvoiceNo & ",'" & gstrUNITID & "')"
                    rs_hdr = New ClsResultSetDB
                    rs_hdr.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
                    Do While Not rs_hdr.EOFRecord
                        strSBUFolder = strFTPPath & "\Invoices\"
                        strEDIFolder = strFTPEDIPATH & "\Invoices\"

                        If (rs_hdr.GetValue("schedulecode").ToString).Trim.Length = 0 Then
                            strSBUFolder = strSBUFolder & "XXX\"
                            strEDIFolder = strEDIFolder & "XXX\"
                        Else
                            If UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "U5J" Or UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "UJW" Or UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "19J" Then
                                strSBUFolder = strSBUFolder & "U09\"
                                strEDIFolder = strEDIFolder & "U09\"
                            Else
                                strSBUFolder = strSBUFolder & Trim(rs_hdr.GetValue("schedulecode").ToString) & "\"
                                strEDIFolder = strEDIFolder & Trim(rs_hdr.GetValue("schedulecode").ToString) & "\"
                            End If
                        End If

                        If Not Directory.Exists(strSBUFolder) Then
                            Directory.CreateDirectory(strSBUFolder)
                        End If
                        If Not Directory.Exists(strEDIFolder) Then
                            Directory.CreateDirectory(strEDIFolder)
                        End If

                        minv = strInvoiceNo ' rs_hdr.GetValue("doc_no").ToString
                        strInvList = strInvList & "'" & minv & "',"
                        invnotopost = minv
                        mSuff = ""
                        invnotopost = CStr(Val(invnotopost))
                        '**************.inv file*********************
                        If UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "U5J" Or UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "UJW" Or UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "19J" Then
                            mfile = Trim(strSBUFolder) & sFolderName & "U09INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".inv"
                            minvEDIFile = Trim(strEDIFolder) & "U09INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".inv"
                        Else
                            mfile = Trim(strSBUFolder) & sFolderName & "" & Trim(rs_hdr.GetValue("schedulecode").ToString) & "INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".inv"
                            minvEDIFile = Trim(strEDIFolder) & Trim(rs_hdr.GetValue("schedulecode").ToString) & "INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".inv"
                        End If

                        '**************.ds file*********************
                        If UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "U5J" Or UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "UJW" Or UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "19J" Then
                            varDSFile = Trim(strSBUFolder) & sFolderName & "U09INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".ds"
                            mdsEDIFile = Trim(strEDIFolder) & "U09INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".ds"
                        Else
                            varDSFile = Trim(strSBUFolder) & sFolderName & "" & Trim(rs_hdr.GetValue("schedulecode").ToString) & "INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".ds"
                            mdsEDIFile = Trim(strEDIFolder) & Trim(rs_hdr.GetValue("schedulecode").ToString) & "INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".ds"
                        End If
                        '************************************************************

                        strHdr = Trim(rs_hdr.GetValue("account_code").ToString) & "|" & Trim(invnotopost) & "|" & rs_hdr.GetValue("INVOICE_DATE").ToString & "|" & Trim(invnotopost) & "|" & rs_hdr.GetValue("INVOICE_DATE").ToString & "|"

                        strSql = "select b.cust_item_code,b.Rate, b.sales_quantity,b.to_box From sales_dtl b "
                        strSql = strSql & " Where b.unit_code='" & gstrUNITID & "' and  b.doc_no=" & minv & " And b.suffix = '" & mSuff & "' "
                        strSql = strSql & " Order by b.cust_item_code"

                        rs_dtl = New ClsResultSetDB
                        rs_dtl.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

                        strDtl = ""
                        mValue = 0
                        Do While Not rs_dtl.EOFRecord
                            strDtl = strDtl & Trim(rs_dtl.GetValue("cust_item_code").ToString) & "|" & Trim(Str(rs_dtl.GetValue("sales_quantity").ToString)) & "|" & Trim(Str(rs_dtl.GetValue("to_box").ToString)) & "^"
                            mValue = mValue + (rs_dtl.GetValue("sales_quantity") * rs_dtl.GetValue("Rate"))
                            rs_dtl.MoveNext()
                        Loop

                        '***Query to Fetch Data DS File**********
                        strSql = "SELECT CUST_PART_CODE, DSNO,QUANTITYKNOCKEDOFF  FROM MKT_INVDSHISTORY H  INNER JOIN SALESCHALLAN_DTL SC ON SC.UNIT_CODE=H.UNIT_CODE AND SC.DOC_NO = H.DOC_NO  AND SC.LOCATION_CODE = H.LOCATION_CODE  AND SC.ACCOUNT_CODE = H.CUSTOMER_CODE INNER JOIN SALES_DTL SD ON SC.UNIT_CODE=SD.UNIT_CODE AND SC.DOC_NO = SD.DOC_NO AND SC.LOCATION_CODE = SD.LOCATION_CODE AND SD.ITEM_CODE = H.ITEM_CODE AND SC.UNIT_CODE='" & gstrUNITID & "' AND SC.DOC_NO = " & minv
                        objGetDSData = New ClsResultSetDB
                        objGetDSData.GetResult(strSql)
                        strDSFileData = ""
                        Do While Not objGetDSData.EOFRecord
                            strDSFileData = strDSFileData & objGetDSData.GetValue("CUST_PART_CODE").ToString & "|" & objGetDSData.GetValue("DSNO").ToString & "|" & objGetDSData.GetValue("QUANTITYKNOCKEDOFF").ToString & "^"
                            objGetDSData.MoveNext()
                        Loop
                        objGetDSData.ResultSetClose()

                        rsGetAmount = New ClsResultSetDB
                        stramt = " SELECT TOTAL_AMOUNT FROM SALESCHALLAN_DTL WHERE UNIT_CODE='" & gstrUNITID & "' AND DOC_NO=" & minv & ""
                        rsGetAmount.GetResult(stramt, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
                        If rsGetAmount.GetNoRows > 0 Then
                            stramt = rsGetAmount.GetValue("total_amount").ToString
                        End If

                        strHdr = strHdr & (stramt) & "|" & rs_hdr.GetValue("cust_ref").ToString & "##"
                        rsGetAmount.ResultSetClose()
                        rsGetAmount = Nothing

                        strRec = strHdr & strDtl
                        ' ***----  working with text file
                        nInvWrite = FreeFile()
                        FileOpen(nInvWrite, mfile, OpenMode.Output)
                        PrintLine(nInvWrite, strRec)
                        FileClose(nInvWrite)

                        '*********** Working in .DS file*************
                        intDSWrite = FreeFile()
                        FileOpen(intDSWrite, varDSFile, OpenMode.Output)
                        PrintLine(intDSWrite, strDSFileData)
                        FileClose(intDSWrite)


                        ' ***----  working with text file for EDI
                        'intinvEDIFile = FreeFile()
                        'FileOpen(intinvEDIFile, minvEDIFile, OpenMode.Output)
                        'PrintLine(intinvEDIFile, strRec)
                        'FileClose(intinvEDIFile)
                        objFSO.CopyFile(mfile, minvEDIFile)

                        '*********** Working in .DS file*********for EDI
                        'intdsEDIFile = FreeFile()
                        'FileOpen(intdsEDIFile, mdsEDIFile, OpenMode.Output)
                        'PrintLine(intdsEDIFile, strDSFileData)
                        'FileClose(intdsEDIFile)
                        objFSO.CopyFile(varDSFile, mdsEDIFile)

                        rs_dtl.ResultSetClose()

                        'intInvoice = intInvoice + 1
                        rs_hdr.MoveNext()
                    Loop
                    strInvList = Mid(strInvList, Len(Trim(strInvList)) + 1)
                    If rs_hdr.GetNoRows > 0 Then
                        'strLogMsg = Trim(Str(intInvoice)) & " Invoice(s) Found."
                        blnNewData = True
                    Else
                        strLogMsg = "No Invoices Found For MSSL."
                        'MsgBox("No Invoices Found For MSSL.", MsgBoxStyle.Information, ResolveResString(100))
                        blnNewData = False
                    End If
                    rs_hdr.ResultSetClose()
                    pgBar.PerformStep()
                    Application.DoEvents()
                Next
                If blnRecordExists = False Then
                    MsgBox("Please Select Invoice No. From Invoice List.", MsgBoxStyle.Information, ResolveResString(100))
                    Return False
                End If
            Else
                Dim strCustomer As String = String.Empty
                Dim dtInv As DataTable

                If optAllCust.Checked Then
                    strSql = "SELECT DOC_NO,ACCOUNT_CODE AS CUSTOMER_CODE FROM SALESCHALLAN_DTL A WHERE BILL_FLAG=1 AND CANCEL_FLAG=0 AND UNIT_CODE='" & gstrUNITID & "' " & _
                             " AND INVOICE_DATE BETWEEN '" & getDateForDB(dtFrom.Value) & "'  AND '" & getDateForDB(dtTo.Value) & "' " & _
                             " AND EXISTS(SELECT TOP 1 1 FROM CUSTOMER_MST WHERE UNIT_CODE =A.UNIT_CODE AND CUSTOMER_CODE = A.ACCOUNT_CODE AND SCHEDULECODE IS NOT NULL) " & _
                             " ORDER BY INVOICE_DATE "

                Else
                    If lvwCustomer.Items.Count > 0 Then
                        For Each lvwItem As ListViewItem In lvwCustomer.Items
                            If lvwItem.Checked Then
                                strCustomer += "'" + lvwItem.Text + "',"
                            End If
                        Next
                    Else
                        MsgBox("No Customer Record Found", MsgBoxStyle.Information, ResolveResString(100))
                        Return False
                    End If
                    If strCustomer.Trim.Length > 0 Then
                        strCustomer = "(" + strCustomer.Substring(0, strCustomer.Trim.Length - 1) + ")"
                    End If
                    strSql = "SELECT DOC_NO,ACCOUNT_CODE AS CUSTOMER_CODE FROM SALESCHALLAN_DTL WHERE BILL_FLAG=1 AND CANCEL_FLAG=0 AND UNIT_CODE='" & gstrUNITID & "' " & _
                             " AND INVOICE_DATE BETWEEN '" & getDateForDB(dtFrom.Value) & "'  AND '" & getDateForDB(dtTo.Value) & "' " & _
                             " AND ACCOUNT_CODE IN " + strCustomer
                End If

                dtInv = SqlConnectionclass.GetDataTable(strSql)

                lvwInvoices.Items.Clear()
                If dtInv.Rows.Count > 0 Then
                    pgBar.Maximum = dtInv.Rows.Count
                    For Each Row As DataRow In dtInv.Rows
                        strInvoiceNo = ""
                        strAccountCode = ""
                        strInvoiceNo = Convert.ToString(Row("DOC_NO"))
                        strAccountCode = Convert.ToString(Row("CUSTOMER_CODE"))
                        strSql = "select * from dbo.FN_GET_FTP_FILEDATA(" & strInvoiceNo & ",'" & gstrUNITID & "')"
                        rs_hdr = New ClsResultSetDB
                        rs_hdr.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
                        Do While Not rs_hdr.EOFRecord
                            strSBUFolder = strFTPPath & "\Invoices\"
                            strEDIFolder = strFTPEDIPATH & "\Invoices\"

                            If (rs_hdr.GetValue("schedulecode").ToString).Trim.Length = 0 Then
                                strSBUFolder = strSBUFolder & "XXX\"
                                strEDIFolder = strEDIFolder & "XXX\"
                            Else
                                If UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "U5J" Or UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "UJW" Or UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "19J" Then
                                    strSBUFolder = strSBUFolder & "U09\"
                                    strEDIFolder = strEDIFolder & "U09\"
                                Else
                                    strSBUFolder = strSBUFolder & Trim(rs_hdr.GetValue("schedulecode").ToString) & "\"
                                    strEDIFolder = strEDIFolder & Trim(rs_hdr.GetValue("schedulecode").ToString) & "\"
                                End If
                            End If

                            If Not Directory.Exists(strSBUFolder) Then
                                Directory.CreateDirectory(strSBUFolder)
                            End If
                            If Not Directory.Exists(strEDIFolder) Then
                                Directory.CreateDirectory(strEDIFolder)
                            End If

                            minv = strInvoiceNo
                            strInvList = strInvList & "'" & minv & "',"
                            invnotopost = minv
                            mSuff = ""
                            invnotopost = CStr(Val(invnotopost))
                            '**************.inv file*********************
                            If gstrUNITID = "SML" And (UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "U5J" Or UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "UJW" Or UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "19J") Then
                                mfile = Trim(strSBUFolder) & sFolderName & "U09INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".inv"
                                minvEDIFile = Trim(strEDIFolder) & "U09INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".inv"
                            Else
                                mfile = Trim(strSBUFolder) & sFolderName & "" & Trim(rs_hdr.GetValue("schedulecode").ToString) & "INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".inv"
                                minvEDIFile = Trim(strEDIFolder) & Trim(rs_hdr.GetValue("schedulecode").ToString) & "INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".inv"
                            End If

                            '**************.ds file*********************

                            If gstrUNITID = "SML" And (UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "U5J" Or UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "UJW" Or UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "19J") Then
                                varDSFile = Trim(strSBUFolder) & sFolderName & "U09INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".ds"
                                mdsEDIFile = Trim(strEDIFolder) & "U09INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".ds"
                            Else
                                varDSFile = Trim(strSBUFolder) & sFolderName & "" & Trim(rs_hdr.GetValue("schedulecode").ToString) & "INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".ds"
                                mdsEDIFile = Trim(strEDIFolder) & Trim(rs_hdr.GetValue("schedulecode").ToString) & "INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".ds"
                            End If

                            '************************************************************

                            strHdr = Trim(rs_hdr.GetValue("account_code").ToString) & "|" & Trim(invnotopost) & "|" & rs_hdr.GetValue("INVOICE_DATE").ToString & "|" & Trim(invnotopost) & "|" & rs_hdr.GetValue("INVOICE_DATE").ToString & "|"

                            strSql = "select b.cust_item_code,b.Rate, b.sales_quantity,b.to_box From sales_dtl b "
                            strSql = strSql & " Where b.unit_code='" & gstrUNITID & "' and  b.doc_no=" & minv & " And b.suffix = '" & mSuff & "' "
                            strSql = strSql & " Order by b.cust_item_code"

                            rs_dtl = New ClsResultSetDB
                            rs_dtl.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

                            strDtl = ""
                            mValue = 0
                            Do While Not rs_dtl.EOFRecord
                                strDtl = strDtl & Trim(rs_dtl.GetValue("cust_item_code").ToString) & "|" & Trim(Str(rs_dtl.GetValue("sales_quantity").ToString)) & "|" & Trim(Str(rs_dtl.GetValue("to_box").ToString)) & "^"
                                mValue = mValue + (rs_dtl.GetValue("sales_quantity") * rs_dtl.GetValue("Rate"))
                                rs_dtl.MoveNext()
                            Loop

                            '***Query to Fetch Data DS File**********
                            strSql = "SELECT CUST_PART_CODE, DSNO,QUANTITYKNOCKEDOFF  FROM MKT_INVDSHISTORY H  INNER JOIN SALESCHALLAN_DTL SC ON SC.UNIT_CODE=H.UNIT_CODE AND SC.DOC_NO = H.DOC_NO  AND SC.LOCATION_CODE = H.LOCATION_CODE  AND SC.ACCOUNT_CODE = H.CUSTOMER_CODE INNER JOIN SALES_DTL SD ON SC.UNIT_CODE=SD.UNIT_CODE AND SC.DOC_NO = SD.DOC_NO AND SC.LOCATION_CODE = SD.LOCATION_CODE AND SD.ITEM_CODE = H.ITEM_CODE AND SC.UNIT_CODE='" & gstrUNITID & "' AND SC.DOC_NO = " & minv
                            objGetDSData = New ClsResultSetDB
                            objGetDSData.GetResult(strSql)
                            strDSFileData = ""
                            Do While Not objGetDSData.EOFRecord
                                strDSFileData = strDSFileData & objGetDSData.GetValue("CUST_PART_CODE").ToString & "|" & objGetDSData.GetValue("DSNO").ToString & "|" & objGetDSData.GetValue("QUANTITYKNOCKEDOFF").ToString & "^"
                                objGetDSData.MoveNext()
                            Loop
                            objGetDSData.ResultSetClose()

                            rsGetAmount = New ClsResultSetDB
                            stramt = " SELECT TOTAL_AMOUNT FROM SALESCHALLAN_DTL WHERE UNIT_CODE='" & gstrUNITID & "' AND DOC_NO=" & minv & ""
                            rsGetAmount.GetResult(stramt, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
                            If rsGetAmount.GetNoRows > 0 Then
                                stramt = rsGetAmount.GetValue("total_amount").ToString
                            End If

                            strHdr = strHdr & (stramt) & "|" & rs_hdr.GetValue("cust_ref").ToString & "##"
                            rsGetAmount.ResultSetClose()
                            rsGetAmount = Nothing

                            strRec = strHdr & strDtl

                            ' ***----  working with text file
                            nInvWrite = FreeFile()
                            FileOpen(nInvWrite, mfile, OpenMode.Output)
                            PrintLine(nInvWrite, strRec)
                            FileClose(nInvWrite)

                            '*********** Working in .DS file*************
                            intDSWrite = FreeFile()
                            FileOpen(intDSWrite, varDSFile, OpenMode.Output)
                            PrintLine(intDSWrite, strDSFileData)
                            FileClose(intDSWrite)

                            ' ***----  working with text file for EDI
                            'intinvEDIFile = FreeFile()
                            'FileOpen(intinvEDIFile, minvEDIFile, OpenMode.Output)
                            'PrintLine(intinvEDIFile, strRec)
                            'FileClose(intinvEDIFile)
                            objFSO.CopyFile(mfile, minvEDIFile)

                            '*********** Working in .DS file*********for EDI
                            objFSO.CopyFile(varDSFile, mdsEDIFile)
                            'intdsEDIFile = FreeFile()
                            'FileOpen(intdsEDIFile, mdsEDIFile, OpenMode.Output)
                            'PrintLine(intdsEDIFile, strDSFileData)
                            'FileClose(intdsEDIFile)

                            rs_dtl.ResultSetClose()
                            rs_dtl = Nothing

                            rs_hdr.MoveNext()
                        Loop
                        strInvList = Mid(strInvList, Len(Trim(strInvList)) + 1)
                        If rs_hdr.GetNoRows > 0 Then
                            'strLogMsg = Trim(Str(intInvoice)) & " Invoice(s) Found."
                            blnNewData = True
                        Else
                            strLogMsg = "No Invoices Found For MSSL."
                            'MsgBox("No Invoices Found For MSSL.", MsgBoxStyle.Information, ResolveResString(100))
                            blnNewData = False
                        End If
                        rs_hdr.ResultSetClose()
                        pgBar.PerformStep()
                        Application.DoEvents()
                    Next

                Else
                    MsgBox("No Invoices Found", MsgBoxStyle.Information, ResolveResString(100))
                    Return False
                End If

            End If
            FileClose(nInFile)
            CheckInvoices = True
            MsgBox("File(s) Reprocessed Successfully.", MsgBoxStyle.Information, ResolveResString(100))
            Exit Function
        Catch ex As Exception
            strLogMsg = "Error " & Err.Description & " " & CStr(Now)
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
            CheckInvoices = False
        Finally
            objFSO = Nothing
            FileClose(nInFile)
            rs_hdr = Nothing
            objGetDSData = Nothing
            objGetDSData = Nothing
            rs_dtl = Nothing
        End Try
    End Function

    Private Sub SearchText(ByRef lvwListView As System.Windows.Forms.ListView, ByRef txtSearchBox As System.Windows.Forms.TextBox, ByRef optFistOption As System.Windows.Forms.RadioButton, ByRef optSecOption As System.Windows.Forms.RadioButton)

        Dim Intcounter As Short
        Try
            With lvwListView
                If optFistOption.Checked = True Then
                    For Intcounter = 0 To .Items.Count - 1
                        If .Items.Item(Intcounter).Font.Bold = True Then
                            .Items.Item(Intcounter).Font = VB6.FontChangeBold(.Items.Item(Intcounter).Font, False)
                            .Refresh()
                        End If
                        If .Items.Item(Intcounter).SubItems.Item(0).Font.Bold = True Then
                            .Items.Item(Intcounter).SubItems.Item(0).Font = VB6.FontChangeBold(.Items.Item(Intcounter).SubItems.Item(1).Font, False)
                            .Refresh()
                        End If
                    Next
                    If Len(txtSearchBox.Text) = 0 Then Exit Sub
                    For Intcounter = 0 To .Items.Count - 1
                        If Trim(UCase(Mid(.Items.Item(Intcounter).Text, 1, Len(txtSearchBox.Text)))) = Trim(UCase(txtSearchBox.Text)) Then
                            .Items.Item(Intcounter).Font = VB6.FontChangeBold(.Items.Item(Intcounter).Font, True)
                            Call .Items.Item(Intcounter).EnsureVisible()
                            .Refresh()
                            Exit For
                        End If
                    Next
                ElseIf optSecOption.Checked Then
                    For Intcounter = 0 To .Items.Count - 1
                        If .Items.Item(Intcounter).Font.Bold = True Then
                            .Items.Item(Intcounter).Font = VB6.FontChangeBold(.Items.Item(Intcounter).Font, False)
                            .Refresh()
                        End If
                        If .Items.Item(Intcounter).SubItems.Item(0).Font.Bold = True Then
                            .Items.Item(Intcounter).SubItems.Item(0).Font = VB6.FontChangeBold(.Items.Item(Intcounter).SubItems.Item(1).Font, False)
                            .Refresh()
                        End If
                    Next
                    If Len(txtSearchBox.Text) = 0 Then Exit Sub
                    For Intcounter = 0 To .Items.Count - 1
                        If Trim(UCase(Mid(.Items.Item(Intcounter).SubItems.Item(1).Text, 1, Len(txtSearchBox.Text)))) = Trim(UCase(txtSearchBox.Text)) Then
                            .Items.Item(Intcounter).SubItems.Item(0).Font = VB6.FontChangeBold(.Items.Item(Intcounter).SubItems.Item(0).Font, True)
                            Call .Items.Item(Intcounter).EnsureVisible()
                            .Refresh()
                            Exit For
                        End If
                    Next
                End If
            End With
            Exit Sub
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End Try
    End Sub

#End Region

#Region "Control's Events"

    Private Sub optAllCust_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAllCust.CheckedChanged
        Try
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
            If sender.Checked Then
                With lvwCustomer
                    .Items.Clear()
                    .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    TxtSearchCust.Text = ""
                    grpSearchCust.Enabled = False
                End With
                optAllInv.Checked = True
            Else
                lvwCustomer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                TxtSearchCust.Text = ""
                grpSearchCust.Enabled = True
                FillCustomer()
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub optAllInv_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAllInv.CheckedChanged
        Try
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
            If sender.Checked Then
                With lvwInvoices
                    .Items.Clear()
                    .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtSearchInv.Text = ""
                    grpSearchInv.Enabled = False
                End With
            Else
                lvwInvoices.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                FillInvoices()
                txtSearchInv.Text = ""
                grpSearchInv.Enabled = True
                'If lvwInvoices.Items.Count = 0 Then optAllInv.Checked = True
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub lvwCustomer_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles lvwCustomer.ItemChecked
        e.Item.Selected = True
    End Sub

    Private Sub lvwInvoices_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles lvwInvoices.ItemChecked
        e.Item.Selected = True
    End Sub

    Private Sub cmdGrp_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtnEditGrp.ButtonClickEventArgs) Handles cmdGrp.ButtonClick
        Try
            Select Case e.Button
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                    ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
                    cmdGrp.Revert()
                    cmdGrp.Caption(0) = "Create"
                    pgBar.Maximum = 0
                    CheckInvoices()
                    ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                    optAllCust.Checked = True
                    optAllInv.Checked = True
                    dtFrom.Value = GetServerDate()
                    dtTo.Value = GetServerDate()
                    lvwCustomer.Items.Clear()
                    lvwInvoices.Items.Clear()
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                    Me.Close()
            End Select
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub TxtSearchCust_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtSearchCust.TextChanged
        Try
            SearchText(lvwCustomer, TxtSearchCust, OptSearchCustCode, OptSearchCustName)
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End Try
    End Sub

    Private Sub txtSearchInv_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSearchInv.TextChanged
        Try
            SearchText(lvwInvoices, txtSearchInv, optSearchInvNo, optSearchInvDate)
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End Try
    End Sub

#End Region

   
End Class