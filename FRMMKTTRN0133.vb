Option Strict Off
Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
'----------------------------------------------------
'Copyright (c)  -  MIND
'Name of module -  frmMKTTRN0133.vb
'Created By     -  
'Created Date   -  
'Description    -
'Revised date   -
'============================================================================================
'Revised By         :   Priti Sharma
'Revised On         :   28 Jan 2026
'Reason             :   Added Delivery Note No
'issue id           :    
'============================================================================================
Friend Class frmMKTTRN0133
    Inherits System.Windows.Forms.Form
    Dim strUploadedFileName As String
    Dim mintIndex As Short
    Dim mflag As Short
    Dim strStockLocation As String
    Private Enum enmUploadItem
        TempInvoiceNo
        SaleOrder
        AmendmentNo
        ItemCode
        ItemDesc
        Cust_DrgwNo
        Cust_DrgDesc
        CurBal
        Qty
        Rate
        FromBox
        ToBox
        Currency_Code
        Payment_Terms
        CUST_MTRL
        TOOL_COST
        HSNCode
        CGSTTXRT_TYPE
        SGSTTXRT_TYPE
        IGSTTXRT_TYPE
        UTGSTTXRT_TYPE
        CUST_NAME
        PACKING
    End Enum
    Private Enum enmInvoiceItem
        status = 0
        SaleOrder
        AmendmentNo
        ItemCode
        ItemDesc
        Cust_DrgwNo
        Cust_DrgDesc
        HSNCode
        CurBal
    End Enum
    Private Enum enmInvoiceItemSearch
        SaleOrder
        ItemCode
        ItemDesc
        Cust_DrgwNo
        Cust_DrgDesc
    End Enum
  
    Private Sub RefreshForm()
        Try
            txtcustomercode.Text = ""
            lblcustomername.Text = ""
            txtfilelocation.Text = ""
            TxtSearch.Text = ""
            CmbSearchBy.SelectedIndex = -1
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub chkall_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkall.CheckStateChanged
        Try
            Dim intcounter As Long
            Dim varstatus As Object
            With dgvItemDetail
                For i As Integer = 0 To dgvItemDetail.Rows.Count - 1
                    If chkall.CheckState = 1 Then
                        dgvItemDetail.Rows(i).Cells(enmInvoiceItem.status).Value = True
                    Else
                        dgvItemDetail.Rows(i).Cells(enmInvoiceItem.status).Value = False
                    End If
                Next
            End With
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub CmbSearchBy_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbSearchBy.TextChanged
        Try
            If TxtSearch.Enabled Then
                TxtSearch.Focus()
            End If
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub CmbSearchBy_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbSearchBy.SelectedIndexChanged
        Try
            If TxtSearch.Enabled Then
                TxtSearch.Focus()
            End If
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub AddTransPortTypeToCombo()
        Try
            VB6.SetItemString(CmbTransType, 0, "R - Road") 'Road
            VB6.SetItemString(CmbTransType, 1, "L - Rail") 'Rail
            VB6.SetItemString(CmbTransType, 2, "S - Sea") 'Sea
            VB6.SetItemString(CmbTransType, 3, "A - Air") 'Air
            VB6.SetItemString(CmbTransType, 4, "H - Hand") 'Hand
            VB6.SetItemString(CmbTransType, 5, "C - Courier") 'Courier
            CmbTransType.SelectedIndex = 0
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub cmdCustHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdcusthelp.Click
        Try
            Dim strCustHelp() As String
            Dim strsql As String
            If Len(Me.txtcustomerhelp.Text) = 0 Then
                strCustHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct customer_Code, cust_name as customer_name from customer_mst b(nolock)  " & _
          "where b.unit_code='" & gstrUNITID & "'  and isnull(InvoiceUploading,0)=1  and ((isnull(b.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= b.deactive_date)) ", "Customer Codes List", 1)
            Else
                strCustHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select a.customer_Code, a.cust_name  as customer_name from Customer_Mst a where a.customer_code='" & Trim(txtcustomerhelp.Text) & "' and isnull(InvoiceUploading,0)=1 and a.Unit_code = '" & gstrUNITID & "' and ((isnull(a.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= a.deactive_date))", "Customer Codes List", 1)
            End If
            If UBound(strCustHelp) <> -1 Then
                If strCustHelp(0) <> "0" Then
                    txtcustomerhelp.Text = Trim(strCustHelp(0))
                    Call txtcustomerhelp_Validating(txtcustomerhelp, New System.ComponentModel.CancelEventArgs(False))
                    lblcustname.Text = strCustHelp(1)


                Else
                    txtcustomerhelp.Text = ""
                    MsgBox("No Record Available", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                End If
            Else
                txtcustomerhelp.Text = ""
                If txtcustomerhelp.Enabled Then txtcustomerhelp.Focus()
                Exit Sub
            End If
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    
    Private Sub cmdfilelocation_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdfilelocation.Click
        Try
            Dim strUploadingpath As String
            CommanDLogOpen.FileName = ""
            CommanDLogOpen.InitialDirectory = ""
            CommanDLogOpen.FileName = ""
            CommanDLogOpen.InitialDirectory = gstrLocalCDrive
            CommanDLogOpen.Filter = "Microsoft Excel File (*.xls)|*.xls;*.xlsx;*.CSV"
            CommanDLogOpen.ShowDialog()
            Me.txtfilelocation.Text = CommanDLogOpen.FileName
            'strUploadedFileName = CommanDLogOpen.SafeFileName()
            If txtfilelocation.Text.Trim.Length = 0 Then
                Exit Sub
            End If

            strUploadedFileName = Mid(CommanDLogOpen.FileName, CommanDLogOpen.FileName.LastIndexOf("\") + 2, CommanDLogOpen.FileName.Length - 1)
            strUploadingpath = gstrLocalCDrive & "DSUpload\uploaded_files"
            If txtfilelocation.Text <> "" Then
                If strUploadingpath = Mid(txtfilelocation.Text, 1, Len(txtfilelocation.Text) - 41) Then
                    MsgBox("You cannot select file from folder: " & gstrLocalCDrive & "Forecast\uploaded_files.")
                    txtfilelocation.Text = ""
                    SetItemUploadGridsHeader()
                End If
            End If
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub cmdgenerateformat_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdgenerateformat.Click
        Try
            Dim OBJexlformat As Excel.Application
            Dim objfso As Scripting.FileSystemObject
            Dim OBj_wB As Excel.Workbook
            Dim objfolder As Scripting.Folder
            Dim objfile As Scripting.File
            Dim stroldfilename As String
            Dim strnewfilename As String
            objfso = New Scripting.FileSystemObject
            OBJexlformat = New Excel.Application
            If Len(Trim(txtcustomerhelp.Text)) = 0 Then
                MsgBox("Please Select Customer First", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            ElseIf Len(Trim(txtRefNo.Text)) = 0 Then
                MsgBox("Please Select Sale Order First", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
            If dgvItemDetail.Rows.Count > 0 Then
                Dim intCount As Integer
                For i As Integer = 0 To dgvItemDetail.Rows.Count - 1
                    If dgvItemDetail.Rows(i).Cells(enmInvoiceItem.status).Value = True Then
                        intCount = intCount + 1
                        Exit For
                    End If
                Next
                If intCount = 0 Then
                    MsgBox("Please Select Item Code for generating sheet", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If
            End If

            If dgvItemDetail.Rows.Count > 0 Then
                If objfso.FolderExists(gstrLocalCDrive & "InvoiceUpload") = True Then
                    objfolder = objfso.GetFolder(gstrLocalCDrive & "InvoiceUpload")
                    For Each objfile In objfolder.Files
                        stroldfilename = objfile.Name
                        If UCase(Mid(stroldfilename, 10, 8)) = UCase(txtcustomerhelp.Text) Then
                            If objfso.FolderExists(gstrLocalCDrive & "InvoiceUpload\Backup") = True Then
                                If stroldfilename <> "" Then
                                    objfso.MoveFile(gstrLocalCDrive & "InvoiceUpload\" & stroldfilename, gstrLocalCDrive & "InvoiceUpload\backup\")
                                End If
                            Else
                                objfso.CreateFolder(gstrLocalCDrive & "InvoiceUpload\backup")
                                If stroldfilename <> "" Then
                                    objfso.MoveFile(gstrLocalCDrive & "InvoiceUpload\" & stroldfilename, gstrLocalCDrive & "InvoiceUpload\backup\")
                                End If
                            End If
                        End If
                    Next objfile
                    OBj_wB = OBJexlformat.Workbooks.Add
                    strnewfilename = getfilename(Trim(txtcustomerhelp.Text))
                    OBj_wB.SaveAs(gstrLocalCDrive & "InvoiceUpload\" & Replace(strnewfilename, ":", ""))
                    OBJexlformat.Workbooks.Close()
                    OBj_wB = Nothing
                    OBJexlformat.Workbooks.Open(gstrLocalCDrive & "InvoiceUpload\" & Replace(strnewfilename, ":", ""))

                    OBJexlformat.Cells._Default(1, 1) = "Customer Code"
                    OBJexlformat.Cells._Default(1, 1).Font.Bold = True
                    OBJexlformat.Cells._Default(1, 2) = "Customer Name"
                    OBJexlformat.Cells._Default(1, 2).Font.Bold = True
                    OBJexlformat.Cells._Default(1, 3) = "Sale Order"
                    OBJexlformat.Cells._Default(1, 3).Font.Bold = True
                    OBJexlformat.Cells._Default(1, 4) = "Amendment No"
                    OBJexlformat.Cells._Default(1, 4).Font.Bold = True
                    OBJexlformat.Cells._Default(1, 5) = "Item Code"
                    OBJexlformat.Cells._Default(1, 5).Font.Bold = True
                    OBJexlformat.Cells._Default(1, 6) = "Item Desc"
                    OBJexlformat.Cells._Default(1, 6).Font.Bold = True
                    OBJexlformat.Cells._Default(1, 7) = "Cust Drgno"
                    OBJexlformat.Cells._Default(1, 7).Font.Bold = True
                    OBJexlformat.Cells._Default(1, 8) = "Cust DrgDesc"
                    OBJexlformat.Cells._Default(1, 8).Font.Bold = True
                    OBJexlformat.Cells._Default(1, 9) = "Qty"
                    OBJexlformat.Cells._Default(1, 9).Font.Bold = True
                    OBJexlformat.Cells._Default(1, 10) = "From Box"
                    OBJexlformat.Cells._Default(1, 10).Font.Bold = True
                    OBJexlformat.Cells._Default(1, 11) = "To Box"
                    OBJexlformat.Cells._Default(1, 11).Font.Bold = True
                    'OBJexlformat.Cells(3).numberformat = "@"
                    OBJexlformat.Cells(3).columns.numberformat = "@"
                    OBJexlformat.Cells(3).entirecolumn.numberformat = "@"

                    OBJexlformat.Cells(4).columns.numberformat = "@"
                    OBJexlformat.Cells(4).entirecolumn.numberformat = "@"

                    OBJexlformat.Cells(5).columns.numberformat = "@"
                    OBJexlformat.Cells(5).entirecolumn.numberformat = "@"

                    OBJexlformat.Cells(7).columns.numberformat = "@"
                    OBJexlformat.Cells(7).entirecolumn.numberformat = "@"

                    Call fillitems(OBJexlformat)
                    OBJexlformat.ActiveWorkbook.Save()
                    OBJexlformat.Quit()

                    MsgBox("Excel Sheet Generated Successfully " & vbCrLf & "location--> " & " " & gstrLocalCDrive & "DSUpload\" & Replace(strnewfilename, ":", ""), MsgBoxStyle.Information, ResolveResString(100))
                    clear_fields()
                    chkall.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkall.Enabled = False
                    OBJexlformat = Nothing
                Else
                    objfso.CreateFolder(gstrLocalCDrive & "InvoiceUpload")
                    OBj_wB = OBJexlformat.Workbooks.Add
                    strnewfilename = getfilename(Trim(txtcustomerhelp.Text))
                    OBj_wB.SaveAs(gstrLocalCDrive & "InvoiceUpload\" & Replace(strnewfilename, ":", ""))
                    OBJexlformat.Workbooks.Close()
                    OBj_wB = Nothing
                    OBJexlformat.Workbooks.Open(gstrLocalCDrive & "InvoiceUpload\" & Replace(strnewfilename, ":", ""))
                    OBJexlformat.Cells._Default(1, 1) = "Customer Code"
                    OBJexlformat.Cells._Default(1, 1).Font.Bold = True
                    OBJexlformat.Cells._Default(1, 2) = "Customer Name"
                    OBJexlformat.Cells._Default(1, 2).Font.Bold = True
                    OBJexlformat.Cells._Default(1, 3) = "Sale Order"
                    OBJexlformat.Cells._Default(1, 3).Font.Bold = True
                    OBJexlformat.Cells._Default(1, 4) = "Amendment No"
                    OBJexlformat.Cells._Default(1, 4).Font.Bold = True
                    OBJexlformat.Cells._Default(1, 5) = "Item Code"
                    OBJexlformat.Cells._Default(1, 5).Font.Bold = True
                    OBJexlformat.Cells._Default(1, 6) = "Item Desc"
                    OBJexlformat.Cells._Default(1, 6).Font.Bold = True
                    OBJexlformat.Cells._Default(1, 7) = "Cust Drgno"
                    OBJexlformat.Cells._Default(1, 7).Font.Bold = True
                    OBJexlformat.Cells._Default(1, 8) = "Cust DrgDesc"
                    OBJexlformat.Cells._Default(1, 8).Font.Bold = True
                    OBJexlformat.Cells._Default(1, 9) = "Qty"
                    OBJexlformat.Cells._Default(1, 9).Font.Bold = True
                    OBJexlformat.Cells._Default(1, 10) = "From Box"
                    OBJexlformat.Cells._Default(1, 10).Font.Bold = True
                    OBJexlformat.Cells._Default(1, 11) = "To Box"
                    OBJexlformat.Cells._Default(1, 11).Font.Bold = True
                    'OBJexlformat.Cells(3).numberformat = "@"
                    OBJexlformat.Cells(9).columns.numberformat = "@"
                    OBJexlformat.Cells(9).entirecolumn.numberformat = "@"

                    OBJexlformat.Cells(10).columns.numberformat = "@"
                    OBJexlformat.Cells(10).entirecolumn.numberformat = "@"

                    OBJexlformat.Cells(11).columns.numberformat = "@"
                    OBJexlformat.Cells(11).entirecolumn.numberformat = "@"

                    Call fillitems(OBJexlformat)
                    OBJexlformat.ActiveWorkbook.Save()
                    OBJexlformat.Quit()

                    MsgBox("Excel Sheet Generated Successfully " & vbCrLf & "location--> " & " " & gstrLocalCDrive & "DSUpload\" & Replace(strnewfilename, ":", ""), MsgBoxStyle.Information, ResolveResString(100))
                    clear_fields()
                    chkall.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkall.Enabled = False
                    OBJexlformat = Nothing
                End If
            Else
                MsgBox("No Item Is Selected", MsgBoxStyle.Information, ResolveResString(100))
            End If
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub cmdclose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdclose.Click
        Try
            Me.Close()
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub cmdupload_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdupload.Click
        Try
            If txtcustomercode.Text <> "" Then
                If txtfilelocation.Text <> "" Then
                    'dgvUpload.Rows.Count = 0
                    cmdsave.Enabled = True

                    Call upload_InvoiceData()
                Else
                    MsgBox("Please Select File Location To Upload DS", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If
            Else
                MsgBox("Please Select Customer First", MsgBoxStyle.Information, ResolveResString(100))
                txtcustomercode.Focus()
                Exit Sub
            End If
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function StockLocationSalesConf(ByRef pstrInvType As String, ByRef pstrInvSubtype As String, ByRef pstrFeild As String) As String
        Dim rsSalesConf As ClsResultSetDB
        Dim StockLocation As String
        Try
            rsSalesConf = New ClsResultSetDB
            Select Case pstrFeild
                Case "DESCRIPTION"
                    rsSalesConf.GetResult("Select Stock_Location from SaleConf (nolock) WHERE UNIT_CODE='" + gstrUNITID + "' AND  Description ='" & Trim(pstrInvType) & "' and Sub_type_Description ='" & Trim(pstrInvSubtype) & "' AND Location_Code='" & gstrUNITID & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                Case "TYPE"
                    rsSalesConf.GetResult("Select Stock_Location from SaleConf (nolock) WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_type ='" & Trim(pstrInvType) & "' and Sub_type ='" & Trim(pstrInvSubtype) & "' AND Location_Code='" & gstrUNITID & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
            End Select
            If rsSalesConf.GetNoRows > 0 Then
                StockLocation = rsSalesConf.GetValue("Stock_Location")
            End If
            rsSalesConf.ResultSetClose()
            StockLocationSalesConf = StockLocation
            Exit Function
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub cmdviewitems_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdviewitems.Click
        Try
            Dim objgetitems As New ClsResultSetDB
            Dim strSQL As String
            Dim sqlCmd As New SqlCommand
            Dim strDate As String
            If Len(Trim(txtcustomerhelp.Text)) = 0 Then
                MsgBox("Please Select Customer Code First", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            ElseIf Len(Trim(txtRefNo.Text)) = 0 Then
                MsgBox("Please Select Reference No First", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
            SetItemGridsHeader()
            Dim i As Integer = 0
            strDate = VB6.Format(GetServerDate, gstrDateFormat)
            strSQL = makeSelectSql(txtcustomerhelp.Text, Trim(CUSTREFLIST), "", "", "", strStockLocation, strDate, "'F','S'", "")
            Using dtItem As DataTable = SqlConnectionclass.GetDataTable(strSQL)
                If dtItem.Rows.Count > 0 Then
                    dgvItemDetail.Rows.Clear()
                    dgvItemDetail.Rows.Add(dtItem.Rows.Count)
                    For Each dr As DataRow In dtItem.Rows

                        dgvItemDetail.Rows(i).Cells(enmInvoiceItem.status).Value = False
                        dgvItemDetail.Rows(i).Cells(enmInvoiceItem.SaleOrder).Value = dr("Cust_ref")
                        dgvItemDetail.Rows(i).Cells(enmInvoiceItem.AmendmentNo).Value = dr("AMENDMENT_NO")
                        dgvItemDetail.Rows(i).Cells(enmInvoiceItem.ItemCode).Value = dr("Item_Code")
                        dgvItemDetail.Rows(i).Cells(enmInvoiceItem.ItemDesc).Value = dr("Description")
                        dgvItemDetail.Rows(i).Cells(enmInvoiceItem.Cust_DrgwNo).Value = dr("Cust_DrgNo")
                        dgvItemDetail.Rows(i).Cells(enmInvoiceItem.Cust_DrgDesc).Value = dr("Cust_Drg_Desc")
                        dgvItemDetail.Rows(i).Cells(enmInvoiceItem.HSNCode).Value = dr("HSN_SAC_CODE")
                        dgvItemDetail.Rows(i).Cells(enmInvoiceItem.CurBal).Value = dr("Cur_Bal")
                        i += 1
                    Next
                    chkall.Enabled = True
                    cmdgenerateformat.Enabled = True
                Else
                    MsgBox("No Item found for selected sale order", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If
            End Using

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Function makeSelectSql(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, ByRef effectyrmon As String, ByRef Validyrmon As String, ByRef pstrstockLocation As String, ByRef strDate As String, ByRef pstrItemin As String, Optional ByRef pstrCondition As String = "", Optional ByRef pstrConsCode As String = "") As String
        '=======================================================================================
        'Revised By     : Ashutosh Verma
        'Revised On     : 22-01-2007 ,Issue ID:19352
        'Revised Reason : Consider Current month schedule while Invoicing.
        '=======================================================================================
        Dim strSelectSql As String
        strDate = getDateForDB(strDate)
        If gblnGSTUnit = True Then
            strSelectSql = "Select distinct b.Item_Code,d.Description,c.Cust_DrgNo,c.Cust_Drg_Desc,d.HSN_SAC_CODE,C.Cust_ref, c.EXTERNAL_SALESORDER_NO , C.AMENDMENT_NO , SHOP_CODE "
        Else
            strSelectSql = "Select distinct b.Item_Code,d.Description,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code ,C.Cust_ref, c.EXTERNAL_SALESORDER_NO , C.AMENDMENT_NO , SHOP_CODE "
        End If

        strSelectSql = strSelectSql & " , CUR_BAL= (	SELECT CUR_BAL FROM ITEMBAL_MST I WHERE B.UNIT_CODE = I.UNIT_CODE 	AND B.ITEM_CODE = I.ITEM_CODE AND I.UNIT_CODE='" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " AND I.LOCATION_CODE='" & pstrstockLocation & "')"
        strSelectSql = strSelectSql & " from Cust_Ord_hdr a,MonthlyMktSchedule b,Cust_ord_dtl c,Item_Mst d , custitem_mst cm "

        strSelectSql = strSelectSql & " where a.unit_code = b.unit_code and b.unit_code = c.unit_code and B.unit_code = d.unit_code "
        strSelectSql = strSelectSql & " and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code And c.Active_Flag ='A' And c.Authorized_flag = 1"
        strSelectSql = strSelectSql & " and cm.unit_code=b.unit_code and cm.account_code =b.account_code and cm.item_Code=b.item_code and cm.cust_drgno=b.Cust_drgNo "

        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code = d.Item_code and a.Account_Code='" & Trim(pstrCustno) & "' and a.unit_code='" & gstrUNITID & "'"
        If Len(Trim(pstrConsCode)) > 0 Then
            strSelectSql = strSelectSql & " and a.consignee_code='" & Trim(pstrConsCode) & "' and a.Consignee_code=b.Consignee_code "
        End If
        strSelectSql = strSelectSql & " and b.ITem_code = c.Item_code and a.Cust_Ref in ('" & Trim(pstrRefNo) & "') and b.status = 1 and b.Schedule_flag =1 and b.Year_Month =  '2024'"
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code = b.unit_code and a.unit_code='" & gstrUNITID & "' and a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        '101254587
        'strSelectSql = strSelectSql & CreateSubQueryForGlobalToolItemCheck(pstrItemin, pstrCustno)
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in(" & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " UNION "
        If gblnGSTUnit = True Then
            strSelectSql = strSelectSql & "Select distinct b.Item_Code,d.Description,c.Cust_DrgNo,c.Cust_Drg_Desc,d.HSN_SAC_CODE,C.Cust_ref, c.EXTERNAL_SALESORDER_NO , C.AMENDMENT_NO , SHOP_CODE "
        Else
            strSelectSql = strSelectSql & "Select distinct b.Item_Code,d.Description,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code,C.Cust_ref , c.EXTERNAL_SALESORDER_NO , C.AMENDMENT_NO , SHOP_CODE "
        End If

        strSelectSql = strSelectSql & " , CUR_BAL=(	SELECT CUR_BAL FROM ITEMBAL_MST I WHERE B.UNIT_CODE = I.UNIT_CODE 	AND B.ITEM_CODE = I.ITEM_CODE AND I.UNIT_CODE='" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " AND I.LOCATION_CODE='" & pstrstockLocation & "')"
        strSelectSql = strSelectSql & " from Cust_Ord_hdr a,Dailymktschedule   b,Cust_ord_dtl c,Item_Mst d , custitem_mst cm  "

        strSelectSql = strSelectSql & " where a.unit_code=b.unit_code and b.unit_code=c.unit_code and B.unit_code  = d.unit_code and "
        strSelectSql = strSelectSql & " a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code"
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code =d.ITem_code "


        strSelectSql = strSelectSql & " and cm.unit_code=b.unit_code and cm.account_code =b.account_code and cm.item_Code=b.item_code and cm.cust_drgno=B.Cust_drgNo and "
        strSelectSql = strSelectSql & " b.status = 1 and b.Schedule_Flag = 1 And c.Active_Flag ='A' And c.Authorized_flag = 1 and a.Account_Code='" & Trim(pstrCustno) & "' "
        strSelectSql = strSelectSql & " and a.unit_code='" & gstrUNITID & "' and a.Valid_date >='" & getDateForDB(GetServerDate()) & "' and effect_Date <= '" & getDateForDB(GetServerDate()) & "'"
        ''Changes done by Ashutosh on 16 Apr 2007, issue id:19731
        If Len(Trim(pstrConsCode)) > 0 Then
            strSelectSql = strSelectSql & " and a.consignee_code='" & Trim(pstrConsCode) & "' and a.Consignee_code=b.Consignee_code "
        End If
        ''Changes for issue id:19731 end here.
        ''*** Changes done By ashutosh on 22-01-2007, issue Id: 19352, Consider Current month schedule.
        strSelectSql = strSelectSql & " and b.ITem_code = c.Item_code and a.Cust_Ref in('" & Trim(pstrRefNo) & "') and b.trans_date <= '" & strDate & "' And datepart(mm,b.trans_date) = '" & Month(CDate(strDate)) & "' And DatePart(yyyy, b.Trans_Date) = '" & Year(CDate(strDate)) & "'"
        ''*** Changes for Issue Id:19352 end here.
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code = b.unit_code and a.unit_code ='" & gstrUNITID & "' and a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        '101254587
        'strSelectSql = strSelectSql & CreateSubQueryForGlobalToolItemCheck(pstrItemin, pstrCustno)

        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in( " & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " ORDER BY CUST_REF,ITEM_CODE "
        makeSelectSql = strSelectSql
    End Function
    Private Sub frmMKTTRN0063_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Try
            mdifrmMain.CheckFormName = mintIndex
            frmModules.NodeFontBold(Tag) = True
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub frmMKTTRN0063_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        Try
            frmModules.NodeFontBold(Me.Tag) = False
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub frmMKTTRN0063_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Try
            mintIndex = mdifrmMain.AddFormNameToWindowList(ctlDSWiseScheduleStatus.Tag)
            FitToClient(Me, framain, ctlDSWiseScheduleStatus, frabuttons)
            Call FillLabelFromResFile(Me)
            FillItemSearchCategory()
            cmdcustomercode.Image = My.Resources.ico111.ToBitmap
            ResetData()
            SSTab1.SelectedIndex = 0
            SelectInvoiceTypeFromSaleConf()
            SelectInvoiceSubTypeFromSaleConf("NORMAL INVOICE")
            CmbInvType.Text = "NORMAL INVOICE"
            CmbInvSubType.Text = "FINISHED GOODS"
            strStockLocation = StockLocationSalesConf("INV", "F", "TYPE")
            AddTransPortTypeToCombo()
            dtpDateDesc.Enabled = False
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub ResetData()
        SetItemGridsHeader()
        SetItemUploadGridsHeader()
        txtfilelocation.Enabled = False
        chkall.Enabled = False
        cmdgenerateformat.Enabled = False
        cmdsave.Enabled = False
        txtcustomercode.Text = ""
        lblcustomername.Text = ""
        txtcustomerhelp.Text = ""
        lblcustomername.Text = ""
        txtRefNo.Text = ""
        txtCarrServices.Text = ""
        txtVehNo.Text = ""
        TxtLRNO.Text = ""
        txtfilelocation.Text = ""
        txtDeliveryNoteNo.Text = ""
    End Sub
    Private Sub frmMKTTRN0063_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            frmModules.NodeFontBold(Me.Tag) = False
            mdifrmMain.RemoveFormNameFromWindowList = mintIndex
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub cmdcustomercode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdcustomercode.Click
        Try
            Dim strCustHelp() As String
            If Len(Me.txtcustomerhelp.Text) = 0 Then
                strCustHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct customer_Code, cust_name as customer_name from customer_mst b(nolock)  " & _
          "where b.unit_code='" & gstrUNITID & "'  and isnull(InvoiceUploading,0)=1  and ((isnull(b.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= b.deactive_date)) ", "Customer Codes List", 1)
            Else
                strCustHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select a.customer_Code, a.cust_name  as customer_name from Customer_Mst a where a.customer_code='" & Trim(txtcustomercode.Text) & "' and isnull(InvoiceUploading,0)=1 and a.Unit_code = '" & gstrUNITID & "' and ((isnull(a.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= a.deactive_date))", "Customer Codes List", 1)
            End If
            If UBound(strCustHelp) <> -1 Then
                If strCustHelp(0) <> "0" Then
                    txtcustomercode.Text = Trim(strCustHelp(0))
                    Call TxtCustomerCode_Validating(txtcustomercode, New System.ComponentModel.CancelEventArgs(False))
                    lblcustomername.Text = strCustHelp(1)
                Else
                    RefreshForm()
                    txtcustomercode.Text = ""
                    MsgBox("No Record Available", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                End If
            Else
                RefreshForm()
                txtcustomercode.Text = ""
                If txtcustomercode.Enabled Then txtcustomercode.Focus()
                Exit Sub
            End If
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub SSTab1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SSTab1.SelectedIndexChanged
        Try
            If SSTab1.SelectedIndex = 0 Then
                cmdsave.Enabled = False
            Else
                cmdsave.Enabled = True
            End If
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub txtCustomerCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtcustomercode.TextChanged
        Try
            lblcustomername.Text = ""
            SetItemUploadGridsHeader()
            txtfilelocation.Text = ""
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub txtCustomerCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtcustomercode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Try
            If KeyCode = System.Windows.Forms.Keys.F1 Then
                Call cmdcustomercode_Click(cmdcustomercode, New System.EventArgs())
            End If
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub TxtCustomerCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtcustomercode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then
            Call TxtCustomerCode_Validating(txtcustomercode, New System.ComponentModel.CancelEventArgs(False))
            If Len(Trim(txtcustomercode.Text)) <> 0 Then
                System.Windows.Forms.SendKeys.Send("{tab}")
            End If
        End If
        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = System.Windows.Forms.Keys.Back Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 22 Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
            KeyAscii = 0
        ElseIf KeyAscii = System.Windows.Forms.Keys.Return Then
            If Len(Trim(txtcustomercode.Text)) <> 0 Then
                Call TxtCustomerCode_Validating(txtcustomercode, New System.ComponentModel.CancelEventArgs(False))
            End If
        End If
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub TxtCustomerCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtcustomercode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim oRs As ADODB.Recordset
        Dim strSQL As String
        If Len(txtcustomercode.Text) > 0 Then
            txtcustomercode.Text = Replace(txtcustomercode.Text, "'", "")
            strSQL = "select a.customer_Code, a.cust_name from Customer_Mst a where a.customer_code='" & Trim(txtcustomercode.Text) & "' and isnull(manual_ds_closure,0)=0  and a.Unit_code = '" & gstrUNITID & "' and ((isnull(a.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= a.deactive_date))"
            oRs = New ADODB.Recordset
            oRs = mP_Connection.Execute(strSQL)
            If Not (oRs.EOF And oRs.BOF) Then
                lblcustomername.Text = oRs.Fields("cust_name").Value
                strSQL = "select dbo.UDF_ISDOCKCODE_INVOICE( '" & gstrUNITID & "','" & txtcustomercode.Text.Trim & "','INV','F' )"
                TxtLRNO.Enabled = True
                If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
                    TxtLRNO.MaxLength = 11
                Else
                    TxtLRNO.MaxLength = 30
                End If
            Else
                MsgBox("Invalid Customer Code", MsgBoxStyle.Information, ResolveResString(100))
                lblcustomername.Text = ""
                txtcustomercode.Text = ""
                txtcustomercode.Enabled = True
                txtcustomercode.Focus()
            End If
            oRs = Nothing
        Else
            lblcustomername.Text = ""
        End If
        GoTo EventExitSub
ErrHandler:
        oRs = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub FillItemSearchCategory()
        Try
            With CmbSearchBy
                .DataSource = Nothing
                .Items.Clear()
                .DataSource = [Enum].GetNames(GetType(enmInvoiceItemSearch))
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub SearchItem()
        Dim intCounter As Integer
        Dim strText As String
        Dim Col As enmInvoiceItemSearch
        Try

            If Len(TxtSearch.Text) = 0 Then Exit Sub
            For intCounter = 0 To dgvItemDetail.Rows.Count - 1
                If CmbSearchBy.Text = "SaleOrder" Then
                    strText = dgvItemDetail.Rows(intCounter).Cells(enmInvoiceItem.SaleOrder).Value
                    If Trim(UCase(Mid(strText, 1, Len(TxtSearch.Text)))) = Trim(UCase(TxtSearch.Text)) Then
                        dgvItemDetail.CurrentCell = dgvItemDetail.Rows(intCounter).Cells(enmInvoiceItem.SaleOrder)
                        Exit For
                    End If
                ElseIf CmbSearchBy.Text = "ItemCode" Then
                    strText = dgvItemDetail.Rows(intCounter).Cells(enmInvoiceItem.ItemCode).Value
                    If Trim(UCase(Mid(strText, 1, Len(TxtSearch.Text)))) = Trim(UCase(TxtSearch.Text)) Then
                        dgvItemDetail.CurrentCell = dgvItemDetail.Rows(intCounter).Cells(enmInvoiceItem.ItemCode)
                        Exit For
                    End If
                ElseIf CmbSearchBy.Text = "ItemDesc" Then
                    strText = dgvItemDetail.Rows(intCounter).Cells(enmInvoiceItem.ItemDesc).Value
                    If Trim(UCase(Mid(strText, 1, Len(TxtSearch.Text)))) = Trim(UCase(TxtSearch.Text)) Then
                        dgvItemDetail.CurrentCell = dgvItemDetail.Rows(intCounter).Cells(enmInvoiceItem.ItemDesc)
                        Exit For
                    End If
                ElseIf CmbSearchBy.Text = "Cust_DrgwNo" Then
                    strText = dgvItemDetail.Rows(intCounter).Cells(enmInvoiceItem.Cust_DrgwNo).Value
                    If Trim(UCase(Mid(strText, 1, Len(TxtSearch.Text)))) = Trim(UCase(TxtSearch.Text)) Then
                        dgvItemDetail.CurrentCell = dgvItemDetail.Rows(intCounter).Cells(enmInvoiceItem.Cust_DrgwNo)
                        Exit For
                    End If
                ElseIf CmbSearchBy.Text = "Cust_DrgDesc" Then
                    strText = dgvItemDetail.Rows(intCounter).Cells(enmInvoiceItem.Cust_DrgDesc).Value
                    If Trim(UCase(Mid(strText, 1, Len(TxtSearch.Text)))) = Trim(UCase(TxtSearch.Text)) Then
                        dgvItemDetail.CurrentCell = dgvItemDetail.Rows(intCounter).Cells(enmInvoiceItem.Cust_DrgDesc)
                        Exit For
                    End If
                End If
            Next
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub txtcustomerhelp_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtcustomerhelp.TextChanged
        Try
            lblcustname.Text = ""
            'dgvItemDetail.MaxCols = 0
            'dgvItemDetail.MaxRows = 0
            chkall.Enabled = False
            cmdgenerateformat.Enabled = False
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub txtcustomerhelp_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtcustomerhelp.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call cmdCustHelp_Click(cmdcusthelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtcustomerhelp_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtcustomerhelp.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then
            Call txtcustomerhelp_Validating(txtcustomerhelp, New System.ComponentModel.CancelEventArgs(False))
            If Len(Trim(txtcustomerhelp.Text)) <> 0 Then
                System.Windows.Forms.SendKeys.Send("{tab}")
            End If
        End If
        If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = System.Windows.Forms.Keys.Back Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 22 Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
            KeyAscii = 0
        ElseIf KeyAscii = System.Windows.Forms.Keys.Return Then
            If Len(Trim(txtcustomerhelp.Text)) <> 0 Then
                Call txtcustomerhelp_Validating(txtcustomerhelp, New System.ComponentModel.CancelEventArgs(False))
            End If
        End If
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtcustomerhelp_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtcustomerhelp.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim oRs As ADODB.Recordset
        Dim strSQL As String
        If Len(txtcustomerhelp.Text) > 0 Then
            txtcustomerhelp.Text = Replace(txtcustomerhelp.Text, "'", "")
            strSQL = "select a.customer_Code, a.cust_name from Customer_Mst a where a.customer_code='" & Trim(txtcustomerhelp.Text) & "' and isnull(InvoiceUploading,0)=1  and a.Unit_code = '" & gstrUNITID & "' and ((isnull(a.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= a.deactive_date))"
            oRs = New ADODB.Recordset
            oRs = mP_Connection.Execute(strSQL)
            If Not (oRs.EOF And oRs.BOF) Then
                lblcustname.Text = oRs.Fields("cust_name").Value
            Else
                MsgBox("Invalid Customer Code", MsgBoxStyle.Information, ResolveResString(100))
                lblcustname.Text = ""
                txtcustomerhelp.Text = ""
                txtcustomerhelp.Enabled = True
                txtcustomerhelp.Focus()
            End If
            oRs = Nothing
        Else
            lblcustname.Text = ""
        End If
        GoTo EventExitSub
ErrHandler:
        oRs = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub TxtSearch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtSearch.TextChanged
        On Error GoTo ErrHandler
        Call SearchItem()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function SAVEDATA() As Boolean
        Dim strInsert As String
        Dim strSalesOrderNo As String = ""
        Dim strAmendmentNo As String = ""
        Dim strItemCode As String = ""
        Dim strCustDrgNo As String = ""
        Dim strCustDrgDesc As String = ""
        Dim strQty As String = ""
        Dim strRate As String = ""
        Dim strFromBox As String = ""
        Dim strToBox As String = ""
        Dim StrMessage As String
        Dim strCurrency As String = ""
        Dim strPaymentTerm As String = ""
        Dim strCustMtrl As String = ""
        Dim strToolCost As String = ""
        Dim strHSNcode As String = ""
        Dim CGSTTXRT_TYPE As String = ""
        Dim SGSTTXRT_TYPE As String = ""
        Dim UTGSTTXRT_TYPE As String = ""
        Dim IGSTTXRT_TYPE As String = ""
        Dim Packing As String = ""
        Dim pkg_amount As String = ""
        Dim MaxNoifItem As Integer = 0
        Dim objfso As New Scripting.FileSystemObject

        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)

            If Validations() = True Then
                Dim intDocNo As Integer = SqlConnectionclass.ExecuteScalar("select isnull(max( UploadDoc_No),0) + 1 from SalesChallan_Dtl_Upload  where ip_address='" & gstrIpaddressWinSck & "' and unit_code='" & gstrUNITID & "'")
                MaxNoifItem = SqlConnectionclass.ExecuteScalar("select isnull(MaximumItemsInInvoices,0) from Sales_Parameter  where unit_code='" & gstrUNITID & "'")
                If MaxNoifItem = 0 Then
                    MsgBox("Please set MaxNoifItem in sales parameter", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Function
                End If
                For intcounter As Integer = 0 To dgvUpload.Rows.Count - 1
                    strSalesOrderNo = dgvUpload.Rows(intcounter).Cells(enmUploadItem.SaleOrder).Value
                    strAmendmentNo = dgvUpload.Rows(intcounter).Cells(enmUploadItem.AmendmentNo).Value
                    strItemCode = dgvUpload.Rows(intcounter).Cells(enmUploadItem.ItemCode).Value
                    strCustDrgNo = dgvUpload.Rows(intcounter).Cells(enmUploadItem.Cust_DrgwNo).Value
                    strCustDrgDesc = dgvUpload.Rows(intcounter).Cells(enmUploadItem.Cust_DrgDesc).Value
                    strQty = dgvUpload.Rows(intcounter).Cells(enmUploadItem.Qty).Value
                    strRate = dgvUpload.Rows(intcounter).Cells(enmUploadItem.Rate).Value
                    strFromBox = dgvUpload.Rows(intcounter).Cells(enmUploadItem.FromBox).Value
                    strToBox = dgvUpload.Rows(intcounter).Cells(enmUploadItem.ToBox).Value
                    strHSNcode = dgvUpload.Rows(intcounter).Cells(enmUploadItem.HSNCode).Value
                    CGSTTXRT_TYPE = dgvUpload.Rows(intcounter).Cells(enmUploadItem.CGSTTXRT_TYPE).Value
                    SGSTTXRT_TYPE = dgvUpload.Rows(intcounter).Cells(enmUploadItem.SGSTTXRT_TYPE).Value
                    UTGSTTXRT_TYPE = dgvUpload.Rows(intcounter).Cells(enmUploadItem.UTGSTTXRT_TYPE).Value
                    IGSTTXRT_TYPE = dgvUpload.Rows(intcounter).Cells(enmUploadItem.IGSTTXRT_TYPE).Value
                    strCurrency = dgvUpload.Rows(intcounter).Cells(enmUploadItem.Currency_Code).Value
                    strPaymentTerm = dgvUpload.Rows(intcounter).Cells(enmUploadItem.Payment_Terms).Value
                    strCustMtrl = dgvUpload.Rows(intcounter).Cells(enmUploadItem.CUST_MTRL).Value
                    strToolCost = dgvUpload.Rows(intcounter).Cells(enmUploadItem.TOOL_COST).Value
                    Packing = dgvUpload.Rows(intcounter).Cells(enmUploadItem.PACKING).Value


                    strInsert = "insert into SalesChallan_Dtl_Upload([UploadDoc_No],[Location_Code],[Vehicle_No],[Invoice_Date], " &
                    "[Account_Code],[Cust_Ref],[Amendment_No],[Item_Code],[Sales_Quantity],[From_Box],[To_Box],[Cust_Item_Code],[Cust_Item_Desc], " &
                    "[Carriage_Name],[Invoice_Type],[Cust_Name],[Sub_Category],[Ent_UserId],[Ent_dt],[Upd_Userid],[Upd_dt], " &
                    "[UNIT_CODE],IP_ADDRESS,Rate,HSNSACCODE,CGSTTXRT_TYPE,SGSTTXRT_TYPE,IGSTTXRT_TYPE,UTGSTTXRT_TYPE,Currency_Code, " &
                    "Payment_Terms,CUST_MTRL,TOOL_COST,Packing,TRANSPORT_TYPE,LorryNo_Date,DeliveryNoteNo) values " &
                    "(" & intDocNo & ",'" & gstrUNITID & "','" & txtVehNo.Text & "',getdate(),'" & txtcustomercode.Text & "', " &
                    "'" & strSalesOrderNo & "','" & strAmendmentNo & "','" & strItemCode & "', '" & strQty & "', " &
                    "'" & strFromBox & "','" & strToBox & "','" & strCustDrgNo & "','" & strCustDrgDesc & "', " &
                    "'" & txtCarrServices.Text & "','INV','" & lblcustomername.Text & "','F','" & mP_User & "',getdate(), " &
                    "'" & mP_User & "',getdate(),'" & gstrUNITID & "','" & gstrIpaddressWinSck & "','" & strRate & "', " &
                    "'" & strHSNcode & "','" & CGSTTXRT_TYPE & "','" & SGSTTXRT_TYPE & "','" & IGSTTXRT_TYPE & "','" & UTGSTTXRT_TYPE & "', " &
                    "'" & strCurrency & "','" & strPaymentTerm & "','" & strCustMtrl & "','" & strToolCost & "','" & Packing & "','" & Mid(Trim(CmbTransType.Text), 1, 1) & "','" & TxtLRNO.Text & "','" & txtDeliveryNoteNo.Text.Trim & "')"
                    SqlConnectionclass.ExecuteNonQuery(strInsert)

                Next

                Dim NoifItemInvoice As Integer = 0
                Dim intRowNo As Integer = 0
                Dim intInvoiceNo As Integer = 0
                Dim strFirstSaleOrder As String = ""
                Dim strFirstAmedment As String = ""
                strSalesOrderNo = ""
                strAmendmentNo = ""
                strItemCode = ""
                strCustDrgNo = ""
                strInsert = "select ROW_NUMBER() OVER (partition by  cust_ref,amendment_no ORDER BY  cust_ref,amendment_no ) row_num, * from SalesChallan_Dtl_Upload  where ip_address='" & gstrIpaddressWinSck & "' and uploaddoc_no='" & intDocNo & "'  order by cust_ref,amendment_no"
                Dim dt As DataTable = SqlConnectionclass.GetDataTable(strInsert)
                If dt.Rows.Count > 0 Then
                    For Each Row As DataRow In dt.Rows
                        intRowNo = Convert.ToInt32(Row("row_num"))
                        strItemCode = Convert.ToString(Row("Item_Code"))
                        strCustDrgNo = Convert.ToString(Row("Cust_Item_Code"))
                        If intRowNo = 1 Then
                            strFirstSaleOrder = Convert.ToString(Row("Cust_Ref"))
                            strFirstAmedment = Convert.ToString(Row("Amendment_no"))
                            intInvoiceNo = intInvoiceNo + 1
                            NoifItemInvoice = 1
                        Else
                            strSalesOrderNo = Convert.ToString(Row("Cust_Ref"))
                            strAmendmentNo = Convert.ToString(Row("Amendment_no"))
                            If Trim(strFirstSaleOrder) = Trim(strSalesOrderNo) Then

                                'NoifItemInvoice = NoifItemInvoice + 1
                                'If NoifItemInvoice > 7 And NoifItemInvoice = 8 Then
                                '    intInvoiceNo = intInvoiceNo + 1
                                'ElseIf NoifItemInvoice > 14 And NoifItemInvoice = 15 Then
                                '    intInvoiceNo = intInvoiceNo + 1
                                'ElseIf NoifItemInvoice > 21 And NoifItemInvoice = 22 Then
                                '    intInvoiceNo = intInvoiceNo + 1
                                'ElseIf NoifItemInvoice > 28 And NoifItemInvoice = 29 Then
                                '    intInvoiceNo = intInvoiceNo + 1
                                'End If
                            Else
                                intInvoiceNo = intInvoiceNo + 1
                            End If
                        End If
                        Dim strSql = "Update SalesChallan_Dtl_Upload set TMP_DOC_NO=" & intInvoiceNo & "  where unit_code='" & gstrUNITID & "' and  ip_address='" & gstrIpaddressWinSck & "' and uploaddoc_no='" & intDocNo & "'  and cust_ref='" & strFirstSaleOrder & "' and amendment_no='" & strFirstAmedment & "' and Item_code='" & strItemCode & "' and Cust_Item_Code='" & strCustDrgNo & "'  "
                        SqlConnectionclass.ExecuteNonQuery(strSql)
                    Next

                End If
                SqlConnectionclass.CloseGlobalConnection()
                SqlConnectionclass.OpenGlobalConnection()
                SqlConnectionclass.BeginTrans()

                Using sqlCmd As SqlCommand = New SqlCommand
                    With sqlCmd
                        .CommandText = "USP_GENERATE_AUTO_INVOICE_UPLOADING"
                        .CommandTimeout = 300
                        .CommandType = CommandType.StoredProcedure
                        .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                        .Parameters.Add("@CUSTOMER_CODE", SqlDbType.VarChar, 10).Value = txtcustomercode.Text
                        .Parameters.Add("@UPLOAD_NO", SqlDbType.VarChar, 10).Value = intDocNo
                        .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                        .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                        SqlConnectionclass.ExecuteNonQuery(sqlCmd)

                        If Convert.ToString(.Parameters("@MESSAGE").Value).Length > 0 Then
                            SqlConnectionclass.RollbackTran()
                            MessageBox.Show(Convert.ToString(.Parameters("@MESSAGE").Value), "eMPro", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                            Exit Function
                        Else
                            SqlConnectionclass.CommitTran()
                            fillInvoiceData(intDocNo)

                            cmdsave.Enabled = False
                        End If
                    End With
                End Using

                MsgBox("Invoice Generated Successfully ", MsgBoxStyle.Information, ResolveResString(100))

            End If

            'ResetData()
            Return True

ErrHandler:
        Catch ex As Exception
            SqlConnectionclass.RollbackTran()
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    Private Sub fillInvoiceData(ByRef intUploadDocNo As Integer)
        Try
            Dim intCount As Integer = 0
            Dim strSql As String = "SELECT * FROM SalesChallan_Dtl_Upload WHERE UNIT_CODE='" & gstrUNITID & "'  and IP_ADDRESS='" & gstrIpaddressWinSck & "' and UploadDoc_No='" & intUploadDocNo & "' ORDER BY DOC_NO "
            Dim dt As DataTable = SqlConnectionclass.GetDataTable(strSql)
            If dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.TempInvoiceNo).Value = dr("DOC_NO")
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.SaleOrder).Value = dr("Cust_Ref")
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.AmendmentNo).Value = dr("Amendment_No")
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.ItemCode).Value = dr("Item_Code")
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.Cust_DrgwNo).Value = dr("Cust_Item_Code")
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.Cust_DrgDesc).Value = dr("Cust_Item_Desc")
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.Qty).Value = dr("Sales_Quantity")
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.Rate).Value = dr("Rate")
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.FromBox).Value = dr("From_Box")
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.ToBox).Value = dr("To_Box")
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.HSNCode).Value = dr("HSNSACCODE")
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.CGSTTXRT_TYPE).Value = dr("CGSTTXRT_TYPE")
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.SGSTTXRT_TYPE).Value = dr("SGSTTXRT_TYPE")
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.UTGSTTXRT_TYPE).Value = dr("UTGSTTXRT_TYPE")
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.IGSTTXRT_TYPE).Value = dr("IGSTTXRT_TYPE")
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.Currency_Code).Value = dr("Currency_Code")
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.Payment_Terms).Value = dr("Payment_Terms")
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.CUST_MTRL).Value = dr("CUST_MTRL")
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.TOOL_COST).Value = dr("TOOL_COST")
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.PACKING).Value = dr("PACKING")
                    Dim strSTockQty As Integer = SqlConnectionclass.ExecuteScalar("Select Cur_Bal From ItemBal_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Location_Code ='" & strStockLocation & "' and item_Code ='" & dr("Item_Code") & "'")
                    dgvUpload.Rows(intCount).Cells(enmUploadItem.CurBal).Value = strSTockQty

                    intCount = intCount + 1
                Next
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub
    Private Function get_itemdescription(ByVal stritemcode As String) As String
        On Error GoTo ErrHandler
        Dim objgetitemdesc As New ClsResultSetDB
        objgetitemdesc.GetResult("select description from item_mst where item_code='" & stritemcode & "' and Unit_code = '" & gstrUNITID & "'")
        If Not (objgetitemdesc.EOFRecord And objgetitemdesc.BOFRecord) Then
            get_itemdescription = objgetitemdesc.GetValue("description")
        Else
            get_itemdescription = ""
        End If
        objgetitemdesc = Nothing
        Exit Function
ErrHandler:
        objgetitemdesc = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Function Validatecustomer_drgno(ByRef stritemcode As String) As Boolean
        On Error GoTo ErrHandler
        Dim objgetdrgno As New ClsResultSetDB
        If SSTab1.SelectedIndex = 0 Then
            objgetdrgno.GetResult("select top 1 cust_drgno from custitem_mst where account_code='" & Trim(txtcustomerhelp.Text) & "' and item_code='" & stritemcode & "' and active=1 and Unit_code = '" & gstrUNITID & "'")
        Else
            objgetdrgno.GetResult("select top 1 cust_drgno from custitem_mst where account_code='" & Trim(txtcustomercode.Text) & "' and item_code='" & stritemcode & "' and active=1 and Unit_code = '" & gstrUNITID & "'")
        End If
        If Not (objgetdrgno.BOFRecord And objgetdrgno.EOFRecord) Then
            Return True
        Else
            Return False
        End If
        objgetdrgno = Nothing
        Exit Function
ErrHandler:
        objgetdrgno = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function


    Private Sub upload_InvoiceData()
        On Error GoTo ErrHandler
        Dim objexl As New Excel.Application
        Dim objfso As New Scripting.FileSystemObject
        Dim row As Long
        Dim strSaleOrder As String
        Dim strAmedmentNo As String
        Dim stritemcode As String
        Dim strItemDesc As String
        Dim strcustomerDrgNo As String
        Dim strcustomerDrgDesc As String
        Dim strQuantity As String
        Dim strFromBox As String
        Dim strToBox As String
        Dim strcustomercode As String
        Dim strDate As String

        Dim col As Long
        Dim row1 As Long
        Dim col1 As Long
        Dim row2 As Long
        Dim col2 As Long
        Dim strSQL As String
        If objfso.FileExists(txtfilelocation.Text) = False Then
            MsgBox("File Does Not Exists", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        Else
            objexl = New Excel.Application
            objexl.Workbooks.Open(Trim(Me.txtfilelocation.Text))
            If validate_excel_format(objexl) = True Then
                row = 2
                strcustomercode = objexl.Cells(row, 1).Value
                If UCase(strcustomercode) <> UCase(txtcustomercode.Text) Then
                    MsgBox("Customer Code In File Is Not Matching With Customer Code Selected", MsgBoxStyle.Information, ResolveResString(100))
                    objexl.Quit()
                    objexl = Nothing
                    objfso = Nothing
                    Exit Sub
                Else
                    SetItemUploadGridsHeader()
                    Dim intCount As Integer = 0
                    dgvUpload.Rows.Clear()
                    'dgvUpload.Rows.Add(objexl.Rows.Count)
                    While (row <= objexl.Rows.Count)
                        strSaleOrder = objexl.Cells(row, 3).Value
                        strAmedmentNo = objexl.Cells(row, 4).Value
                        stritemcode = objexl.Cells(row, 5).Value
                        strItemDesc = objexl.Cells(row, 6).Value
                        strcustomerDrgNo = objexl.Cells(row, 7).Value
                        strcustomerDrgDesc = objexl.Cells(row, 8).Value
                        strQuantity = objexl.Cells(row, 9).Value
                        strFromBox = objexl.Cells(row, 10).Value
                        strToBox = objexl.Cells(row, 11).Value
                        If strSaleOrder <> "" And stritemcode <> "" And strcustomerDrgNo <> "" Then
                            With dgvUpload
                                dgvUpload.Rows.Add(1)
                                dgvUpload.Rows(intCount).Cells(enmUploadItem.SaleOrder).Value = strSaleOrder
                                dgvUpload.Rows(intCount).Cells(enmUploadItem.AmendmentNo).Value = strAmedmentNo
                                dgvUpload.Rows(intCount).Cells(enmUploadItem.ItemCode).Value = stritemcode
                                dgvUpload.Rows(intCount).Cells(enmUploadItem.ItemDesc).Value = strItemDesc
                                dgvUpload.Rows(intCount).Cells(enmUploadItem.Cust_DrgwNo).Value = strcustomerDrgNo
                                dgvUpload.Rows(intCount).Cells(enmUploadItem.Cust_DrgDesc).Value = strcustomerDrgDesc
                                dgvUpload.Rows(intCount).Cells(enmUploadItem.Qty).Value = strQuantity
                                dgvUpload.Rows(intCount).Cells(enmUploadItem.FromBox).Value = strFromBox
                                dgvUpload.Rows(intCount).Cells(enmUploadItem.ToBox).Value = strToBox
                                dgvUpload.Rows(intCount).Cells(enmUploadItem.CUST_NAME).Value = lblcustomername.Text
                                Dim strSTockQty As Integer = SqlConnectionclass.ExecuteScalar("Select Cur_Bal From ItemBal_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Location_Code ='" & strStockLocation & "' and item_Code ='" & stritemcode & "'")
                                dgvUpload.Rows(intCount).Cells(enmUploadItem.CurBal).Value = strSTockQty
                            End With
                            row = row + 1
                            intCount = intCount + 1
                        Else
                            Exit While
                        End If

                    End While
                    objexl.Quit()
                    objexl = Nothing
                    objfso = Nothing
                    If Validations() = False Then
                        'ResetData()
                        SetItemUploadGridsHeader()
                        Exit Sub

                    End If

                End If
            Else
                MsgBox("Excel Format Is Not Correct", MsgBoxStyle.Information, ResolveResString(100))
                objexl.Quit()
                objexl = Nothing
                objfso = Nothing
            End If
        End If
        Exit Sub
ErrHandler:
        objexl = Nothing
        objfso = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub SetItemGridsHeader()
        Try

            dgvItemDetail.Columns.Clear()

            Dim objChkBox As New DataGridViewCheckBoxColumn
            objChkBox.Name = "Selection"
            objChkBox.HeaderText = "Select"

            dgvItemDetail.Columns.Add(objChkBox)
            dgvItemDetail.Columns.Add("SaleOrder", "CustRef/Po No")
            dgvItemDetail.Columns.Add("AmendmentNo", "Amendment")
            dgvItemDetail.Columns.Add("ItemCode", "Item Code")
            dgvItemDetail.Columns.Add("ItemDesc", "Item Desc")
            dgvItemDetail.Columns.Add("Cust_DrgwNo", "Cust DrgNo")
            dgvItemDetail.Columns.Add("Cust_DrgDesc", "Drg Desc")
            dgvItemDetail.Columns.Add("HSNCode", "HSN Code")
            dgvItemDetail.Columns.Add("CurBal", "Cur Bal")

            dgvItemDetail.Columns(enmInvoiceItem.status).Visible = True
            dgvItemDetail.Columns(enmInvoiceItem.status).Width = 50
            dgvItemDetail.Columns(enmInvoiceItem.SaleOrder).Width = 140
            dgvItemDetail.Columns(enmInvoiceItem.AmendmentNo).Width = 100
            dgvItemDetail.Columns(enmInvoiceItem.ItemCode).Width = 100
            dgvItemDetail.Columns(enmInvoiceItem.ItemDesc).Width = 150
            dgvItemDetail.Columns(enmInvoiceItem.Cust_DrgwNo).Width = 100
            dgvItemDetail.Columns(enmInvoiceItem.Cust_DrgDesc).Width = 150
            dgvItemDetail.Columns(enmInvoiceItem.HSNCode).Width = 150
            dgvItemDetail.Columns(enmInvoiceItem.CurBal).Width = 150
            dgvItemDetail.Columns(enmInvoiceItem.HSNCode).Visible = False
            dgvItemDetail.Columns(enmInvoiceItem.CurBal).Visible = False

            dgvItemDetail.Columns(enmInvoiceItem.status).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvItemDetail.Columns(enmInvoiceItem.SaleOrder).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvItemDetail.Columns(enmInvoiceItem.AmendmentNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvItemDetail.Columns(enmInvoiceItem.ItemCode).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvItemDetail.Columns(enmInvoiceItem.ItemDesc).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvItemDetail.Columns(enmInvoiceItem.Cust_DrgwNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvItemDetail.Columns(enmInvoiceItem.Cust_DrgDesc).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvItemDetail.Columns(enmInvoiceItem.HSNCode).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvItemDetail.Columns(enmInvoiceItem.CurBal).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            dgvItemDetail.Columns(enmInvoiceItem.SaleOrder).ReadOnly = True
            dgvItemDetail.Columns(enmInvoiceItem.AmendmentNo).ReadOnly = True
            dgvItemDetail.Columns(enmInvoiceItem.ItemCode).ReadOnly = True
            dgvItemDetail.Columns(enmInvoiceItem.ItemDesc).ReadOnly = True
            dgvItemDetail.Columns(enmInvoiceItem.Cust_DrgwNo).ReadOnly = True
            dgvItemDetail.Columns(enmInvoiceItem.Cust_DrgDesc).ReadOnly = True
            dgvItemDetail.Columns(enmInvoiceItem.HSNCode).ReadOnly = True
            dgvItemDetail.Columns(enmInvoiceItem.CurBal).ReadOnly = True

            dgvItemDetail.Columns(enmInvoiceItem.status).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvItemDetail.Columns(enmInvoiceItem.ItemCode).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvItemDetail.Columns(enmInvoiceItem.ItemDesc).SortMode = DataGridViewColumnSortMode.NotSortable

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub SetItemUploadGridsHeader()
        Try

            dgvUpload.Columns.Clear()

            dgvUpload.Columns.Add("TempInvoiceNo", "Temp InvoiceNo")
            dgvUpload.Columns.Add("SaleOrder", "CustRef/Po No")
            dgvUpload.Columns.Add("AmendmentNo", "Amendment")
            dgvUpload.Columns.Add("ItemCode", "Item Code")
            dgvUpload.Columns.Add("ItemDesc", "Item Desc")
            dgvUpload.Columns.Add("Cust_DrgwNo", "Cust DrgNo")
            dgvUpload.Columns.Add("Cust_DrgDesc", "Drg Desc")
            dgvUpload.Columns.Add("CurBal", "Cur Bal")
            dgvUpload.Columns.Add("Qty", "Quantity")
            dgvUpload.Columns.Add("Rate", "Rate")
            dgvUpload.Columns.Add("FromBox", "From Box")
            dgvUpload.Columns.Add("ToBox", "To Box")
            dgvUpload.Columns.Add("Currency_Code", "Currency Code")
            dgvUpload.Columns.Add("Payment_Terms", "Payment Terms")
            dgvUpload.Columns.Add("CUST_MTRL", "CUST MTRL")
            dgvUpload.Columns.Add("TOOL_COST", "TOOL COST")
            dgvUpload.Columns.Add("HSNCode", "HSN Code")
            dgvUpload.Columns.Add("CGSTTXRT_TYPE", "CGSTTXRT TYPE")
            dgvUpload.Columns.Add("SGSTTXRT_TYPE", "SGSTTXRT TYPE")
            dgvUpload.Columns.Add("IGSTTXRT_TYPE", "IGSTTXRT TYPE")
            dgvUpload.Columns.Add("UTGSTTXRT_TYPE", "UTGSTTXRT TYPE")
            dgvUpload.Columns.Add("CUST_NAME", "CUST NAME")
            dgvUpload.Columns.Add("PACKING", "PACKING")


            dgvUpload.Columns(enmUploadItem.TempInvoiceNo).Width = 70
            dgvUpload.Columns(enmUploadItem.SaleOrder).Width = 100
            dgvUpload.Columns(enmUploadItem.AmendmentNo).Width = 100
            dgvUpload.Columns(enmUploadItem.ItemCode).Width = 80
            dgvUpload.Columns(enmUploadItem.ItemDesc).Width = 100
            dgvUpload.Columns(enmUploadItem.Cust_DrgwNo).Width = 80
            dgvUpload.Columns(enmUploadItem.Cust_DrgDesc).Width = 100
            dgvUpload.Columns(enmUploadItem.HSNCode).Width = 50
            dgvUpload.Columns(enmUploadItem.CurBal).Width = 50
            dgvUpload.Columns(enmUploadItem.Qty).Width = 50
            dgvUpload.Columns(enmUploadItem.Rate).Width = 50
            dgvUpload.Columns(enmUploadItem.FromBox).Width = 50
            dgvUpload.Columns(enmUploadItem.ToBox).Width = 50
            dgvUpload.Columns(enmUploadItem.Currency_Code).Width = 50
            dgvUpload.Columns(enmUploadItem.Payment_Terms).Width = 50
            dgvUpload.Columns(enmUploadItem.CUST_MTRL).Width = 50
            dgvUpload.Columns(enmUploadItem.TOOL_COST).Width = 50
            dgvUpload.Columns(enmUploadItem.CGSTTXRT_TYPE).Width = 50
            dgvUpload.Columns(enmUploadItem.SGSTTXRT_TYPE).Width = 50
            dgvUpload.Columns(enmUploadItem.IGSTTXRT_TYPE).Width = 50
            dgvUpload.Columns(enmUploadItem.UTGSTTXRT_TYPE).Width = 50
            dgvUpload.Columns(enmUploadItem.CUST_NAME).Width = 50
            dgvUpload.Columns(enmUploadItem.PACKING).Width = 50

            dgvUpload.Columns(enmUploadItem.HSNCode).Visible = True
            'dgvUpload.Columns(enmUploadItem.CurBal).Visible = True
            dgvUpload.Columns(enmUploadItem.Rate).Visible = True


            dgvUpload.Columns(enmUploadItem.TempInvoiceNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvUpload.Columns(enmUploadItem.SaleOrder).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvUpload.Columns(enmUploadItem.AmendmentNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvUpload.Columns(enmUploadItem.ItemCode).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvUpload.Columns(enmUploadItem.ItemDesc).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvUpload.Columns(enmUploadItem.Cust_DrgwNo).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvUpload.Columns(enmUploadItem.Cust_DrgDesc).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvUpload.Columns(enmUploadItem.HSNCode).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvUpload.Columns(enmUploadItem.CurBal).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvUpload.Columns(enmUploadItem.Qty).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvUpload.Columns(enmUploadItem.FromBox).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvUpload.Columns(enmUploadItem.ToBox).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            dgvUpload.Columns(enmUploadItem.TempInvoiceNo).ReadOnly = True
            dgvUpload.Columns(enmUploadItem.SaleOrder).ReadOnly = True
            dgvUpload.Columns(enmUploadItem.AmendmentNo).ReadOnly = True
            dgvUpload.Columns(enmUploadItem.ItemCode).ReadOnly = True
            dgvUpload.Columns(enmUploadItem.ItemDesc).ReadOnly = True
            dgvUpload.Columns(enmUploadItem.Cust_DrgwNo).ReadOnly = True
            dgvUpload.Columns(enmUploadItem.Cust_DrgDesc).ReadOnly = True
            dgvUpload.Columns(enmUploadItem.HSNCode).ReadOnly = True
            dgvUpload.Columns(enmUploadItem.CurBal).ReadOnly = True
            dgvUpload.Columns(enmUploadItem.Qty).ReadOnly = True
            dgvUpload.Columns(enmUploadItem.FromBox).ReadOnly = True
            dgvUpload.Columns(enmUploadItem.ToBox).ReadOnly = True
            dgvUpload.Columns(enmUploadItem.TOOL_COST).ReadOnly = True
            dgvUpload.Columns(enmUploadItem.CUST_MTRL).ReadOnly = True
            dgvUpload.Columns(enmUploadItem.CGSTTXRT_TYPE).ReadOnly = True
            dgvUpload.Columns(enmUploadItem.SGSTTXRT_TYPE).ReadOnly = True
            dgvUpload.Columns(enmUploadItem.IGSTTXRT_TYPE).ReadOnly = True
            dgvUpload.Columns(enmUploadItem.UTGSTTXRT_TYPE).ReadOnly = True


            dgvUpload.Columns(enmUploadItem.ItemCode).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvUpload.Columns(enmUploadItem.ItemDesc).SortMode = DataGridViewColumnSortMode.NotSortable

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function getfilename(ByVal strcustname As String) As String
        On Error GoTo ErrHandler
        Dim objgetdate As New ClsResultSetDB
        Dim strFileName As String
        objgetdate.GetResult("select convert(varchar(11),getdate(),106) + '-' + convert(varchar(8),getdate(),108) as date1")
        If Not objgetdate.EOFRecord Then
            strFileName = objgetdate.GetValue("date1")
        End If
        ' getfilename = "Forecast_" & strcustname & "_" & strFileName & ".xlsx"
        getfilename = "DSUpload_" & strcustname & "_" & strFileName & ".xls"
        objgetdate = Nothing
        Exit Function
ErrHandler:
        objgetdate = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub fillitems(ByRef objexl As Excel.Application)
        Dim intcounter As Long
        Dim varstatus As Object
        Dim varSaleOrder As Object
        Dim varAmendmentNo As Object
        Dim varItemCode As Object
        Dim varItemDesc As Object
        Dim varCustdrgno As Object
        Dim varCustDrgdesc As Object
        Dim itemcount As Long
        Dim intcounter1 As Long
        Dim interval As Long
        Dim dtDate As Date
        Dim datecounter As Long
        Dim objcalender As New ClsResultSetDB
        Try
            itemcount = 2
            interval = DateDiff(Microsoft.VisualBasic.DateInterval.Day, dtfromdate.Value, dttodate.Value) + 1
            With dgvItemDetail
                objexl.Cells._Default(1, 1).Font.Bold = True


                For intcount As Integer = 0 To dgvItemDetail.Rows.Count - 1

                    varstatus = Nothing
                    varSaleOrder = Nothing
                    varAmendmentNo = Nothing
                    varItemCode = Nothing
                    varItemDesc = Nothing
                    varCustdrgno = Nothing
                    varCustDrgdesc = Nothing
                    If dgvItemDetail.Rows(intcount).Cells(enmInvoiceItem.status).Value = True Then
                        varSaleOrder = dgvItemDetail.Rows(intcount).Cells(enmInvoiceItem.SaleOrder).Value
                        varAmendmentNo = dgvItemDetail.Rows(intcount).Cells(enmInvoiceItem.AmendmentNo).Value
                        varItemCode = dgvItemDetail.Rows(intcount).Cells(enmInvoiceItem.ItemCode).Value
                        varItemDesc = dgvItemDetail.Rows(intcount).Cells(enmInvoiceItem.ItemDesc).Value
                        varCustdrgno = dgvItemDetail.Rows(intcount).Cells(enmInvoiceItem.Cust_DrgwNo).Value
                        varCustDrgdesc = dgvItemDetail.Rows(intcount).Cells(enmInvoiceItem.Cust_DrgDesc).Value
                    End If

                    If dgvItemDetail.Rows(intcount).Cells(enmInvoiceItem.status).Value = True Then
                        objexl.Cells._Default(itemcount, 1) = txtcustomerhelp.Text
                        objexl.Cells._Default(itemcount, 2) = lblcustname.Text
                        objexl.Cells._Default(itemcount, 3) = varSaleOrder
                        objexl.Cells._Default(itemcount, 4) = varAmendmentNo
                        objexl.Cells._Default(itemcount, 5) = varItemCode
                        objexl.Cells._Default(itemcount, 6) = varItemDesc
                        objexl.Cells._Default(itemcount, 7) = varCustdrgno
                        objexl.Cells._Default(itemcount, 8) = varCustDrgdesc
                        objexl.Cells._Default(itemcount, 1).ColumnWidth = Len(varSaleOrder) + 5
                        objexl.Cells._Default(itemcount, 2).ColumnWidth = Len(lblcustname.Text) + 5
                        objexl.Cells._Default(itemcount, 3).ColumnWidth = Len(varSaleOrder) + 5
                        objexl.Cells._Default(itemcount, 4).ColumnWidth = Len(varAmendmentNo) + 5
                        objexl.Cells._Default(itemcount, 5).ColumnWidth = Len(varItemCode) + 5
                        objexl.Cells._Default(itemcount, 6).ColumnWidth = Len(varItemDesc) + 5
                        objexl.Cells._Default(itemcount, 7).ColumnWidth = Len(varCustdrgno) + 5
                        objexl.Cells._Default(itemcount, 8).ColumnWidth = Len(varCustDrgdesc) + 5
                        itemcount = itemcount + 1
                    End If
                Next
            End With

            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub clear_fields()
        On Error GoTo ErrHandler
        SetItemGridsHeader()
        SetItemUploadGridsHeader()
        lblcustname.Text = ""
        txtcustomerhelp.Text = ""
        dtfromdate.Value = GetServerDate()
        dttodate.Value = GetServerDate()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
 
    Private Function validate_excel_format(ByRef objexl As Excel.Application) As Boolean

        Dim row As Long
        Dim strcustomercode As String
        Dim stritemcode As String
        Dim strCustDrgNo As String
        Dim strItemDesc As String
        Dim strDate As String
        Dim lngquantity As Integer
        Dim col As Long
        Dim row1 As Long
        Dim col1 As Long
        Dim row2 As Long
        Dim col2 As Long
        Dim rowDrgNo As Long
        Dim rowItemDesc As Long
        Dim intRow As Long
        Dim strTransDate As String
        Dim strBalanceQty As Integer
        Try
            validate_excel_format = True
            row = 2
            col = 3

            While (row <= objexl.Rows.Count)
                stritemcode = objexl.Cells(row, 1).Value
                strCustDrgNo = objexl.Cells(row, 3).Value
                If stritemcode <> "" AndAlso strCustDrgNo <> "" Then
                    row = row + 1
                Else
                    Exit While
                End If

            End While
            Exit Function
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub dtfromdate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtfromdate.ValueChanged
        Try
            If dtfromdate.Value < GetServerDate() Then
                dtfromdate.Value = GetServerDate()
            ElseIf dtfromdate.Value > dttodate.Value Then
                dtfromdate.Value = GetServerDate()
            End If
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub dttodate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dttodate.ValueChanged
        Try
            If dttodate.Value < dtfromdate.Value Then
                dttodate.Value = dtfromdate.Value
            End If
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub cmdsave_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdsave.Click
        Try
            Dim strcustomercode As String
            Dim stritemlist As String
            Dim intvalue As Long

            If dgvUpload.Rows.Count > 0 Then
                If txtcustomercode.Text <> "" Then
                    If SAVEDATA() = False Then Exit Sub
                Else
                    MsgBox("Customer Code Can Not Be Blank", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If
            Else
                MsgBox("Please Upload DS First", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If

            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function Validations() As Boolean
        Dim strSalesOrderNo As String = ""
        Dim strAmendmentNo As String = ""
        Dim strItemCode As String = ""
        Dim strCustDrgNo As String = ""
        Dim strQty As String = ""
        Dim strFromBox As String = ""
        Dim strtoBox As String = ""
        Dim strInnerSalesOrderNo As String = ""
        Dim strInnerAmendmentNo As String = ""
        Dim strInnerItemCode As String = ""
        Dim strInnerCustDrgNo As String = ""
        Dim strInnerQty As String = ""
        Dim strQuery As String = ""
        Dim strExist As String = ""
        Dim dt As DataTable
        Try
            Dim strSQL = "select dbo.UDF_ISDOCKCODE_INVOICE( '" & gstrUNITID & "','" & txtcustomercode.Text.Trim & "','" & CmbInvType.Text.Trim.ToUpper & "','" & CmbInvSubType.Text.Trim.ToUpper & "' )"
            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
                If TxtLRNO.Text = "" Then
                    MsgBox("Please enter LR No", MsgBoxStyle.Information, ResolveResString(100))
                    Return False
                End If
            End If
            If Len(strStockLocation) = 0 Then
                MsgBox("Please Define Stock Location in Sales Conf First", MsgBoxStyle.Information, ResolveResString(100))
                Return False
            End If

            For intRow As Integer = 0 To dgvUpload.Rows.Count - 1
                strSalesOrderNo = dgvUpload.Rows(intRow).Cells(enmUploadItem.SaleOrder).Value
                strAmendmentNo = dgvUpload.Rows(intRow).Cells(enmUploadItem.AmendmentNo).Value
                strItemCode = dgvUpload.Rows(intRow).Cells(enmUploadItem.ItemCode).Value
                strCustDrgNo = dgvUpload.Rows(intRow).Cells(enmUploadItem.Cust_DrgwNo).Value
                strQty = dgvUpload.Rows(intRow).Cells(enmUploadItem.Qty).Value
                strFromBox = dgvUpload.Rows(intRow).Cells(enmUploadItem.FromBox).Value
                strtoBox = dgvUpload.Rows(intRow).Cells(enmUploadItem.ToBox).Value

                strQuery = "Select Item_code from item_mst (Nolock) where item_code='" & strItemCode & "' and unit_code='" & gstrUNITID & "'"
                strExist = SqlConnectionclass.ExecuteScalar(strQuery)
                If strExist = "" Then
                    MsgBox("Item Code is not correct at row no  - " & Convert.ToString(intRow + 1), MsgBoxStyle.Information, ResolveResString(100))
                    Return False
                End If
                strExist = ""

                strQuery = "select item_code from custitem_mst  (Nolock)  where account_code='" & txtcustomercode.Text & "' and item_code='" & strItemCode & "' and cust_drgno='" & strCustDrgNo & "'  and unit_code='" & gstrUNITID & "' and active=1 "
                strExist = SqlConnectionclass.ExecuteScalar(strQuery)
                If strExist = "" Then
                    MsgBox("Please check Customer Item mapiping at row no  - " & Convert.ToString(intRow + 1), MsgBoxStyle.Information, ResolveResString(100))
                    Return False
                End If
                strExist = ""

                strQuery = ""

                strQuery = "select top 1 Item_code from cust_ord_dtl  where UNIT_CODE='" + gstrUNITID + "' and account_code='" & Trim(txtcustomercode.Text) & "' and item_code='" & strItemCode & "' and cust_drgno = '" & strCustDrgNo & "' and cust_ref='" & strSalesOrderNo & "' and Amendment_no='" & strAmendmentNo & "' and  active_flag='A' and authorized_flag=1"
                strExist = SqlConnectionclass.ExecuteScalar(strQuery)
                If strExist = "" Then
                    MsgBox("Please check Sale Order mapping at row no  - " & Convert.ToString(intRow + 1), MsgBoxStyle.Information, ResolveResString(100))
                    Return False
                Else
                    strQuery = "select top 1 CURRENCY_CODE,TERM_PAYMENT from cust_ord_Hdr where UNIT_CODE='" + gstrUNITID + "' and account_code='" & Trim(txtcustomercode.Text) & "' and cust_ref='" & strSalesOrderNo & "' and Amendment_no='" & strAmendmentNo & "' and  active_flag='A' and authorized_flag=1"
                    dt = New DataTable
                    dt = SqlConnectionclass.GetDataTable(strQuery)
                    If dt.Rows.Count > 0 Then
                        dgvUpload.Rows(intRow).Cells(enmUploadItem.Currency_Code).Value = Convert.ToString(dt.Rows(0)("CURRENCY_CODE"))
                        dgvUpload.Rows(intRow).Cells(enmUploadItem.Payment_Terms).Value = Convert.ToString(dt.Rows(0)("TERM_PAYMENT"))
                    End If
                    strQuery = "select top 1 Rate,CUST_MTRL,TOOL_COST,PACKING from cust_ord_dtl where UNIT_CODE='" + gstrUNITID + "' and account_code='" & Trim(txtcustomercode.Text) & "' and item_code='" & strItemCode & "' and cust_drgno = '" & strCustDrgNo & "' and cust_ref='" & strSalesOrderNo & "' and Amendment_no='" & strAmendmentNo & "' and  active_flag='A' and authorized_flag=1"
                    dt = SqlConnectionclass.GetDataTable(strQuery)
                    dgvUpload.Rows(intRow).Cells(enmUploadItem.Rate).Value = Convert.ToString(dt.Rows(0)("Rate"))
                    dgvUpload.Rows(intRow).Cells(enmUploadItem.CUST_MTRL).Value = Convert.ToString(dt.Rows(0)("CUST_MTRL"))
                    dgvUpload.Rows(intRow).Cells(enmUploadItem.TOOL_COST).Value = Convert.ToString(dt.Rows(0)("TOOL_COST"))
                    dgvUpload.Rows(intRow).Cells(enmUploadItem.PACKING).Value = Convert.ToString(dt.Rows(0)("PACKING"))
                End If

                If Convert.ToInt32(strQty) <= 0 Then
                    MsgBox("Sale Quantity cannot be 0  at row no  - " & Convert.ToString(intRow + 1), MsgBoxStyle.Information, ResolveResString(100))
                    Return False
                ElseIf Convert.ToInt32(strFromBox) <= 0 Then
                    MsgBox("From Box cannot be 0  at row no  - " & Convert.ToString(intRow + 1), MsgBoxStyle.Information, ResolveResString(100))
                    Return False
                ElseIf Convert.ToInt32(strtoBox) <= 0 Then
                    MsgBox("To Box cannot be 0  at row no  - " & Convert.ToString(intRow + 1), MsgBoxStyle.Information, ResolveResString(100))
                    Return False
                End If

                Dim strSTockQty As Integer = SqlConnectionclass.ExecuteScalar("Select Cur_Bal From ItemBal_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Location_Code ='" & strStockLocation & "' and item_Code ='" & strItemCode & "'")

                If Convert.ToInt32(strQty) > strSTockQty Then
                    MsgBox("Quantity should not be Greater then Current Balance  at row no  - " & Convert.ToString(intRow + 1), MsgBoxStyle.Information, ResolveResString(100))
                    Return False
                End If
                For innerRow As Integer = 0 To dgvUpload.Rows.Count - 1
                    strInnerSalesOrderNo = dgvUpload.Rows(innerRow).Cells(enmUploadItem.SaleOrder).Value
                    strInnerAmendmentNo = dgvUpload.Rows(innerRow).Cells(enmUploadItem.AmendmentNo).Value
                    strInnerItemCode = dgvUpload.Rows(innerRow).Cells(enmUploadItem.ItemCode).Value
                    strInnerCustDrgNo = dgvUpload.Rows(innerRow).Cells(enmUploadItem.Cust_DrgwNo).Value
                    If intRow <> innerRow And strItemCode = strInnerItemCode And strCustDrgNo = strInnerCustDrgNo And strSalesOrderNo = strInnerSalesOrderNo And strAmendmentNo = strInnerAmendmentNo Then
                        MsgBox("Duplicate Item cannot be uploaded  at row no  - " & Convert.ToString(intRow + 1) & "  and row no -  " & Convert.ToString(innerRow + 1), MsgBoxStyle.Information, ResolveResString(100))
                        Return False
                    End If
                    If intRow <> innerRow And strItemCode = strInnerItemCode And strCustDrgNo = strInnerCustDrgNo And strSalesOrderNo = strInnerSalesOrderNo And strAmendmentNo <> strInnerAmendmentNo Then
                        MsgBox("Multiple Sale order No with different amendment no for same customer Item cannot be uploaded  at row no  - " & Convert.ToString(intRow + 1) & "  and row no -  " & Convert.ToString(innerRow + 1), MsgBoxStyle.Information, ResolveResString(100))
                        Return False
                    End If

                Next

                If ValidateScheduleQuantity(strItemCode, strCustDrgNo, strQty, intRow, strSalesOrderNo, strAmendmentNo) = False Then
                    Return False
                End If

                strQuery = "set dateformat 'dmy' SELECT * FROM DBO.UFN_GST_ITEMWISETAXES('" & gstrUNITID & "','" & txtcustomercode.Text.Trim & "','" & strItemCode & "','','')"
                dt = New DataTable
                dt = SqlConnectionclass.GetDataTable(strQuery)
                If dt.Rows.Count > 0 Then
                    dgvUpload.Rows(intRow).Cells(enmUploadItem.HSNCode).Value = Convert.ToString(dt.Rows(0)("HSNSACCODE"))
                    dgvUpload.Rows(intRow).Cells(enmUploadItem.CGSTTXRT_TYPE).Value = Convert.ToString(dt.Rows(0)("CGST_TXRT_HEAD"))
                    dgvUpload.Rows(intRow).Cells(enmUploadItem.SGSTTXRT_TYPE).Value = Convert.ToString(dt.Rows(0)("SGST_TXRT_HEAD"))
                    dgvUpload.Rows(intRow).Cells(enmUploadItem.UTGSTTXRT_TYPE).Value = Convert.ToString(dt.Rows(0)("UGST_TXRT_HEAD"))
                    dgvUpload.Rows(intRow).Cells(enmUploadItem.IGSTTXRT_TYPE).Value = Convert.ToString(dt.Rows(0)("IGST_TXRT_HEAD"))
                    Dim strHSN = dgvUpload.Rows(intRow).Cells(enmUploadItem.HSNCode).Value
                    If strHSN = "" Then
                        MsgBox("HSN/SAC CODE can't be blank for item codes " & strItemCode & "  at row no  - " & Convert.ToString(intRow + 1), MsgBoxStyle.Information, ResolveResString(100))
                        Return False
                    End If
                End If

            Next
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function ValidateScheduleQuantity(ByVal varItemCode As String, ByVal varDrgNo As String, ByVal varItemQty As String, ByVal intRwCount As Integer, ByRef strRefNo As String, ByRef strAmend As String) As Boolean
        '------------------------------------------------------------------
        'Name       :   Validate For Schedule Quantity
        'Type       :   Function
        'Author     :   Priti Sharma
        'Arguments  :
        'Return     :   True : If validation is Successfully Completed
        'Purpose    :
        '------------------------------------------------------------------
        Try
            Dim strInvoiceType As String
            Dim strInvoiceSubType As String
            Dim ldblNetDispatchQty As Double
            Dim blnDSTracking As Boolean
            Dim strMakeDate As String
            '*********************************************************
            'Validation For Schedule Start From Here
            '*********************************************************
            ValidateScheduleQuantity = True

            strInvoiceType = UCase(Trim(CmbInvType.Text))
            strInvoiceSubType = UCase(Trim(CmbInvSubType.Text))
            If ((UCase(Trim(strInvoiceType)) = "NORMAL INVOICE") And (UCase(Trim(strInvoiceSubType)) = "FINISHED GOODS")) Then
                blnDSTracking = SqlConnectionclass.ExecuteScalar("Select DSWiseTracking From Sales_parameter  (Nolock)  WHERE UNIT_CODE='" + gstrUNITID + "'")
                If blnDSTracking Then
                    If CheckMeasurmentUnit(varItemCode, varItemQty, intRwCount) = False Then
                        ValidateScheduleQuantity = False
                        Exit Function
                    End If
                    If CheckcustorddtlQty(CStr(varItemCode), CStr(varDrgNo), CDbl(varItemQty), strRefNo, strAmend) = True Then
                        ValidateScheduleQuantity = True
                    Else
                        ValidateScheduleQuantity = False
                        Exit Function
                    End If

                    ldblNetDispatchQty = GetTotalDispatchQuantityFromDailySchedule(Trim(txtcustomercode.Text), Trim(varDrgNo), Trim(varItemCode))
                    If Val(varItemQty) > Val(CStr(ldblNetDispatchQty)) Then
                        ValidateScheduleQuantity = False
                        MsgBox("Quantity should not be Greater than Schedule Quantity " & CStr(ldblNetDispatchQty) & " For Item Code " & varItemCode, MsgBoxStyle.Information, "eMPro")
                        Exit Function
                    Else
                        ValidateScheduleQuantity = True
                    End If
                End If
            End If
            ValidateScheduleQuantity = True
            Exit Function 'This is to avoid the execution of the error handler
        Catch ex As Exception
            RaiseException(ex)

        End Try
    End Function
    Private Function GetTotalDispatchQuantityFromDailySchedule(ByVal pstrAccountCode As String, ByVal pstrCustomerDrawNo As String, ByVal pstrItemCode As String) As Double
        Dim strScheduleSql As String
        Dim ldblTotalDispatchQuantity As Double
        Dim ldblTotalScheduleQuantity As Double
        Dim lintLoopCounter As Short
        Dim pstrDate As String
        Try
            ldblTotalDispatchQuantity = 0
            ldblTotalScheduleQuantity = 0
            With dtpDateDesc
                .Format = DateTimePickerFormat.Custom
                .CustomFormat = gstrDateFormat
                .Value = GetServerDate()
                .Visible = True 'Don't Show DatePicker
            End With
            pstrDate = getDateForDB(dtpDateDesc.Text)
            SqlConnectionclass.ExecuteNonQuery("SET DATEFORMAT 'mdy'")

            strScheduleSql = "Select isnull(Schedule_Quantity,0) Schedule_Quantity,isnull(Despatch_Qty,0) Despatch_Qty from DailyMktSchedule WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & pstrAccountCode & "' and "
            strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(CDate(pstrDate)) & "'"
            strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date)='" & Month(CDate(pstrDate)) & "'"
            strScheduleSql = strScheduleSql & " and Trans_Date <='" & pstrDate & "'"
            strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND ITEM_CODE  = '" & pstrItemCode & "' and Status =1  ORDER BY Trans_Date DESC"
            Dim dtItem As DataTable = SqlConnectionclass.GetDataTable(strScheduleSql)
            If dtItem.Rows.Count = 0 Then
                GetTotalDispatchQuantityFromDailySchedule = -1
                'mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                Exit Function
            Else
                For Each dr As DataRow In dtItem.Rows
                    ldblTotalScheduleQuantity = ldblTotalScheduleQuantity + Val(dr("Schedule_Quantity"))
                    ldblTotalDispatchQuantity = ldblTotalDispatchQuantity + Val(dr("Despatch_Qty"))
                Next
                GetTotalDispatchQuantityFromDailySchedule = Val(CStr(ldblTotalScheduleQuantity - ldblTotalDispatchQuantity))
                'mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                Exit Function
            End If

            Exit Function 'This is to avoid the execution of the error handler
        Catch ex As Exception
            GetTotalDispatchQuantityFromDailySchedule = -1
            RaiseException(ex)
        End Try
    End Function
    Public Function CheckcustorddtlQty(ByRef pstrItemCode As String, ByRef pstrDrgno As String, ByRef pdblQty As Double, ByRef strRefNo As String, ByRef strAmend As String) As Boolean
        Dim rsCustOrdDtl As New ClsResultSetDB
        Dim rssaledtl As ClsResultSetDB
        Dim dblSaleQuantity As Double
        Dim strCustOrdDtl As String
        Dim blnOpenSO As Boolean
        Dim dblBalanceqty As Double = 0
        Try
            strCustOrdDtl = "Select openso,balance_Qty = order_qty - Despatch_qty from Cust_ord_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
            strCustOrdDtl = strCustOrdDtl & "Account_code ='" & txtcustomercode.Text & "'" & " and Item_code ='"
            strCustOrdDtl = strCustOrdDtl & pstrItemCode & "' and cust_drgNo ='" & pstrDrgno
            strCustOrdDtl = strCustOrdDtl & "' and Authorized_flag = 1 and Active_Flag = 'A' and cust_ref = '" & strRefNo
            strCustOrdDtl = strCustOrdDtl & "' and Amendment_no = '" & strAmend & "'"
            Dim dt As DataTable = SqlConnectionclass.GetDataTable(strCustOrdDtl)
            If dt.Rows.Count > 0 Then
                blnOpenSO = Convert.ToString(dt.Rows(0)("openso"))
                dblBalanceqty = Convert.ToDouble(dt.Rows(0)("balance_Qty"))
                If blnOpenSO = True Then
                    CheckcustorddtlQty = True
                Else
                    If Val(dblBalanceqty) < Val(pdblQty) Then
                        MsgBox("Balance Quantity available in SO for Customer Part code [ " & pstrDrgno & "] is " & Val(dblBalanceqty) & ".", MsgBoxStyle.Information, "empower")
                        CheckcustorddtlQty = False
                    Else
                        CheckcustorddtlQty = True
                    End If
                End If
            End If
            Exit Function
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function
    Public Function CheckMeasurmentUnit(ByRef strItem As String, ByRef strQuantity As Double, ByRef intRow As Short) As Boolean
        Dim strMeasure As String
        Dim rsMeasure As ClsResultSetDB
        Dim blnDecimal_allowed_flag As Boolean = False
        Try
            blnDecimal_allowed_flag = SqlConnectionclass.ExecuteScalar("select a.Decimal_allowed_flag from Measure_Mst a,Item_Mst b  " & _
            "where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND b.cons_Measure_Code=a.Measure_Code and b.Item_Code = '" & strItem & "'")

            If blnDecimal_allowed_flag = False Then
                If System.Math.Round(strQuantity, 0) - Val(strQuantity) <> 0 Then
                    Call ConfirmWindow(10455, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    CheckMeasurmentUnit = False
                    Exit Function
                Else
                    CheckMeasurmentUnit = True
                End If
            Else
                CheckMeasurmentUnit = True
            End If

            Exit Function
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function

    Private Sub cmdRefNoHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefNoHelp.Click
        Dim frmMKTTRN0020NEW As New frmMKTTRN0020NEW
        Try
            If Len(txtcustomerhelp.Text) = 0 Then
                Call ConfirmWindow(10416, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                txtcustomercode.Focus()
                Exit Sub
            End If
            Dim strRefAmm As String
            Dim intPos As Short
            Dim mstrRefNo As String
            strRefAmm = frmMKTTRN0020NEW.SelectDataFromCustOrd_DtlUploadExcel(txtcustomerhelp.Text, "NORMAL INVOICE")
            If Len(strRefAmm) > 0 Then
                intPos = InStr(1, Trim(strRefAmm), ",", CompareMethod.Text)
                mstrRefNo = Mid(Trim(strRefAmm), 2, intPos - 3)
                txtRefNo.Text = Trim(mstrRefNo)
            End If

            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Private Function SelectInvoiceTypeFromSaleConf() As Object
        Dim strSaleConfSql As String
        Dim rsSaleConf As ClsResultSetDB
        Dim intRecCount As Short
        Dim intLoopCounter As Short
        Try
            strSaleConfSql = "Select Distinct(Description) from SaleConf (nolock) WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type Not in('STX') and (fin_start_date <= getdate() and fin_end_date >= getdate()) "
            rsSaleConf = New ClsResultSetDB
            rsSaleConf.GetResult(strSaleConfSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsSaleConf.GetNoRows > 0 Then
                intRecCount = rsSaleConf.GetNoRows
                rsSaleConf.MoveFirst()
                For intLoopCounter = 0 To intRecCount - 1
                    VB6.SetItemString(CmbInvType, intLoopCounter, rsSaleConf.GetValue("Description"))
                    rsSaleConf.MoveNext()
                Next intLoopCounter
            End If
            rsSaleConf.ResultSetClose()
            rsSaleConf = Nothing
            Exit Function
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub SelectInvoiceSubTypeFromSaleConf(ByRef pstrInvType As String)
        '****************************************************
        'Description    -  Select Invoice SubTypeDescription From SaleConf Acc. to Inv. Type
        '****************************************************

        Dim strSaleConfSql As String
        Dim rsSaleConf As ClsResultSetDB
        Dim intRecCount As Short
        Dim intLoopCounter As Short
        Try
            strSaleConfSql = "Select Distinct(Sub_Type_Description) from SaleConf (nolock)  WHERE UNIT_CODE='" + gstrUNITID + "' AND  Description='" & Trim(pstrInvType) & "' and  (fin_start_date <= getdate() and fin_end_date >= getdate()) "
            rsSaleConf = New ClsResultSetDB
            rsSaleConf.GetResult(strSaleConfSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsSaleConf.GetNoRows > 0 Then
                intRecCount = rsSaleConf.GetNoRows
                rsSaleConf.MoveFirst()
                CmbInvSubType.Items.Clear()
                For intLoopCounter = 0 To intRecCount - 1
                    VB6.SetItemString(CmbInvSubType, intLoopCounter, rsSaleConf.GetValue("Sub_Type_Description"))
                    rsSaleConf.MoveNext()
                Next intLoopCounter
                CmbInvSubType.SelectedIndex = 0
            End If
            rsSaleConf.ResultSetClose()
            rsSaleConf = Nothing
            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
  
    Private Sub cmdVehicleCodeHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdVehicleCodeHelp.Click
        Dim strSql As String = ""
        Dim strVehicle As String = ""
        Dim varRetVal As Object
        Try
            With txtVehNo
                If Len(.Text) = 0 Then
                    varRetVal = ShowList(1, .MaxLength, "", "vehicle_no", "Transporter_name", "vehicle_mst", " and active=1 ", "Help", "", "", 0, "transporter_code")
                    If varRetVal = "-1" Then
                        Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        .Text = ""
                    Else
                        .Text = varRetVal

                    End If
                Else
                    varRetVal = ShowList(1, .MaxLength, , "vehicle_no", "Transporter_name", "vehicle_mst", " and active=1 ", "Help", "", "", 0, "transporter_code")
                    If varRetVal = "-1" Then
                        Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        .Text = ""
                    Else
                        .Text = varRetVal
                    End If
                End If
                .Focus()
            End With

            Exit Sub
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Try
            ResetData()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class