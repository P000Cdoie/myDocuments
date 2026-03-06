Option Strict Off
Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
'----------------------------------------------------
'Copyright (c)  -  MIND
'Name of module -  frmMKTTRN0063.vb
'Created By     -  
'Created Date   -  
'Description    -
'Revised date   -
'-----------------------------------------------------------------------------
'Modified by    :   Virendra Gupta
'Modified ON    :   26/05/2011
'Modified to support MultiUnit functionality
'-----------------------------------------------------------------------
'Revised By         :   Shubhra Verma
'Revised On         :   23 Dec 2011
'Reason             :   Arithmetic Overflow error
'issue id           :   10174369
'============================================================================================
'MODIFIED BY NITIN MEHTA 0N 31 JAN 2012 FOR CHANGE MANAGEMENT
'============================================================================================
'Revised By         :   PRASHANT RAJPAL
'Revised On         :   14 aug 2012
'Reason             :   Changes in Forecast uploading 
'issue id           :   10261997 
'============================================================================================
'Revised By         :   Vinod Singh
'Revised On         :   26 Oct 2012
'Reason             :   SafeFileName method is replaced as this method is only available 
'                       in Framework 2.0 SP1 or later
'Issue id           :   10276584 
'============================================================================================
'Revised By         :   Vinod Singh
'Revised On         :   05 March 2013
'Reason             :   Error in uploading excel file.
'Issue id           :   10352802 
'============================================================================================
'Revised By         :   Shubhra Verma
'Revised On         :   11 March 2013
'Reason             :   not generating files for excel 2003
'Issue id           :   10354980
'============================================================================================
'Revised By         :   Shubhra Verma
'Revised On         :   16 Apr 2013
'Reason             :   Excel 2007 files not uploading
'issue id           :   10309901
'============================================================================================
'Revised By         :   Milind Mishra
'Revised On         :   26 May 2017
'Reason             :   Forecast Uploading Error
'issue id           :   101289013 
'============================================================================================
Friend Class frmMKTTRN0121
    Inherits System.Windows.Forms.Form
    Dim strUploadedFileName As String
    Dim mintIndex As Short
    Dim mflag As Short
    Private Enum enmDSdetail
        'ItemCode = 1
        'Item_Description
        'Cust_DrgwNo
        'UOM
        'Quantity
        'DELIVERY_DATE
        CustomerCode = 0
        ItemCode
        ItemDesc
        Cust_DrgwNo
        Trans_date
        ScheduleQty
        DispatchQty
        BalanceQty
        Serial_No
        DSNO
        Product_Start_date
        Product_End_date
    End Enum
    Private Enum enmDSitem
        'status = 1
        'ItemCode
        'Item_Description
        'Cust_DrgwNo
        status = 1
        CustomerCode
        ItemCode
        ItemDesc
        Cust_DrgwNo
        Trans_date
        ScheduleQty
        DispatchQty
        BalanceQty
        Serial_No
        DSNO

    End Enum
    Private Sub SetSpreadProperty()
        On Error GoTo Errorhandler
        With fspforecastdetail
            .MaxRows = 0
            .MaxCols = 0
            .set_RowHeight(0, 400)
            .Row = 0
            .ColsFrozen = 0
            .MaxCols = enmDSdetail.Product_End_date
            .Col = enmDSdetail.CustomerCode : .Text = "Cutomer Code" : .set_ColWidth(enmDSdetail.CustomerCode, 1600)
            .Col = enmDSdetail.ItemCode : .Text = "Item Code" : .set_ColWidth(enmDSdetail.ItemCode, 1000)
            .Col = enmDSdetail.ItemDesc : .Text = "Item Desc" : .set_ColWidth(enmDSdetail.ItemDesc, 2000)
            .Col = enmDSdetail.Cust_DrgwNo : .Text = "Cust Drgn No" : .set_ColWidth(enmDSdetail.Cust_DrgwNo, 2000)
            .Col = enmDSdetail.Trans_date : .Text = "Trans Date" : .set_ColWidth(enmDSdetail.Trans_date, 1200)
            .Col = enmDSdetail.ScheduleQty : .Text = "Schedule Qty" : .set_ColWidth(enmDSdetail.ScheduleQty, 1000) : .ColHidden = True
            .Col = enmDSdetail.DispatchQty : .Text = "Dispatch Qty" : .set_ColWidth(enmDSdetail.DispatchQty, 1000) : .ColHidden = True
            .Col = enmDSdetail.BalanceQty : .Text = "Balance Qty" : .set_ColWidth(enmDSdetail.BalanceQty, 1000)
            .Col = enmDSdetail.Serial_No : .Text = "Serial No" : .set_ColWidth(enmDSdetail.Serial_No, 1000) : .ColHidden = True
            .Col = enmDSdetail.DSNO : .Text = "DS No" : .set_ColWidth(enmDSdetail.DSNO, 1000) : .ColHidden = True
            .Col = enmDSdetail.Product_Start_date : .Text = "Start Date" : .set_ColWidth(enmDSdetail.Product_Start_date, 1200)
            .Col = enmDSdetail.Product_End_date : .Text = "End Date" : .set_ColWidth(enmDSdetail.Product_End_date, 1200)

        End With
        fspforecastdetail.ReDraw = True
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SetSpreadColTypes(ByRef pintRowNo As Long)
        On Error GoTo ErrHandler
        Dim i As Long
        With Me.fspforecastdetail
            .Row = pintRowNo

            .Col = enmDSdetail.CustomerCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmDSdetail.ItemCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmDSdetail.ItemDesc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmDSdetail.Cust_DrgwNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmDSdetail.Trans_date : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmDSdetail.ScheduleQty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmDSdetail.DispatchQty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmDSdetail.BalanceQty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmDSdetail.Serial_No : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmDSdetail.DSNO : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmDSdetail.Product_Start_date : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmDSdetail.Product_End_date : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
        End With
        fspforecastdetail.ReDraw = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub addRowAtEnterKeyPress(ByRef pintRows As Long)
        On Error GoTo ErrHandler
        Dim intRowHeight As Long
        With Me.fspforecastdetail
            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            For intRowHeight = 1 To pintRows
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .set_RowHeight(.Row, 300)
                Call SetSpreadColTypes(.Row)
            Next intRowHeight
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub RefreshForm()
        On Error GoTo ErrHandler
        fspforecastdetail.MaxCols = 0
        fspforecastdetail.MaxRows = 0
        txtcustomercode.Text = ""
        lblcustomername.Text = ""
        txtfilelocation.Text = ""
        TxtSearch.Text = ""
        CmbSearchBy.SelectedIndex = -1
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub chkall_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkall.CheckStateChanged
        On Error GoTo ErrHandler
        Dim intcounter As Long
        Dim varstatus As Object
        With fspitem
            For intcounter = 1 To .MaxRows
                If chkall.CheckState = 1 Then
                    Call .SetText(enmDSitem.status, intcounter, "1")
                Else
                    Call .SetText(enmDSitem.status, intcounter, "0")
                End If
            Next
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub CmbSearchBy_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbSearchBy.TextChanged
        On Error GoTo ErrHandler
        If TxtSearch.Enabled Then
            TxtSearch.Focus()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub CmbSearchBy_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbSearchBy.SelectedIndexChanged
        On Error GoTo ErrHandler
        If TxtSearch.Enabled Then
            TxtSearch.Focus()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdCustHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdcusthelp.Click
        On Error GoTo ErrHandler
        Dim strCustHelp() As String
        Dim strsql As String
        If Len(Me.txtcustomerhelp.Text) = 0 Then
            '    strCustHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct customer_Code, cust_name as customer_name from DailyMktSchedule a ( nolock) ,customer_mst b(nolock) ,Item_mst c (Nolock) " & _
            '"where  a.UNIT_CODE=b.UNIT_CODE and a.Account_Code=b.Customer_Code  and a.unit_code=c.unit_code and " & _
            '"a.item_code=c.item_code and a.unit_code='" & gstrUNITID & "' and c.Status ='A' and c.Hold_Flag <> 1  and isnull(manual_ds_closure,0)=0  and a.Trans_date between '" & VB6.Format(Me.dtfromdate.Value, "dd/mmm/yyyy") & "' and '" & VB6.Format(Me.dttodate.Value, "dd/mmm/yyyy") & "'", "Customer Codes List", 1)
            strCustHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct customer_Code, cust_name as customer_name from customer_mst b(nolock)  " & _
      "where b.unit_code='" & gstrUNITID & "'  and isnull(manual_ds_closure,0)=0  and ((isnull(b.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= b.deactive_date)) ", "Customer Codes List", 1)
        Else
            strCustHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select a.customer_Code, a.cust_name  as customer_name from Customer_Mst a where a.customer_code='" & Trim(txtcustomerhelp.Text) & "' and isnull(manual_ds_closure,0)=0 and a.Unit_code = '" & gstrUNITID & "' and ((isnull(a.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= a.deactive_date))", "Customer Codes List", 1)
        End If
        If UBound(strCustHelp) <> -1 Then
            If strCustHelp(0) <> "0" Then
                txtcustomerhelp.Text = Trim(strCustHelp(0))
                Call txtcustomerhelp_Validating(txtcustomerhelp, New System.ComponentModel.CancelEventArgs(False))
                lblcustname.Text = strCustHelp(1)

                InsertCustomerItem(txtcustomerhelp.Text)

               
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
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub InsertCustomerItem(ByVal strCustomer As String)
        On Error GoTo ErrHandler
        Dim strsql As String = ""
        mP_Connection.Execute("Delete DSCustomerCode Where IP_ADDRESS = '" & gstrIpaddressWinSck & "' AND UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.Execute("Delete DSItemCode Where IP_ADDRESS = '" & gstrIpaddressWinSck & "' AND UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

        strsql = "insert into DSCustomerCode (Account_Code,UNIT_CODE,IP_ADDRESS) Values('" & strCustomer & "','" & gstrUNITID & "','" & gstrIpaddressWinSck & "')"
        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

        strsql = "select Distinct(a.Item_code) as Item_code,a.Unit_code,'" & gstrIpaddressWinSck & "' from DailyMktSchedule a ( nolock) ,customer_mst b(nolock) ,Item_mst c (Nolock) " & _
       "where a.Account_code='" & strCustomer & "' and a.UNIT_CODE=b.UNIT_CODE and a.Account_Code=b.Customer_Code  and a.unit_code=c.unit_code and " & _
       "a.item_code=c.item_code and a.unit_code='" & gstrUNITID & "' and c.Status ='A' and c.Hold_Flag <> 1  " & _
       "and isnull(manual_ds_closure,0)=0  and a.Trans_date between '" & VB6.Format(Me.dtfromdate.Value, "dd/mmm/yyyy") & "' and '" & VB6.Format(Me.dttodate.Value, "dd/mmm/yyyy") & "' "
        mP_Connection.Execute("insert into DSItemCode (Item_Code,UNIT_CODE,IP_ADDRESS) " & strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub
    Private Sub cmdfilelocation_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdfilelocation.Click
        On Error GoTo ErrHandler
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
                fspforecastdetail.MaxRows = 0
                fspforecastdetail.MaxRows = 0
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdgenerateformat_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdgenerateformat.Click
        On Error GoTo ErrHandler
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
        End If
        'If validate_daterange() = False Then
        '    MsgBox("Date Range Cannot Be More Than 24 Months ( 2 years) ", MsgBoxStyle.Information, ResolveResString(100))
        '    Exit Sub
        'End If
        If validate_calender() = False Then
            MsgBox("Sales Calender Is Not Defined For Specified Date Range", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If
        If fspitem.MaxRows > 0 And validate_itemselected() = True Then
            If objfso.FolderExists(gstrLocalCDrive & "DSUpload") = True Then
                objfolder = objfso.GetFolder(gstrLocalCDrive & "DSUpload")
                For Each objfile In objfolder.Files
                    stroldfilename = objfile.Name
                    If UCase(Mid(stroldfilename, 10, 8)) = UCase(txtcustomerhelp.Text) Then
                        If objfso.FolderExists(gstrLocalCDrive & "DSUpload\Backup") = True Then
                            If stroldfilename <> "" Then
                                objfso.MoveFile(gstrLocalCDrive & "DSUpload\" & stroldfilename, gstrLocalCDrive & "DSUpload\backup\")
                            End If
                        Else
                            objfso.CreateFolder(gstrLocalCDrive & "DSUpload\backup")
                            If stroldfilename <> "" Then
                                objfso.MoveFile(gstrLocalCDrive & "DSUpload\" & stroldfilename, gstrLocalCDrive & "DSUpload\backup\")
                            End If
                        End If
                    End If
                Next objfile
                OBj_wB = OBJexlformat.Workbooks.Add
                strnewfilename = getfilename(Trim(txtcustomerhelp.Text))
                OBj_wB.SaveAs(gstrLocalCDrive & "DSUpload\" & Replace(strnewfilename, ":", ""))
                OBJexlformat.Workbooks.Close()
                OBj_wB = Nothing
                OBJexlformat.Workbooks.Open(gstrLocalCDrive & "DSUpload\" & Replace(strnewfilename, ":", ""))
                OBJexlformat.Cells._Default(1, 1) = "DS UPLOADER"
                OBJexlformat.Cells._Default(2, 1) = "Customer Code"
                OBJexlformat.Cells._Default(2, 1).ColumnWidth = 15
                OBJexlformat.Cells._Default(2, 1).Font.Bold = True
                OBJexlformat.Cells._Default(2, 2) = txtcustomerhelp.Text
                OBJexlformat.Cells._Default(2, 2).Font.Bold = True
                OBJexlformat.Cells._Default(3, 1) = "Customer Name"
                OBJexlformat.Cells._Default(3, 1).ColumnWidth = 15
                OBJexlformat.Cells._Default(3, 1).Font.Bold = True
                OBJexlformat.Cells._Default(3, 2) = lblcustname.Text
                OBJexlformat.Cells._Default(3, 2).Font.Bold = True
                OBJexlformat.Cells._Default(5, 1) = "Date Range"
                OBJexlformat.Cells._Default(5, 1).Font.Bold = True
                OBJexlformat.Cells._Default(5, 2) = dtfromdate.Value
                OBJexlformat.Cells._Default(5, 3) = "To"
                OBJexlformat.Cells._Default(5, 4) = dttodate.Value
                OBJexlformat.Cells._Default(6, 1) = "Item Code"
                OBJexlformat.Cells._Default(6, 1).Font.Bold = True
                OBJexlformat.Cells._Default(6, 2) = "Item Desc"
                OBJexlformat.Cells._Default(6, 2).Font.Bold = True
                OBJexlformat.Cells._Default(6, 2) = "Item Desc"
                OBJexlformat.Cells._Default(6, 2).Font.Bold = True
                OBJexlformat.Cells._Default(6, 3) = "Cust DrgNo"
                OBJexlformat.Cells._Default(6, 3).Font.Bold = True
                OBJexlformat.Cells._Default(6, 4) = "Trans Date"
                OBJexlformat.Cells._Default(6, 4).Font.Bold = True
                OBJexlformat.Cells._Default(6, 5) = "Balance Qty"
                OBJexlformat.Cells._Default(6, 5).Font.Bold = True
                'OBJexlformat.Cells(3).numberformat = "@"
                OBJexlformat.Cells(1).columns.numberformat = "@"
                OBJexlformat.Cells(1).entirecolumn.numberformat = "@"

                OBJexlformat.Cells(3).columns.numberformat = "@"
                OBJexlformat.Cells(3).entirecolumn.numberformat = "@"

                Call fillitems(OBJexlformat)
                OBJexlformat.ActiveWorkbook.Save()
                OBJexlformat.Quit()

                MsgBox("Excel Sheet Generated Successfully " & vbCrLf & "location--> " & " " & gstrLocalCDrive & "DSUpload\" & Replace(strnewfilename, ":", ""), MsgBoxStyle.Information, ResolveResString(100))
                clear_fields()
                chkall.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkall.Enabled = False
                OBJexlformat = Nothing
            Else
                objfso.CreateFolder(gstrLocalCDrive & "DSUpload")
                OBj_wB = OBJexlformat.Workbooks.Add
                strnewfilename = getfilename(Trim(txtcustomerhelp.Text))
                OBj_wB.SaveAs(gstrLocalCDrive & "DSUpload\" & Replace(strnewfilename, ":", ""))
                OBJexlformat.Workbooks.Close()
                OBj_wB = Nothing
                OBJexlformat.Workbooks.Open(gstrLocalCDrive & "DSUpload\" & Replace(strnewfilename, ":", ""))
                OBJexlformat.Cells._Default(2, 1) = "Customer Code"
                OBJexlformat.Cells._Default(2, 1).ColumnWidth = 15
                OBJexlformat.Cells._Default(2, 1).Font.Bold = True
                OBJexlformat.Cells._Default(2, 2) = txtcustomerhelp.Text
                OBJexlformat.Cells._Default(2, 2).Font.Bold = True
                OBJexlformat.Cells._Default(3, 1) = "Customer Name"
                OBJexlformat.Cells._Default(3, 1).ColumnWidth = 15
                OBJexlformat.Cells._Default(3, 1).Font.Bold = True
                OBJexlformat.Cells._Default(3, 2) = lblcustname.Text
                OBJexlformat.Cells._Default(3, 2).Font.Bold = True
                OBJexlformat.Cells._Default(5, 1) = "Date Range"
                OBJexlformat.Cells._Default(5, 1).Font.Bold = True
                OBJexlformat.Cells._Default(5, 2) = dtfromdate.Value
                OBJexlformat.Cells._Default(5, 3) = "To"
                OBJexlformat.Cells._Default(5, 3) = dttodate.Value
                OBJexlformat.Cells._Default(6, 1) = "Item Code"
                OBJexlformat.Cells._Default(6, 1).Font.Bold = True
                OBJexlformat.Cells._Default(6, 2) = "Item Desc"
                OBJexlformat.Cells._Default(6, 2).Font.Bold = True
                OBJexlformat.Cells._Default(6, 2) = "Item Desc"
                OBJexlformat.Cells._Default(6, 2).Font.Bold = True
                OBJexlformat.Cells._Default(6, 3) = "Cust DrgNo"
                OBJexlformat.Cells._Default(6, 3).Font.Bold = True
                OBJexlformat.Cells._Default(6, 4) = "Trans Date"
                OBJexlformat.Cells._Default(6, 4).Font.Bold = True
                OBJexlformat.Cells._Default(6, 5) = "Balance Qty"
                OBJexlformat.Cells._Default(6, 5).Font.Bold = True
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
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdclose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdclose.Click
        On Error GoTo ErrHandler
        Me.Close()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdupload_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdupload.Click
        On Error GoTo ErrHandler
        If txtcustomercode.Text <> "" Then
            If txtfilelocation.Text <> "" Then
                fspforecastdetail.MaxRows = 0
                fspforecastdetail.MaxRows = 0
                cmdsave.Enabled = True
                Call upload_DS_data()
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
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdviewitems_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdviewitems.Click
        On Error GoTo ErrHandler
        Dim objgetitems As New ClsResultSetDB
        Dim strSQL As String
        Dim sqlCmd As New SqlCommand
        If Len(Trim(txtcustomerhelp.Text)) = 0 Then
            MsgBox("Please Select Customer Code First", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If
        InsertCustomerItem(txtcustomerhelp.Text)
        With sqlCmd
            .CommandType = CommandType.StoredProcedure
            .CommandTimeout = 300 ' 5 Minute
            .CommandText = "DS_CLOSURE"
            .Parameters.Clear()
            .Parameters.AddWithValue("@unitcode", gstrUNITID)
            .Parameters.AddWithValue("@DateFrom", VB6.Format("01 Jan 2022", "dd/mmm/yyyy"))
            .Parameters.AddWithValue("@DateTo", VB6.Format(Me.dttodate.Value, "dd/mmm/yyyy"))
            .Parameters.AddWithValue("@IP_ADDRESS", gstrIpaddressWinSck)
            .Parameters.AddWithValue("@User_Id", mP_User)
            .Parameters.AddWithValue("@TYPE", "UPLOAD")
            SqlConnectionclass.ExecuteNonQuery(sqlCmd)

            strSQL = "SELECT * FROM TEMP_Dailymktschedule_CLOSURE WHERE UNIT_CODE ='" & gstrUNITID & "' AND IP_ADDRESS ='" & gstrIpaddressWinSck & "' and type='UPLOAD'"
            objgetitems.GetResult(strSQL)
            If Not objgetitems.EOFRecord Then
                Call setitemgrdproperty()
                While Not objgetitems.EOFRecord
                    With fspitem
                        addrowinitemgrd((1))
                        Call .SetText(enmDSitem.CustomerCode, .MaxRows, objgetitems.GetValue("Account_Code"))
                        Call .SetText(enmDSitem.ItemCode, .MaxRows, objgetitems.GetValue("Item_Code"))
                        Call .SetText(enmDSitem.ItemDesc, .MaxRows, objgetitems.GetValue("Item_Desc"))
                        Call .SetText(enmDSitem.Cust_DrgwNo, .MaxRows, objgetitems.GetValue("Cust_Drgno"))
                        Call .SetText(enmDSitem.Trans_date, .MaxRows, objgetitems.GetValue("Trans_date"))
                        Call .SetText(enmDSitem.ScheduleQty, .MaxRows, objgetitems.GetValue("Schedule_Quantity"))
                        Call .SetText(enmDSitem.DispatchQty, .MaxRows, objgetitems.GetValue("Despatch_Qty"))
                        Call .SetText(enmDSitem.BalanceQty, .MaxRows, objgetitems.GetValue("Schedule_Quantity") - objgetitems.GetValue("Despatch_Qty"))
                        Call .SetText(enmDSitem.Serial_No, .MaxRows, objgetitems.GetValue("Serial_No"))
                        Call .SetText(enmDSitem.DSNO, .MaxRows, objgetitems.GetValue("DSNO"))
                        objgetitems.MoveNext()
                    End With
                End While
                chkall.Enabled = True
                cmdgenerateformat.Enabled = True
            Else
                MsgBox("No Item Is Defined For The Specified Customer", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
        End With

ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0063_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0063_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0063_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlDSWiseScheduleStatus.Tag)
        FitToClient(Me, framain, ctlDSWiseScheduleStatus, frabuttons)
        Call FillLabelFromResFile(Me)
        cmdcustomercode.Image = My.Resources.ico111.ToBitmap
        ResetData()
        SSTab1.SelectedIndex = 0
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ResetData()
        fspforecastdetail.MaxCols = 0
        fspforecastdetail.MaxRows = 0
        txtfilelocation.Enabled = False
        dtfromdate.Format = DateTimePickerFormat.Custom
        dtfromdate.CustomFormat = gstrDateFormat
        dtfromdate.Value = GetServerDate()
        dttodate.Format = DateTimePickerFormat.Custom
        dttodate.CustomFormat = gstrDateFormat
        dttodate.Value = GetServerDate()
        txtfilelocation.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        chkall.Enabled = False
        cmdgenerateformat.Enabled = False
        cmdsave.Enabled = False
    End Sub
    Private Sub frmMKTTRN0063_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdcustomercode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdcustomercode.Click
        On Error GoTo ErrHandler
        Dim strCustHelp() As String
        If Len(Me.txtcustomercode.Text) = 0 Then
            '    strCustHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct customer_Code, cust_name as customer_name from  DailyMktSchedule a ( nolock) ,customer_mst b(nolock) ,Item_mst c (Nolock) " & _
            '"where  a.UNIT_CODE=b.UNIT_CODE and a.Account_Code=b.Customer_Code  and a.unit_code=c.unit_code and " & _
            '"a.item_code=c.item_code and a.unit_code='" & gstrUNITID & "' and c.Status ='A' and c.Hold_Flag <> 1  and isnull(manual_ds_closure,0)=0  and a.Trans_date between '" & VB6.Format(Me.dtfromdate.Value, "dd/mmm/yyyy") & "' and '" & VB6.Format(Me.dttodate.Value, "dd/mmm/yyyy") & "'", "Customer Codes List", 1)

            strCustHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct customer_Code, cust_name as customer_name from customer_mst b(nolock)  " & _
     "where b.unit_code='" & gstrUNITID & "'  and isnull(manual_ds_closure,0)=0  and ((isnull(b.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= b.deactive_date)) ", "Customer Codes List", 1)
        Else
            strCustHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select a.customer_Code, a.cust_name  as customer_name from Customer_Mst a where a.customer_code='" & Trim(txtcustomercode.Text) & "' and a.Unit_code = '" & gstrUNITID & "' and isnull(manual_ds_closure,0)=0 and ((isnull(a.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= a.deactive_date))", "Customer Codes List", 1)
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
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SSTab1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SSTab1.SelectedIndexChanged
        On Error GoTo ErrHandler
        If SSTab1.SelectedIndex = 0 Then
            cmdsave.Enabled = False
        Else
            cmdsave.Enabled = True
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustomerCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtcustomercode.TextChanged
        On Error GoTo ErrHandler
        lblcustomername.Text = ""
        fspforecastdetail.MaxCols = 0
        fspforecastdetail.MaxRows = 0
        txtfilelocation.Text = ""
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustomerCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtcustomercode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call cmdcustomercode_Click(cmdcustomercode, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
    Private Sub fillGrid()
        On Error GoTo ErrHandler
        Dim objgetDetail As New ClsResultSetDB
        Dim strSQL As String
        Dim intRowCount As Long
        Dim intcounter As Long
        Dim intMaxCounter As Long
        strSQL = "select product_no,due_date,sum(quantity)as quantity from forecast_mst_temp where customer_code='" & Trim(txtcustomercode.Text) & "' and ipaddress='" & gstrIpaddressWinSck & "' and Unit_code = '" & gstrUNITID & "' and due_date>=convert(varchar(12),getdate(),106) group by product_no,due_date"
        objgetDetail.GetResult(strSQL)
        With fspforecastdetail
            If Not objgetDetail.EOFRecord Then
                intRowCount = objgetDetail.GetNoRows
                Call addRowAtEnterKeyPress(intRowCount)
                objgetDetail.MoveFirst()
                For intcounter = 1 To intRowCount
                    .Row = intcounter
                    Call .SetText(enmDSitem.CustomerCode, intcounter, objgetDetail.GetValue("Account_Code"))
                    Call .SetText(enmDSitem.ItemCode, intcounter, objgetDetail.GetValue("Item_Code"))
                    Call .SetText(enmDSitem.ItemDesc, intcounter, objgetDetail.GetValue("Item_Desc"))
                    Call .SetText(enmDSitem.Trans_date, intcounter, VB6.Format(objgetDetail.GetValue("due_date"), gstrDateFormat))
                    Call .SetText(enmDSitem.ScheduleQty, intcounter, objgetDetail.GetValue("ScheduleQty"))
                    Call .SetText(enmDSitem.DispatchQty, intcounter, objgetDetail.GetValue("DispatchQty"))
                    'Call .SetText(enmDSitem.BalanceQty, intcounter, objgetDetail.GetValue("ScheduleQty") - objgetDetail.GetValue("DispatchQty"))
                    Call .SetText(enmDSitem.Serial_No, intcounter, objgetDetail.GetValue("Serial_No"))
                    Call .SetText(enmDSitem.DSNO, intcounter, objgetDetail.GetValue("DSNO"))

                    objgetDetail.MoveNext()
                Next
            Else
                MsgBox("No Data To Upload", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End With
        objgetDetail = Nothing
        Exit Sub
ErrHandler:
        objgetDetail = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SearchItem()
        On Error GoTo ErrHandler
        Dim intRowNo As Long
        Dim strSearchBy As Object
        Dim strSearchItem As Object
        Dim intCount As Long
        If Len(Trim(CmbSearchBy.Text)) = 0 Then
            If CmbSearchBy.Enabled Then CmbSearchBy.Focus()
            Exit Sub
        End If
        If Len(Trim(TxtSearch.Text)) = 0 Then
            fspitem.TopRow = 1
            fspitem.Font = VB6.FontChangeBold(fspitem.Font, False)
            If TxtSearch.Enabled Then TxtSearch.Focus()
            Exit Sub
        End If
        strSearchBy = CmbSearchBy.Text
        For intCount = 0 To fspitem.MaxRows - 1
            fspitem.Row = intCount
            fspitem.Col = enmDSitem.ItemDesc
            fspitem.Font = VB6.FontChangeBold(fspitem.Font, False)
            fspitem.Col = enmDSitem.ItemCode
            fspitem.Font = VB6.FontChangeBold(fspitem.Font, False)
        Next
        For intRowNo = 0 To fspitem.MaxRows - 1
            strSearchItem = Nothing
            If strSearchBy = "Item Code" Then
                Call fspitem.GetText(enmDSitem.ItemCode, intRowNo, strSearchItem)
            ElseIf strSearchBy = "Item Description" Then
                Call fspitem.GetText(enmDSitem.ItemDesc, intRowNo, strSearchItem)
            End If
            If UCase(strSearchItem) Like UCase(TxtSearch.Text) & "*" Then
                If strSearchBy = "Item Description" Then
                    fspitem.Col = enmDSitem.ItemDesc
                ElseIf strSearchBy = "Item Code" Then
                    fspitem.Col = enmDSitem.ItemCode
                End If
                fspitem.Row = intRowNo
                fspitem.TopRow = intRowNo
                fspitem.Row = intRowNo
                fspitem.Font = VB6.FontChangeBold(fspitem.Font, True)
                Exit For
            End If
        Next intRowNo
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtcustomerhelp_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtcustomerhelp.TextChanged
        On Error GoTo ErrHandler
        lblcustname.Text = ""
        fspitem.MaxCols = 0
        fspitem.MaxRows = 0
        chkall.Enabled = False
        cmdgenerateformat.Enabled = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
            strSQL = "select a.customer_Code, a.cust_name from Customer_Mst a where a.customer_code='" & Trim(txtcustomerhelp.Text) & "' and isnull(manual_ds_closure,0)=0  and a.Unit_code = '" & gstrUNITID & "' and ((isnull(a.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= a.deactive_date))"
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
        Dim intcounter As Long
        Dim varItemCode As Object
        Dim varItemDesc As Object
        Dim VarCustDrgNo As Object
        Dim varquantity As Object
        Dim varTransDate As Object
        Dim StrMessage As String
        Dim objfso As New Scripting.FileSystemObject

        Try

            strInsert = "Delete from TEMP_Dailymktschedule_UPLOADER where unit_code='" & gstrUNITID & "'  and IP_ADDRESS='" & gstrIpaddressWinSck & "'"
            SqlConnectionclass.ExecuteNonQuery(strInsert)

            With fspforecastdetail

                For intcounter = 1 To fspforecastdetail.MaxRows
                    varItemCode = Nothing
                    varItemDesc = Nothing
                    varTransDate = Nothing
                    varquantity = Nothing
                    VarCustDrgNo = Nothing
                    Call .GetText(enmDSdetail.ItemCode, intcounter, varItemCode)
                    Call .GetText(enmDSdetail.ItemDesc, intcounter, varItemDesc)
                    Call .GetText(enmDSdetail.Cust_DrgwNo, intcounter, VarCustDrgNo)
                    Call .GetText(enmDSdetail.Trans_date, intcounter, varTransDate)
                    Call .GetText(enmDSdetail.BalanceQty, intcounter, varquantity)

                    strInsert = "insert into TEMP_Dailymktschedule_UPLOADER(Account_Code,Trans_date,Item_code,Item_desc,Cust_Drgno,Balance_Qty,Unit_Code,ent_userid,ent_dt,upd_userid,upd_dt,IP_ADDRESS)" & " values('" & txtcustomercode.Text & "','" & getDateForDB(varTransDate) & "','" & varItemCode & "','" & varItemDesc & "','" & VarCustDrgNo & "'," & varquantity & ",'" & gstrUNITID & "','" & mP_User & "',getdate(),'" & mP_User & "',getdate(),'" & gstrIpaddressWinSck & "')"
                    SqlConnectionclass.ExecuteNonQuery(strInsert)
                Next
            End With
            Dim Intcount As Integer = SqlConnectionclass.ExecuteScalar("Select isnull(count(*),0 )  from TEMP_Dailymktschedule_UPLOADER a,CustItem_Mst b  " & _
                   "where a.Account_Code=b.Account_Code and a.Item_Code=b.item_code and a.Cust_Drgno=b.Cust_Drgno and a.UNIT_CODE=b.UNIT_CODE  and  " & _
                   "trans_date not between  b.Product_Start_date and b.Product_End_date and a.unit_code='" & gstrUNITID & "'  and IP_ADDRESS='" & gstrIpaddressWinSck & "'")
            If Intcount > 0 Then
                MsgBox("Schedule data is not between start date and end date for some items, Kind refer to data highlighted in GRID.Kindly correct the data and upload sheet again .")
                FillWrongData()
                Exit Function
            End If

            SqlConnectionclass.BeginTrans()

            Using sqlCmd As SqlCommand = New SqlCommand
                With sqlCmd
                    .CommandText = "USP_SAVE_DSUPLOADER"
                    .CommandTimeout = 0
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                    .Parameters.AddWithValue("@TYPE", "UPLOADER")
                    .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 2000).Direction = ParameterDirection.Output
                    SqlConnectionclass.ExecuteNonQuery(sqlCmd)

                    If Convert.ToString(.Parameters("@MESSAGE").Value).Length > 0 Then
                        SqlConnectionclass.RollbackTran()
                        MessageBox.Show(Convert.ToString(.Parameters("@MESSAGE").Value), "eMPro", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                        Exit Function
                    Else
                        If objfso.FolderExists(gstrLocalCDrive & "DSUpload\uploaded_files") = True Then
                            If objfso.FileExists(gstrLocalCDrive & "DSUpload\uploaded_files\" + strUploadedFileName) Then
                                objfso.DeleteFile(gstrLocalCDrive & "DSUpload\uploaded_files\" + strUploadedFileName)
                            End If
                            objfso.MoveFile(txtfilelocation.Text, gstrLocalCDrive & "DSUpload\uploaded_files\")
                        Else
                            objfso.CreateFolder(gstrLocalCDrive & "DSUpload\uploaded_files")
                            objfso.MoveFile(txtfilelocation.Text, gstrLocalCDrive & "DSUpload\uploaded_files\")
                        End If
                        SqlConnectionclass.CommitTran()
                        MsgBox("Transaction Saved Successfully", MsgBoxStyle.Information, ResolveResString(100))
                        ResetData()
                    End If
                End With
            End Using
            Using sqlCmd As SqlCommand = New SqlCommand
                With sqlCmd
                    .CommandText = "USP_DSUploader_AUTOMAILER"
                    .CommandTimeout = 0
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                    .Parameters.AddWithValue("@TYPE", "Uploader")
                    SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                End With
            End Using
            ResetData()
            Return True

ErrHandler:
        Catch ex As Exception
            SqlConnectionclass.RollbackTran()
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Function
    Private Sub FillWrongData()
        Try
            Dim strSQl As String = "Select a.Account_Code,Trans_date,a.ITEM_CODE,a.Cust_Drgno,Balance_Qty,a.Ent_dt,a.Ent_UserId,b.Product_Start_date, " & _
            "b.Product_End_date   from  TEMP_Dailymktschedule_UPLOADER a, CustItem_Mst b  where a.Account_Code=b.Account_Code " & _
            "and a.Item_Code=b.item_code and a.Cust_Drgno=b.Cust_Drgno and a.UNIT_CODE=b.UNIT_CODE  and " & _
            "trans_date NOT between  b.Product_Start_date and b.Product_End_date and  a.unit_code='" & gstrUNITID & "'  and IP_ADDRESS='" & gstrIpaddressWinSck & "'"
            Using dt As DataTable = SqlConnectionclass.GetDataTable(strSQl)
                If dt.Rows.Count > 0 Then
                    Call SetSpreadProperty()
                    With fspforecastdetail
                        For Each row As DataRow In dt.Rows
                            .MaxRows = .MaxRows + 1
                            .Row = .MaxRows
                            .BackColorStyle = FPSpreadADO.BackColorStyleConstants.BackColorStyleUnderGrid
                            .set_RowHeight(.Row, 300)
                            Call SetSpreadColTypes(.Row)
                            .Col = enmDSdetail.CustomerCode : .Text = Convert.ToString(row("Account_Code")) : .ForeColor = System.Drawing.Color.Red
                            .Col = enmDSdetail.ItemCode : .Text = Convert.ToString(row("ITEM_CODE")) : .ForeColor = System.Drawing.Color.Red
                            .Col = enmDSdetail.Cust_DrgwNo : .Text = Convert.ToString(row("Cust_Drgno")) : .ForeColor = System.Drawing.Color.Red
                            .Col = enmDSdetail.Trans_date : .Text = VB6.Format(Convert.ToString(row("Trans_date")), gstrDateFormat) : .ForeColor = System.Drawing.Color.Red
                            .Col = enmDSdetail.Product_Start_date : .Text = VB6.Format(Convert.ToString(row("Product_Start_date")), gstrDateFormat) : .ForeColor = System.Drawing.Color.Red
                            .Col = enmDSdetail.Product_End_date : .Text = VB6.Format(Convert.ToString(row("Product_End_date")), gstrDateFormat) : .ForeColor = System.Drawing.Color.Red
                            .Col = enmDSdetail.ItemDesc : .ColHidden = True
                            .Col = enmDSdetail.Product_Start_date : .ColHidden = False
                            .Col = enmDSdetail.Product_End_date : .ColHidden = False
                            .Col = enmDSdetail.BalanceQty : .ColHidden = True
                        Next
                    End With
                    cmdsave.Enabled = False
                End If

            End Using


        Catch ex As Exception
            SqlConnectionclass.RollbackTran()
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
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

    
    Private Sub upload_DS_data()
        On Error GoTo ErrHandler
        Dim objexl As New Excel.Application
        Dim objfso As New Scripting.FileSystemObject
        Dim row As Long
        Dim strcustomercode As String
        Dim strcustomerDrgNo As String
        Dim strItemDesc As String
        Dim stritemcode As String
        Dim strDate As String
        Dim strQuantity As String
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
                row1 = 7
                strcustomercode = objexl.Cells(row, 2).Value
                If UCase(strcustomercode) <> UCase(txtcustomercode.Text) Then
                    MsgBox("Customer Code In File Is Not Matching With Customer Code Selected", MsgBoxStyle.Information, ResolveResString(100))
                    objexl.Quit()
                    objexl = Nothing
                    objfso = Nothing
                    Exit Sub
                Else
                    Call SetSpreadProperty()
                    Dim intRowCount As Integer = 0
                    Dim intCount As Integer = 1

                    While (row1 <= objexl.Rows.Count)
                        stritemcode = objexl.Cells(row1, 1).Value
                        strItemDesc = objexl.Cells(row1, 2).Value
                        strcustomerDrgNo = objexl.Cells(row1, 3).Value
                        strDate = objexl.Cells(row1, 4).Value
                        strQuantity = objexl.Cells(row1, 5).Value
                        If stritemcode <> "" Then
                            If strDate <> "" And CStr(strQuantity) <> "" Then   '101289013 -on stritemcode- MILIND
                                With fspforecastdetail
                                    .MaxRows = .MaxRows + 1
                                    .Row = .MaxRows
                                    .BackColorStyle = FPSpreadADO.BackColorStyleConstants.BackColorStyleUnderGrid
                                    .set_RowHeight(.Row, 300)
                                    Call SetSpreadColTypes(.Row)
                                    Call .SetText(enmDSdetail.CustomerCode, intCount, txtcustomercode.Text)
                                    Call .SetText(enmDSdetail.ItemCode, intCount, stritemcode)
                                    Call .SetText(enmDSdetail.ItemDesc, intCount, strItemDesc)
                                    Call .SetText(enmDSdetail.Cust_DrgwNo, intCount, strcustomerDrgNo)
                                    Call .SetText(enmDSdetail.Trans_date, intCount, VB6.Format(strDate, gstrDateFormat))
                                    'Call .SetText(enmDSitem.ScheduleQty, intCount, objgetDetail.GetValue("ScheduleQty"))
                                    'Call .SetText(enmDSitem.DispatchQty, intCount, objgetDetail.GetValue("DispatchQty"))
                                    Call .SetText(enmDSdetail.BalanceQty, intCount, strQuantity)
                                    .Col = enmDSdetail.Product_Start_date : .ColHidden = True
                                    .Col = enmDSdetail.Product_End_date : .ColHidden = True
                                End With
                            End If
                            intCount = intCount + 1
                            row1 = row1 + 1
                        Else
                            Exit While
                        End If

                    End While
                    objexl.Quit()
                    objexl = Nothing
                    objfso = Nothing
                    If Validations() = False Then
                        ResetData()
                        Exit Sub

                    End If
                    If validate_holiday() <> "" Then
                        MsgBox("Following Is\Are Not Working Day(s) " & vbCrLf & validate_holiday() & vbCrLf, MsgBoxStyle.Information, ResolveResString(100))
                        Exit Sub
                    End If
                  
                    fspforecastdetail.Enabled = True
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
    
    Private Function validate_itemcode() As String
        On Error GoTo ErrHandler
        Dim strSQL As String
        Dim objgetitem As New ClsResultSetDB
        Dim stritemlist As String
        strSQL = "select distinct product_no from forecast_mst_temp where customer_code='" & Trim(txtcustomercode.Text) & "' and ipaddress='" & gstrIpaddressWinSck & "' and Unit_code = '" & gstrUNITID & "'" & "and product_no not in(select cim.item_code from custitem_mst cim inner join item_mst im on cim.item_code=im.item_code and cim.Unit_Code=im.unit_code where cim.account_code='" & Trim(txtcustomercode.Text) & "' and active=1 and im.status='A' and cim.Unit_code = '" & gstrUNITID & "')"
        objgetitem.GetResult(strSQL)
        If Not objgetitem.EOFRecord Then
            While Not objgetitem.EOFRecord
                If stritemlist = "" Then
                    stritemlist = objgetitem.GetValue("product_no")
                Else
                    stritemlist = stritemlist & vbCrLf & objgetitem.GetValue("product_no")
                End If
                objgetitem.MoveNext()
            End While
            validate_itemcode = stritemlist
        End If
        objgetitem = Nothing
        Exit Function
ErrHandler:
        objgetitem = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function validate_holiday() As String
        On Error GoTo ErrHandler
        Dim strSQL As String
        Dim objgetholiday As New ClsResultSetDB
        Dim strholidaylist As String
        strSQL = "select distinct due_date from forecast_mst_temp where customer_code='" & Trim(txtcustomercode.Text) & "' and ipaddress='" & gstrIpaddressWinSck & "' and Unit_code = '" & gstrUNITID & "' and due_date in(select dt from Calendar_MST where work_flg=1 and Unit_code = '" & gstrUNITID & "')"
        objgetholiday.GetResult(strSQL)
        If Not objgetholiday.EOFRecord Then
            While Not objgetholiday.EOFRecord
                If strholidaylist = "" Then
                    strholidaylist = objgetholiday.GetValue("due_Date")
                Else
                    strholidaylist = strholidaylist & vbCrLf & objgetholiday.GetValue("due_Date")
                End If
                objgetholiday.MoveNext()
            End While
            validate_holiday = VB6.Format(strholidaylist, gstrDateFormat)
        End If
        objgetholiday = Nothing
        Exit Function
ErrHandler:
        objgetholiday = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub setitemgrdproperty()
        On Error GoTo Errorhandler
        With fspitem
            .MaxRows = 0
            .MaxCols = 0
            .set_RowHeight(0, 400)
            .Row = 0
            .ColsFrozen = 2
            .MaxCols = enmDSitem.DSNO
            .Col = enmDSitem.status : .Text = "Select" : .set_ColWidth(enmDSitem.status, 600)
            .Col = enmDSitem.CustomerCode : .Text = "Cutomer Code" : .set_ColWidth(enmDSitem.CustomerCode, 1600)
            .Col = enmDSitem.ItemCode : .Text = "Item Code" : .set_ColWidth(enmDSitem.ItemCode, 1000)
            .Col = enmDSitem.ItemDesc : .Text = "Item Desc" : .set_ColWidth(enmDSitem.ItemDesc, 2000)
            .Col = enmDSitem.Cust_DrgwNo : .Text = "Cust Drgn No" : .set_ColWidth(enmDSitem.Cust_DrgwNo, 2000)
            .Col = enmDSitem.Trans_date : .Text = "Trans Date" : .set_ColWidth(enmDSitem.Trans_date, 860)
            .Col = enmDSitem.ScheduleQty : .Text = "Schedule Qty" : .set_ColWidth(enmDSitem.ScheduleQty, 1000)
            .Col = enmDSitem.DispatchQty : .Text = "Dispatch Qty" : .set_ColWidth(enmDSitem.DispatchQty, 1000)
            .Col = enmDSitem.BalanceQty : .Text = "Balance Qty" : .set_ColWidth(enmDSitem.BalanceQty, 1000)
            .Col = enmDSitem.Serial_No : .Text = "Serial No" : .set_ColWidth(enmDSitem.Serial_No, 1000)
            .Col = enmDSitem.DSNO : .Text = "DS No" : .set_ColWidth(enmDSitem.DSNO, 1000)
        End With
        fspitem.ReDraw = True
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub setitemgrdcoltypes(ByRef pintRowNo As Long)
        On Error GoTo Errorhandler
        Dim i As Long
        With Me.fspitem
            .Row = pintRowNo
            .Col = enmDSitem.status : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .ForeColor = System.Drawing.Color.Black
            .Col = enmDSitem.CustomerCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .ForeColor = System.Drawing.Color.Black
            .Col = enmDSitem.ItemCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .ForeColor = System.Drawing.Color.Black
            .Col = enmDSitem.ItemDesc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .ForeColor = System.Drawing.Color.Black
            .Col = enmDSitem.Cust_DrgwNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmDSitem.Trans_date : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmDSitem.ScheduleQty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmDSitem.DispatchQty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmDSitem.BalanceQty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmDSitem.Serial_No : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmDSitem.DSNO : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
        End With
        fspitem.ReDraw = True
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub addrowinitemgrd(ByRef pintRows As Long)
        On Error GoTo ErrHandler
        Dim intRowHeight As Long
        With Me.fspitem
            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            For intRowHeight = 1 To pintRows
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .set_RowHeight(.Row, 300)
                Call setitemgrdcoltypes(.Row)
            Next intRowHeight
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
        On Error GoTo ErrHandler
        Dim intcounter As Long
        Dim varstatus As Object
        Dim varItemCode As Object
        Dim varCustdrgno As Object
        Dim varBalQty As Object
        Dim varTransDate As Object
        Dim varItemDesc As Object
        Dim itemcount As Long
        Dim intcounter1 As Long
        Dim interval As Long
        Dim dtDate As Date
        Dim datecounter As Long
        Dim objcalender As New ClsResultSetDB
        itemcount = 6
        interval = DateDiff(Microsoft.VisualBasic.DateInterval.Day, dtfromdate.Value, dttodate.Value) + 1
        With fspitem
            objexl.Cells._Default(1, 1).Font.Bold = True

            For intcounter = 1 To .MaxRows
                varstatus = Nothing
                varItemCode = Nothing
                varCustdrgno = Nothing
                varBalQty = Nothing
                varItemDesc = Nothing
                varTransDate = Nothing
                Call .GetText(enmDSitem.status, intcounter, varstatus)
                varstatus = IIf(varstatus.ToString = "", 0, varstatus)
                If varstatus = 1 Then
                    Call .GetText(enmDSitem.ItemCode, intcounter, varItemCode)
                    Call .GetText(enmDSitem.Cust_DrgwNo, intcounter, varCustdrgno)
                    Call .GetText(enmDSitem.ItemDesc, intcounter, varItemDesc)
                    Call .GetText(enmDSitem.BalanceQty, intcounter, varBalQty)
                    Call .GetText(enmDSitem.Trans_date, intcounter, varTransDate)
                    itemcount = itemcount + 1
                    'objexl.Cells._Default(itemcount, 3).numberformat = "0"
                    objexl.Cells._Default(itemcount, 1) = varItemCode
                    objexl.Cells._Default(itemcount, 2) = varItemDesc
                    objexl.Cells._Default(itemcount, 3) = varCustdrgno
                    objexl.Cells._Default(itemcount, 4) = varTransDate
                    objexl.Cells._Default(itemcount, 5) = varBalQty
                    'objexl.Cells._Default(4, 1 + itemcount) = get_itemdescription(varItemCode)
                    'objexl.Cells._Default(4, 1 + itemcount).Font.Bold = True
                    'objexl.Cells._Default(4, 1 + itemcount).ColumnWidth = Len(get_itemdescription(varItemCode)) + 5
                    'objexl.Cells._Default(5, 1 + itemcount).Font.Bold = True
                    objexl.Cells._Default(itemcount, 1).ColumnWidth = Len(varItemCode) + 5
                    objexl.Cells._Default(itemcount, 2).ColumnWidth = Len(varItemDesc) + 5
                    objexl.Cells._Default(itemcount, 3).ColumnWidth = Len(varTransDate) + 5

                    objexl.Cells._Default(itemcount, 4).ColumnWidth = Len(varTransDate) + 5
                    objexl.Cells._Default(itemcount, 5).ColumnWidth = Len(varItemDesc) + 5

                End If
            Next
        End With
       
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function validate_itemselected() As Boolean
        On Error GoTo ErrHandler
        Dim intcounter As Long
        Dim varstatus As Object
        validate_itemselected = False
        With fspitem
            For intcounter = 1 To .MaxRows
                varstatus = Nothing
                Call .GetText(enmDSitem.status, intcounter, varstatus)
                varstatus = IIf(varstatus.ToString = "", 0, varstatus)
                If varstatus = 1 Then
                    validate_itemselected = True
                    Exit For
                End If
            Next
        End With
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub clear_fields()
        On Error GoTo ErrHandler
        fspitem.MaxCols = 0
        fspitem.MaxRows = 0
        lblcustname.Text = ""
        txtcustomerhelp.Text = ""
        dtfromdate.Value = GetServerDate()
        dttodate.Value = GetServerDate()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function validate_calender() As Boolean
        On Error GoTo ErrHandler
        Dim objgetcaldata As New ClsResultSetDB
        Dim noofdays As Integer
        noofdays = DateDiff(Microsoft.VisualBasic.DateInterval.Day, dtfromdate.Value, dttodate.Value) + 1
        validate_calender = True
        objgetcaldata.GetResult("select count(*) as noofdays from Calendar_mst where dt between '" & getDateForDB(dtfromdate.Value) & "' and '" & getDateForDB(dttodate.Value) & "' and Unit_code = '" & gstrUNITID & "'")
        If Not objgetcaldata.EOFRecord Then
            If noofdays <> objgetcaldata.GetValue("noofdays") Then
                validate_calender = False
                Exit Function
            End If
        End If
        objgetcaldata = Nothing
        Exit Function
ErrHandler:
        objgetcaldata = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function validate_daterange() As Boolean
        On Error GoTo ErrHandler
        Dim intinterval As Integer
        validate_daterange = True
        intinterval = DateDiff(Microsoft.VisualBasic.DateInterval.Month, dtfromdate.Value, dttodate.Value)
        If intinterval > 24 Then
            validate_daterange = False
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function validate_excel_format(ByRef objexl As Excel.Application) As Boolean
        On Error GoTo ErrHandler
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

        validate_excel_format = True
        row1 = 6 : row2 = 6
        row = 7 : col = 1
        stritemcode = Nothing
        stritemcode = objexl.Cells(row, col).Value
        If stritemcode = "" Then
            validate_excel_format = False
            Exit Function
        End If
        intRow = 7
        While (intRow <= objexl.Rows.Count)
            stritemcode = objexl.Cells(intRow, 1).Value
            strCustDrgNo = objexl.Cells(intRow, 3).Value
            strTransDate = objexl.Cells(intRow, 4).Value
            strBalanceQty = objexl.Cells(intRow, 5).Value
            If stritemcode <> "" AndAlso strCustDrgNo <> "" Then
                intRow = intRow + 1
            Else
                Exit While
            End If


        End While
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub dtfromdate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtfromdate.ValueChanged
        On Error GoTo ErrHandler
        If dtfromdate.Value < GetServerDate() Then
            dtfromdate.Value = GetServerDate()
        ElseIf dtfromdate.Value > dttodate.Value Then
            dtfromdate.Value = GetServerDate()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub dttodate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dttodate.ValueChanged
        On Error GoTo ErrHandler
        If dttodate.Value < dtfromdate.Value Then
            dttodate.Value = dtfromdate.Value
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdsave_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdsave.Click
        On Error GoTo ErrHandler
        Dim strcustomercode As String
        Dim stritemlist As String
        Dim intvalue As Long

        If fspforecastdetail.MaxRows > 0 Then
            If txtcustomercode.Text <> "" Then
                If SAVEDATA() = False Then Exit Sub
                RefreshForm()
            Else
                MsgBox("Customer Code Can Not Be Blank", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
        Else
            MsgBox("Please Upload DS First", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If

        Exit Sub
ErrHandler:


        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Function Validations() As Boolean
        Dim strItemCode As String = ""
        Dim strCustDrgNo As String = ""
        Dim strDate As String = ""
        Dim strQty As String = ""
        Dim strInnerItemCode As String = ""
        Dim strInnerCustDrgNo As String = ""
        Dim strInnerDate As String = ""
        Dim strInnerQty As String = ""
        Dim strQuery As String = ""
        Dim strSOItem As String = ""
        Try
            With fspforecastdetail
                For intRow As Integer = 1 To .MaxRows
                    .Row = intRow
                    .Col = enmDSdetail.ItemCode
                    strItemCode = (.Text)

                    .Col = enmDSdetail.Cust_DrgwNo
                    strCustDrgNo = (.Text)

                    .Col = enmDSdetail.Trans_date
                    strDate = (.Text)

                    .Col = enmDSdetail.BalanceQty
                    strQty = (.Text)

                    strQuery = "select top 1 Item_code from cust_ord_dtl where UNIT_CODE='" + gstrUNITID + "' and account_code='" & Trim(txtcustomercode.Text) & "' and item_code='" & strItemCode & "' and cust_drgno = '" & strCustDrgNo & "' and  active_flag='A' and authorized_flag=1"
                    strSOItem = SqlConnectionclass.ExecuteScalar(strQuery)
                    If strSOItem = "" Then
                        MsgBox("Please check mapping for Item in Sale Order " & strItemCode, MsgBoxStyle.Information, ResolveResString(100))
                        Return False

                    ElseIf strDate < GetServerDate() Then
                        MsgBox("Backdate Schedule cannot be uploaded  for Item " & strItemCode, MsgBoxStyle.Information, ResolveResString(100))
                        Return False

                    End If

                    For innerRow As Integer = 1 To .MaxRows
                        .Row = innerRow
                        .Col = enmDSdetail.ItemCode
                        strInnerItemCode = (.Text)

                        .Col = enmDSdetail.Cust_DrgwNo
                        strInnerCustDrgNo = (.Text)

                        .Col = enmDSdetail.Trans_date
                        strInnerDate = (.Text)

                        .Col = enmDSdetail.BalanceQty
                        strInnerQty = (.Text)
                        If intRow <> innerRow And strItemCode = strInnerItemCode And strCustDrgNo = strInnerCustDrgNo And strDate = strInnerDate Then
                            MsgBox("Duplicate Item cannot be uploaded  for Item " & strItemCode, MsgBoxStyle.Information, ResolveResString(100))
                            Return False

                        End If
                    Next

                Next
                Return True
            End With
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
   
End Class