Option Strict Off
Option Explicit On
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
Friend Class frmMKTTRN0063
    Inherits System.Windows.Forms.Form
    Dim strUploadedFileName As String
    Dim mintIndex As Short
    Dim mflag As Short
    Private Enum enmforecastdetail
        ItemCode = 1
        Item_Description
        Cust_DrgwNo
        UOM
        Quantity
        DELIVERY_DATE
    End Enum
    Private Enum enmforecastitem
        status = 1
        ItemCode
        Item_Description
        Cust_DrgwNo
    End Enum
    Private Sub SetSpreadProperty()
        On Error GoTo Errorhandler
        With fspforecastdetail
            .MaxRows = 0
            .MaxCols = 0
            .set_RowHeight(0, 400)
            .Row = 0
            .ColsFrozen = 0
            .MaxCols = enmforecastdetail.DELIVERY_DATE
            .Col = enmforecastdetail.ItemCode : .Text = "Item Code" : .set_ColWidth(enmforecastdetail.ItemCode, 1600)
            .Col = enmforecastdetail.Item_Description : .Text = "Item Description" : .set_ColWidth(enmforecastdetail.Item_Description, 2500)
            .Col = enmforecastdetail.Cust_DrgwNo : .Text = "Cust Drgn No" : .set_ColWidth(enmforecastdetail.Cust_DrgwNo, 2000)
            .Col = enmforecastdetail.UOM : .Text = "UOM" : .set_ColWidth(enmforecastdetail.UOM, 860)
            .Col = enmforecastdetail.Quantity : .Text = "Quantity" : .set_ColWidth(enmforecastdetail.Quantity, 860)
            .Col = enmforecastdetail.DELIVERY_DATE : .Text = "Delivery Date" : .set_ColWidth(enmforecastdetail.DELIVERY_DATE, 1000)
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
            .Col = enmforecastdetail.ItemCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmforecastdetail.Item_Description : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmforecastdetail.Cust_DrgwNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmforecastdetail.UOM : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmforecastdetail.Quantity : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmforecastdetail.DELIVERY_DATE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
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
                    Call .SetText(enmforecastitem.status, intcounter, "1")
                Else
                    Call .SetText(enmforecastitem.status, intcounter, "0")
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
        If Len(Me.txtcustomerhelp.Text) = 0 Then
            strCustHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select a.customer_Code, a.cust_name as customer_name from Customer_Mst a where a.Unit_code = '" & gstrUNITID & "' and ((isnull(a.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= a.deactive_date))", "Customer Codes List", 1)
        Else
            strCustHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select a.customer_Code, a.cust_name  as customer_name from Customer_Mst a where a.customer_code='" & Trim(txtcustomerhelp.Text) & "' and a.Unit_code = '" & gstrUNITID & "' and ((isnull(a.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= a.deactive_date))", "Customer Codes List", 1)
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
        strUploadingpath = gstrLocalCDrive & "Forecast\uploaded_files"
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
        If validate_daterange() = False Then
            MsgBox("Date Range Cannot Be More Than 24 Months ( 2 years) ", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If
        If validate_calender() = False Then
            MsgBox("Sales Calender Is Not Defined For Specified Date Range", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If
        If fspitem.MaxRows > 0 And validate_itemselected() = True Then
            If objfso.FolderExists(gstrLocalCDrive & "forecast") = True Then
                objfolder = objfso.GetFolder(gstrLocalCDrive & "forecast")
                For Each objfile In objfolder.Files
                    stroldfilename = objfile.Name
                    If UCase(Mid(stroldfilename, 10, 8)) = UCase(txtcustomerhelp.Text) Then
                        If objfso.FolderExists(gstrLocalCDrive & "forecast\Backup") = True Then
                            If stroldfilename <> "" Then
                                objfso.MoveFile(gstrLocalCDrive & "forecast\" & stroldfilename, gstrLocalCDrive & "forecast\backup\")
                            End If
                        Else
                            objfso.CreateFolder(gstrLocalCDrive & "forecast\backup")
                            If stroldfilename <> "" Then
                                objfso.MoveFile(gstrLocalCDrive & "forecast\" & stroldfilename, gstrLocalCDrive & "forecast\backup\")
                            End If
                        End If
                    End If
                Next objfile
                OBj_wB = OBJexlformat.Workbooks.Add
                strnewfilename = getfilename(Trim(txtcustomerhelp.Text))
                OBj_wB.SaveAs(gstrLocalCDrive & "forecast\" & Replace(strnewfilename, ":", ""))
                OBJexlformat.Workbooks.Close()
                OBj_wB = Nothing
                OBJexlformat.Workbooks.Open(gstrLocalCDrive & "forecast\" & Replace(strnewfilename, ":", ""))
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
                OBJexlformat.Cells._Default(5, 1) = "Drawing No."
                OBJexlformat.Cells._Default(5, 1).Font.Bold = True
                OBJexlformat.Cells._Default(6, 1) = "Due Date \ Item Code"
                OBJexlformat.Cells._Default(6, 1).Font.Bold = True
                Call fillitems(OBJexlformat)
                OBJexlformat.ActiveWorkbook.Save()
                OBJexlformat.Quit()
                MsgBox("Excel Sheet Generated Successfully " & vbCrLf & "location--> " & " " & gstrLocalCDrive & "Forecast\" & Replace(strnewfilename, ":", ""), MsgBoxStyle.Information, ResolveResString(100))
                clear_fields()
                chkall.CheckState = System.Windows.Forms.CheckState.Unchecked
                chkall.Enabled = False
                OBJexlformat = Nothing
            Else
                objfso.CreateFolder(gstrLocalCDrive & "forecast")
                OBj_wB = OBJexlformat.Workbooks.Add
                strnewfilename = getfilename(Trim(txtcustomerhelp.Text))
                OBj_wB.SaveAs(gstrLocalCDrive & "forecast\" & Replace(strnewfilename, ":", ""))
                OBJexlformat.Workbooks.Close()
                OBj_wB = Nothing
                OBJexlformat.Workbooks.Open(gstrLocalCDrive & "forecast\" & Replace(strnewfilename, ":", ""))
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
                OBJexlformat.Cells._Default(5, 1) = "Due Date"
                OBJexlformat.Cells._Default(5, 1).Font.Bold = True
                Call fillitems(OBJexlformat)
                OBJexlformat.ActiveWorkbook.Save()
                OBJexlformat.Quit()
                MsgBox("Excel Sheet Generated Successfully " & vbCrLf & "location--> " & " " & gstrLocalCDrive & "forecast\" & Replace(strnewfilename, ":", ""), MsgBoxStyle.Information, ResolveResString(100))
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
                Call upload_forecast_data()
            Else
                MsgBox("Please Select File Location To Upload Forecast", MsgBoxStyle.Information, ResolveResString(100))
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
        If Len(Trim(txtcustomerhelp.Text)) = 0 Then
            MsgBox("Please Select Customer Code First", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If
        strSQL = "select distinct a.item_code,b.description,a.cust_drgno from custitem_mst" & " a inner join item_mst b on a.item_code=b.item_code and a.Unit_Code=b.Unit_Code and a.active=1 and b.status='A'" & " and a.account_code='" & Trim(txtcustomerhelp.Text) & "' and a.Unit_code = '" & gstrUNITID & "' "
        objgetitems.GetResult(strSQL)
        If Not objgetitems.EOFRecord Then
            Call setitemgrdproperty()
            While Not objgetitems.EOFRecord
                With fspitem
                    addrowinitemgrd((1))
                    Call .SetText(enmforecastitem.ItemCode, .MaxRows, objgetitems.GetValue("item_code"))
                    Call .SetText(enmforecastitem.Item_Description, .MaxRows, objgetitems.GetValue("description"))
                    Call .SetText(enmforecastitem.Cust_DrgwNo, .MaxRows, objgetitems.GetValue("cust_drgno"))
                    objgetitems.MoveNext()
                End With
            End While
            chkall.Enabled = True
            cmdgenerateformat.Enabled = True
        Else
            MsgBox("No Item Is Defined For The Specified Customer", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If
        Exit Sub
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
        mintIndex = mdifrmMain.AddFormNameToWindowList(CTLFORECASTMASTER.Tag)
        FitToClient(Me, framain, CTLFORECASTMASTER, frabuttons)
        Call FillLabelFromResFile(Me)
        cmdcustomercode.Image = My.Resources.ico111.ToBitmap
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
        SSTab1.SelectedIndex = 0
        cmdsave.Enabled = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
            strCustHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select a.customer_Code, a.cust_name as customer_name from Customer_Mst a where a.Unit_code = '" & gstrUNITID & "' and ((isnull(a.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= a.deactive_date))", "Customer Codes List", 1)
        Else
            strCustHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select a.customer_Code, a.cust_name  as customer_name from Customer_Mst a where a.customer_code='" & Trim(txtcustomercode.Text) & "' and a.Unit_code = '" & gstrUNITID & "' and ((isnull(a.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= a.deactive_date))", "Customer Codes List", 1)
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
            strSQL = "select a.customer_Code, a.cust_name from Customer_Mst a where a.customer_code='" & Trim(txtcustomercode.Text) & "' and a.Unit_code = '" & gstrUNITID & "' and ((isnull(a.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= a.deactive_date))"
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
        strSQL = "select product_no,due_date,sum(quantity)as quantity from forecast_mst_temp where customer_code='" & Trim(txtcustomercode.Text) & "' and ipaddress='" & gstrIpaddressWinSck & "' and Unit_code = '" & gstrUNITID & "' and due_date>=convert(varchar(12),getdate(),106) and quantity >0 group by product_no,due_date"
        objgetDetail.GetResult(strSQL)
        With fspforecastdetail
            If Not objgetDetail.EOFRecord Then
                intRowCount = objgetDetail.GetNoRows
                Call addRowAtEnterKeyPress(intRowCount)
                objgetDetail.MoveFirst()
                For intcounter = 1 To intRowCount
                    .Row = intcounter
                    .Col = enmforecastdetail.ItemCode : Call .SetText(enmforecastdetail.ItemCode, intcounter, objgetDetail.GetValue("product_no"))
                    .Col = enmforecastdetail.Item_Description
                    Call .SetText(enmforecastdetail.Item_Description, intcounter, get_itemdescription(objgetDetail.GetValue("product_no")))
                    .Col = enmforecastdetail.Cust_DrgwNo
                    Call .SetText(enmforecastdetail.Cust_DrgwNo, intcounter, get_customer_drgno(objgetDetail.GetValue("product_no")))
                    .Col = enmforecastdetail.UOM
                    Call .SetText(enmforecastdetail.UOM, intcounter, get_item_uom(objgetDetail.GetValue("product_no")))
                    .Col = enmforecastdetail.DELIVERY_DATE : Call .SetText(enmforecastdetail.DELIVERY_DATE, intcounter, VB6.Format(objgetDetail.GetValue("due_date"), gstrDateFormat))
                    .Col = enmforecastdetail.Quantity : Call .SetText(enmforecastdetail.Quantity, intcounter, objgetDetail.GetValue("quantity"))
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
            fspitem.Col = enmforecastitem.Item_Description
            fspitem.Font = VB6.FontChangeBold(fspitem.Font, False)
            fspitem.Col = enmforecastitem.ItemCode
            fspitem.Font = VB6.FontChangeBold(fspitem.Font, False)
        Next
        For intRowNo = 0 To fspitem.MaxRows - 1
            strSearchItem = Nothing
            If strSearchBy = "Item Code" Then
                Call fspitem.GetText(enmforecastitem.ItemCode, intRowNo, strSearchItem)
            ElseIf strSearchBy = "Item Description" Then
                Call fspitem.GetText(enmforecastitem.Item_Description, intRowNo, strSearchItem)
            End If
            If UCase(strSearchItem) Like UCase(TxtSearch.Text) & "*" Then
                If strSearchBy = "Item Description" Then
                    fspitem.Col = enmforecastitem.Item_Description
                ElseIf strSearchBy = "Item Code" Then
                    fspitem.Col = enmforecastitem.ItemCode
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
            strSQL = "select a.customer_Code, a.cust_name from Customer_Mst a where a.customer_code='" & Trim(txtcustomerhelp.Text) & "' and a.Unit_code = '" & gstrUNITID & "' and ((isnull(a.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= a.deactive_date))"
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
    Private Sub save_data()
        On Error GoTo ErrHandler
        Dim strInsert As String
        Dim intcounter As Long
        Dim varduedate As Object
        Dim varproductno As Object
        Dim VarCustDrgNo As Object
        Dim varquantity As Object
        With fspforecastdetail
            For intcounter = 1 To fspforecastdetail.MaxRows
                varproductno = Nothing
                varduedate = Nothing
                varquantity = Nothing
                VarCustDrgNo = Nothing
                Call .GetText(enmforecastdetail.ItemCode, intcounter, varproductno)
                Call .GetText(enmforecastdetail.DELIVERY_DATE, intcounter, varduedate)
                Call .GetText(enmforecastdetail.Quantity, intcounter, varquantity)
                Call .GetText(enmforecastdetail.Cust_DrgwNo, intcounter, VarCustDrgNo)

                'If varproductno = "MSA65200349" Then
                '    MsgBox("abc", MsgBoxStyle.Information)
                'End If
                strInsert = "insert into forecast_mst(Customer_code,product_no,Due_date,quantity,ent_userid,ent_dt,upd_userid,upd_dt,Enagare_UNLOC, Unit_Code,Cust_Drgno)" & " values('" & txtcustomercode.Text & "','" & varproductno & "','" & getDateForDB(varduedate) & "'," & varquantity & ",'" & mP_User & "',getdate(),'" & mP_User & "',getdate(),'fcst', '" & gstrUNITID & "', '" & VarCustDrgNo & "')"
                Call delete_previous_data(getDateForDB(varduedate), varproductno, varquantity, strInsert)
            Next
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
    Private Function get_item_uom(ByRef stritemcode As String) As String
        On Error GoTo ErrHandler
        Dim objgetitemuom As New ClsResultSetDB
        objgetitemuom.GetResult("select cons_measure_code from item_mst where item_Code='" & stritemcode & "' and Unit_code = '" & gstrUNITID & "'")
        If Not (objgetitemuom.BOFRecord And objgetitemuom.BOFRecord) Then
            get_item_uom = objgetitemuom.GetValue("cons_measure_Code")
        Else
            get_item_uom = ""
        End If
        objgetitemuom = Nothing
        Exit Function
ErrHandler:
        objgetitemuom = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function get_customer_drgno(ByRef stritemcode As String) As String
        On Error GoTo ErrHandler
        Dim objgetdrgno As New ClsResultSetDB
        If SSTab1.SelectedIndex = 0 Then
            objgetdrgno.GetResult("select top 1 cust_drgno from custitem_mst where account_code='" & Trim(txtcustomerhelp.Text) & "' and item_code='" & stritemcode & "' and active=1 and Unit_code = '" & gstrUNITID & "'")
        Else
            objgetdrgno.GetResult("select top 1 cust_drgno from custitem_mst where account_code='" & Trim(txtcustomercode.Text) & "' and item_code='" & stritemcode & "' and active=1 and Unit_code = '" & gstrUNITID & "'")
        End If
        If Not (objgetdrgno.BOFRecord And objgetdrgno.EOFRecord) Then
            get_customer_drgno = objgetdrgno.GetValue("cust_drgno")
        Else
            get_customer_drgno = ""
        End If
        objgetdrgno = Nothing
        Exit Function
ErrHandler:
        objgetdrgno = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub upload_forecast_data()
        On Error GoTo ErrHandler
        Dim objexl As New Excel.Application
        Dim objfso As New Scripting.FileSystemObject
        Dim row As Long
        Dim strcustomercode As String
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
                strcustomercode = objexl.Cells(row, 2).Value
                If UCase(strcustomercode) <> UCase(txtcustomercode.Text) Then
                    MsgBox("Customer Code In File Is Not Matching With Customer Code Selected", MsgBoxStyle.Information, ResolveResString(100))
                    objexl.Quit()
                    objexl = Nothing
                    objfso = Nothing
                    Exit Sub
                Else
                    mP_Connection.Execute("delete from forecast_mst_temp where customer_code='" & txtcustomercode.Text & "' and ipaddress='" & gstrIpaddressWinSck & "' and Unit_code = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    row1 = 7 : row2 = 7
                    row = 6 : col = 2
                    stritemcode = objexl.Cells(row, col).Value
                   
                    While stritemcode <> "" And (col <= objexl.Columns.Count)
                        row = 6 : col = col
                        stritemcode = objexl.Cells(row, col).Value

                        row1 = 7 : col1 = 1
                        strDate = objexl.Cells(row1, col1).Value
                        row2 = 7 : col2 = 2
                        strQuantity = objexl.Cells(row2, col2).Value
                        If stritemcode <> "" Then
                            While strDate <> ""
                                row1 = row1 : col1 = 1
                                strDate = objexl.Cells(row1, col1).Value
                                row2 = row2 : col2 = col
                                strQuantity = objexl.Cells(row2, col2).Value
                                If strDate <> "" And CStr(strQuantity) <> "" Then   '101289013 -on stritemcode- MILIND
                                    strSQL = "insert into forecast_mst_temp(Customer_code,product_no,Due_date,quantity,ent_userid,ent_dt,upd_userid,upd_dt,Enagare_UNLOC,ipaddress,Unit_code)" & " values('" & txtcustomercode.Text & "','" & stritemcode.Trim & "','" & getDateForDB(VB6.Format(strDate, gstrDateFormat)) & "'," & CInt(strQuantity) & ",'" & mP_User & "',getdate(),'" & mP_User & "',getdate(),'fcst','" & gstrIpaddressWinSck & "', '" & gstrUNITID & "')"
                                    mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                End If
                                row1 = row1 + 1
                                row2 = row2 + 1
                            End While
                        End If
                        col = col + 1

                    End While
                    objexl.Quit()
                    objexl = Nothing
                    objfso = Nothing
                    'If validate_itemcode() <> "" Then
                    '    MsgBox("Following Items Are Not Defined Or Inactive " & vbCrLf & " In Item Master Or Customer Item RelationShip " & vbCrLf & validate_itemcode(), MsgBoxStyle.Information, ResolveResString(100))
                    '    Exit Sub
                    'End If
                    If validate_holiday() <> "" Then
                        MsgBox("Following Is\Are Not Working Day(s) " & vbCrLf & validate_holiday() & vbCrLf, MsgBoxStyle.Information, ResolveResString(100))
                        Exit Sub
                    End If
                    Call SetSpreadProperty()
                    fillGrid()
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
    Private Sub delete_previous_data(ByVal strDate As Date, ByVal stritemcode As String, ByVal LngQty As Integer, ByVal stmt As String)
        On Error GoTo ErrHandler
        Dim strSQL As String
        Dim objgetData As New ClsResultSetDB
        Dim objcheckdata As New ClsResultSetDB
        Dim prevquantity As Integer
        Dim strupdate As String
        objcheckdata.GetResult("select * from forecast_mst where customer_code='" & txtcustomercode.Text & "' and product_no='" & stritemcode & "' and due_date='" & strDate & "' and Unit_code = '" & gstrUNITID & "'")
        If Not objcheckdata.EOFRecord Then
            prevquantity = objcheckdata.GetValue("quantity")
            If prevquantity <> LngQty Then
                mflag = 1
                strupdate = "update forecast_mst_history set revisionno=revisionno+1 where customer_code='" & txtcustomercode.Text & "' and product_no='" & stritemcode & "' and due_date='" & strDate & "' and Unit_code = '" & gstrUNITID & "'"
                mP_Connection.Execute(strupdate, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                strSQL = "insert into forecast_mst_history(Customer_code,product_no,Due_date,quantity,ScheduleNo,RevisionNo,ent_userid,ent_dt,upd_userid,upd_dt,Enagare_UNLOC,Unit_Code,Cust_Drgno)" & " select customer_code,product_no,due_Date,quantity,1,0,ent_userid ,ent_dt,upd_userid,upd_dt,enagare_unloc,Unit_code,Cust_Drgno from forecast_mst where customer_code='" & Trim(txtcustomercode.Text) & "' and due_date ='" & strDate & "' and product_no='" & stritemcode & "' and Unit_code = '" & gstrUNITID & "'"
                mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                strSQL = "delete from forecast_mst where customer_code='" & Trim(txtcustomercode.Text) & "' and due_date ='" & strDate & "' and product_no='" & stritemcode & "' and Unit_code = '" & gstrUNITID & "'"
                mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                mP_Connection.Execute(stmt, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
        Else
            mflag = 1
            mP_Connection.Execute(stmt, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        End If
        objgetData = Nothing
        objcheckdata = Nothing
        Exit Sub
ErrHandler:
        objgetData = Nothing
        objcheckdata = Nothing
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
            .MaxCols = enmforecastitem.Cust_DrgwNo
            .Col = enmforecastitem.status : .Text = "Select" : .set_ColWidth(enmforecastitem.status, 600)
            .Col = enmforecastitem.ItemCode : .Text = "Item Code" : .set_ColWidth(enmforecastitem.ItemCode, 1600)
            .Col = enmforecastitem.Item_Description : .Text = "Item Description" : .set_ColWidth(enmforecastitem.Item_Description, 2700)
            .Col = enmforecastitem.Cust_DrgwNo : .Text = "Cust Drgn No" : .set_ColWidth(enmforecastitem.Cust_DrgwNo, 2000)
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
            .Col = enmforecastitem.status : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmforecastitem.ItemCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmforecastitem.Item_Description : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmforecastitem.Cust_DrgwNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
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
        getfilename = "Forecast_" & strcustname & "_" & strFileName & ".xls"
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
        Dim itemcount As Long
        Dim intcounter1 As Long
        Dim interval As Long
        Dim dtDate As Date
        Dim datecounter As Long
        Dim objcalender As New ClsResultSetDB
        interval = DateDiff(Microsoft.VisualBasic.DateInterval.Day, dtfromdate.Value, dttodate.Value) + 1
        With fspitem
            For intcounter = 1 To .MaxRows
                varstatus = Nothing
                varItemCode = Nothing
                varCustdrgno = Nothing
                Call .GetText(enmforecastitem.status, intcounter, varstatus)
                varstatus = IIf(varstatus.ToString = "", 0, varstatus)
                If varstatus = 1 Then
                    Call .GetText(enmforecastitem.ItemCode, intcounter, varItemCode)
                    Call .GetText(enmforecastitem.Cust_DrgwNo, intcounter, varCustdrgno)
                    itemcount = itemcount + 1
                    objexl.Cells._Default(6, 1 + itemcount) = varItemCode
                    objexl.Cells._Default(5, 1 + itemcount) = varCustdrgno
                    objexl.Cells._Default(4, 1 + itemcount) = get_itemdescription(varItemCode)
                    objexl.Cells._Default(4, 1 + itemcount).Font.Bold = True
                    objexl.Cells._Default(4, 1 + itemcount).ColumnWidth = Len(get_itemdescription(varItemCode)) + 5
                    objexl.Cells._Default(5, 1 + itemcount).Font.Bold = True
                    objexl.Cells._Default(6, 1 + itemcount).Font.Bold = True
                End If
            Next
        End With
        objcalender.GetResult("select dt from Calendar_mst where dt between '" & getDateForDB(dtfromdate.Value) & "' and '" & getDateForDB(dttodate.Value) & "' and work_flg=0 and Unit_code = '" & gstrUNITID & "' order by dt")
        If Not objcalender.EOFRecord Then
            While Not objcalender.EOFRecord
                datecounter = datecounter + 1
                objexl.Cells._Default(6 + datecounter, 1) = VB6.Format(objcalender.GetValue("dt"), "DD MMM YYYY")
                objexl.Cells._Default(6 + datecounter, 1).Font.Bold = True
                objexl.Cells._Default(6 + datecounter, 1).ColumnWidth = 15
                objcalender.MoveNext()
            End While
        End If
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
                Call .GetText(enmforecastitem.status, intcounter, varstatus)
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
      
        validate_excel_format = True
        row1 = 7 : row2 = 7
        row = 6 : col = 2
        stritemcode = Nothing
        stritemcode = objexl.Cells(row, col).Value
        If stritemcode = "" Then
            validate_excel_format = False
            Exit Function
        End If
       
        While (col <= objexl.Columns.Count)

            row = 6 : col = col
            stritemcode = objexl.Cells(row, col).Value
            rowDrgNo = 5 : col = col
            strCustDrgNo = objexl.Cells(rowDrgNo, col).Value
            rowItemDesc = 4 : col = col
            strItemDesc = objexl.Cells(rowItemDesc, col).Value

            row1 = 7 : col1 = 1

            If stritemcode <> "" AndAlso strCustDrgNo <> "" Then
                If IsDate(objexl.Cells(row1, col1).value) = True Then
                    validate_excel_format = True
                    strDate = objexl.Cells(row1, col1).Value
                Else
                    validate_excel_format = False
                    Exit Function
                End If
                row2 = 7 : col2 = 2
                If IsNumeric(Val(IIf(objexl.Cells(row2, col2).ToString = "", 0, objexl.Cells(row2, col2).ToString))) = True Then
                    validate_excel_format = True
                Else
                    validate_excel_format = False
                    Exit Function
                End If

                If stritemcode <> "" Then
                    While strDate <> ""
                        row1 = row1 : col1 = 1
                        strDate = objexl.Cells(row1, col1).Value
                        If strDate <> "" Then
                            If IsDate(objexl.Cells(row1, col1).Value) = True Then
                                validate_excel_format = True
                            Else
                                validate_excel_format = False
                                Exit Function
                            End If
                            If IsNumeric(Val(IIf(objexl.Cells(row2, col2).ToString = "", 0, objexl.Cells(row2, col2).ToString))) = True Then
                                validate_excel_format = True
                            Else
                                validate_excel_format = False
                                Exit Function
                            End If
                        End If
                        row1 = row1 + 1
                        row2 = row2 + 1
                        rowDrgNo = rowDrgNo + 1
                    End While
                End If
                col = col + 1
            Else
                If Not (stritemcode = "" And strCustDrgNo = "") Then
                    If stritemcode = "" Then
                        MsgBox("Item Code is blank at Row No " & row & " and Part Code " & strCustDrgNo & " ", MsgBoxStyle.Information, Me.Text)
                    ElseIf strCustDrgNo = "" Then
                        MsgBox("Part Code is blank at Item Code " & stritemcode & " and Part Code " & col & "", MsgBoxStyle.Information, Me.Text)
                    End If
                    validate_excel_format = False
                    col = col + 1
                    Exit Function
                Else
                    If (stritemcode = "" And strItemDesc = "" And stritemcode = "") Then
                        Exit Function
                    Else
                        MsgBox("Item Code and Part Code is blank at Row No " & row & " and Column No " & col & " ", MsgBoxStyle.Information, Me.Text)
                        validate_excel_format = False
                        col = col + 1
                        Exit Function
                    End If

                End If
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
        Dim objfso As New Scripting.FileSystemObject
        If fspforecastdetail.MaxRows > 0 Then
            If txtcustomercode.Text <> "" Then
                'stritemlist = validate_previous_data
                If mP_Connection.State = ADODB.ObjectStateEnum.adStateOpen Then
                    mP_Connection.Close()
                    mP_Connection.Open()
                End If
                mP_Connection.BeginTrans()
                intvalue = 1
                save_data()
                If objfso.FolderExists(gstrLocalCDrive & "forecast\uploaded_files") = True Then
                    If objfso.FileExists(gstrLocalCDrive & "forecast\uploaded_files\" + strUploadedFileName) Then
                        objfso.DeleteFile(gstrLocalCDrive & "forecast\uploaded_files\" + strUploadedFileName)
                    End If
                    objfso.MoveFile(txtfilelocation.Text, gstrLocalCDrive & "forecast\uploaded_files\")
                Else
                    objfso.CreateFolder(gstrLocalCDrive & "forecast\uploaded_files")
                    objfso.MoveFile(txtfilelocation.Text, gstrLocalCDrive & "forecast\uploaded_files\")
                End If
                mP_Connection.CommitTrans()
                If mflag = 1 Then
                    MsgBox("Transaction Saved Successfully", MsgBoxStyle.Information, ResolveResString(100))
                    mflag = 0
                Else
                    MsgBox("File Already Uploaded", MsgBoxStyle.Information, ResolveResString(100))
                    RefreshForm()
                    Exit Sub
                End If
                RefreshForm()
            Else
                MsgBox("Customer Code Can Not Be Blank", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
        Else
            MsgBox("Please Upload Forecast First", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If
        objfso = Nothing
        Exit Sub
ErrHandler:
        If intvalue = 1 Then mP_Connection.RollbackTrans()
        objfso = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
End Class