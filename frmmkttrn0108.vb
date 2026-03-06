Option Strict Off
Option Explicit On
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient

'----------------------------------------------------
'Copyright (c)  -  MIND
'Name of module -  frmMKTTRN0108.vb
'Created By     -  MILIND MISHRA    
'Created Date   -  15 Nov 2018
'Description    -  Generic Marketing Schedule Uploading
'-----------------------------------------------------------------------------

Friend Class frmMKTTRN0108
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
        Try

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

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub SetSpreadColTypes(ByRef pintRowNo As Long)
        Try
            Dim i As Long
            With Me.fspforecastdetail
                .Row = pintRowNo
                .Col = enmforecastdetail.ItemCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                .Col = enmforecastdetail.Item_Description : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                .Col = enmforecastdetail.Cust_DrgwNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                .Col = enmforecastdetail.UOM : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                .Col = enmforecastdetail.Quantity : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                .Col = enmforecastdetail.DELIVERY_DATE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeDate : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = True
            End With
            fspforecastdetail.ReDraw = True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub addRowAtEnterKeyPress(ByRef pintRows As Long)
        Try
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

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub RefreshForm()
        Try
            fspforecastdetail.MaxCols = 0
            fspforecastdetail.MaxRows = 0
            txtcustomercode.Text = ""
            lblcustomername.Text = ""
            txtfilelocation.Text = ""
            TxtSearch.Text = ""
            CmbSearchBy.SelectedIndex = -1

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub chkall_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkall.CheckStateChanged
        Try
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

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub CmbSearchBy_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbSearchBy.TextChanged
        Try
            If TxtSearch.Enabled Then
                TxtSearch.Focus()
            End If
            Exit Sub

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub CmbSearchBy_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbSearchBy.SelectedIndexChanged
        Try
            If TxtSearch.Enabled Then
                TxtSearch.Focus()
            End If
            Exit Sub

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub cmdCustHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdcusthelp.Click
        Dim strSQL As String = String.Empty
        Dim strHelp() As String
        Try
            strSQL = "SELECT ACCOUNT_CODE ,CUST_NAME FROM VW_GENERATE_SCH_CUSTOMER WHERE UNIT_CODE ='" & gstrUNITID & "'"
            strHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL)

            If IsNothing(strHelp) = True Then Exit Sub
            If strHelp.GetUpperBound(0) <> -1 Then
                If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                    MsgBox("No Record found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                    Exit Sub
                Else
                    txtcustomerhelp.Text = strHelp(0)
                    lblcustname.Text = strHelp(1)
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
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
            If txtfilelocation.Text.Trim.Length = 0 Then
                Exit Sub
            End If

            strUploadedFileName = Mid(CommanDLogOpen.FileName, CommanDLogOpen.FileName.LastIndexOf("\") + 2, CommanDLogOpen.FileName.Length - 1)
            strUploadingpath = gstrLocalCDrive & "Firm Schedule\uploaded_files"

        Catch ex As Exception
            MessageBox.Show(ex.Message)
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
                If objfso.FolderExists(gstrLocalCDrive & "Firm Schedule") = True Then
                    objfolder = objfso.GetFolder(gstrLocalCDrive & "Firm Schedule")
                    For Each objfile In objfolder.Files
                        stroldfilename = objfile.Name
                        If UCase(Mid(stroldfilename, 10, 8)) = UCase(txtcustomerhelp.Text) Then
                            If objfso.FolderExists(gstrLocalCDrive & "Firm Schedule\Backup") = True Then
                                If stroldfilename <> "" Then
                                    objfso.MoveFile(gstrLocalCDrive & "Firm Schedule\" & stroldfilename, gstrLocalCDrive & "Firm Schedule\backup\")
                                End If
                            Else
                                objfso.CreateFolder(gstrLocalCDrive & "Firm Schedule\backup")
                                If stroldfilename <> "" Then
                                    objfso.MoveFile(gstrLocalCDrive & "Firm Schedule\" & stroldfilename, gstrLocalCDrive & "Firm Schedule\backup\")
                                End If
                            End If
                        End If
                    Next objfile
                    OBj_wB = OBJexlformat.Workbooks.Add
                    strnewfilename = getfilename(Trim(txtcustomerhelp.Text))
                    OBj_wB.SaveAs(gstrLocalCDrive & "Firm Schedule\" & Replace(strnewfilename, ":", ""))
                    OBJexlformat.Workbooks.Close()
                    OBj_wB = Nothing
                    OBJexlformat.Workbooks.Open(gstrLocalCDrive & "Firm Schedule\" & Replace(strnewfilename, ":", ""))
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
                    MsgBox("Excel Sheet Generated Successfully " & vbCrLf & "location--> " & " " & gstrLocalCDrive & "Firm Schedule\" & Replace(strnewfilename, ":", ""), MsgBoxStyle.Information, ResolveResString(100))
                    clear_fields()
                    chkall.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkall.Enabled = False
                    OBJexlformat = Nothing
                Else
                    objfso.CreateFolder(gstrLocalCDrive & "Firm Schedule")
                    OBj_wB = OBJexlformat.Workbooks.Add
                    strnewfilename = getfilename(Trim(txtcustomerhelp.Text))
                    OBj_wB.SaveAs(gstrLocalCDrive & "Firm Schedule\" & Replace(strnewfilename, ":", ""))
                    OBJexlformat.Workbooks.Close()
                    OBj_wB = Nothing
                    OBJexlformat.Workbooks.Open(gstrLocalCDrive & "Firm Schedule\" & Replace(strnewfilename, ":", ""))
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
                    MsgBox("Excel Sheet Generated Successfully " & vbCrLf & "location--> " & " " & gstrLocalCDrive & "Firm Schedule\" & Replace(strnewfilename, ":", ""), MsgBoxStyle.Information, ResolveResString(100))
                    clear_fields()
                    chkall.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkall.Enabled = False
                    OBJexlformat = Nothing
                End If
            Else
                MsgBox("No Item Is Selected", MsgBoxStyle.Information, ResolveResString(100))
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub cmdclose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdclose.Click
        Try
            Me.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub cmdupload_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdupload.Click
        Try

            If Not optGeneric.Checked = True Then
                MessageBox.Show("Select Generic option to upload file.")
                Exit Sub
            End If

            If txtcustomercode.Text <> "" Then
                If txtfilelocation.Text <> "" Then
                    fspforecastdetail.MaxRows = 0
                    If optGeneric.Checked = True Then
                        Call Upload_Generic_Forecast_Data()
                    Else
                        Call Upload_Toyota_Forecast_Data()
                    End If
                Else
                    MsgBox("Please Select File Location To Upload Schedule", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If
            Else
                MsgBox("Please Select Customer First", MsgBoxStyle.Information, ResolveResString(100))
                txtcustomercode.Focus()
                Exit Sub
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub cmdviewitems_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdviewitems.Click
        Try
            Dim oDr As SqlDataReader

            Dim strSQL As String
            If Len(Trim(txtcustomerhelp.Text)) = 0 Then
                MsgBox("Please Select Customer Code First", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
            strSQL = "select distinct a.item_code,b.description,a.cust_drgno from custitem_mst" & " a inner join item_mst b on a.item_code=b.item_code and a.Unit_Code=b.Unit_Code and a.active=1 and b.status='A'" & " and a.account_code='" & Trim(txtcustomerhelp.Text) & "' and a.Unit_code = '" & gstrUNITID & "' "
            oDr = SqlConnectionclass.ExecuteReader(strSQL)

            If oDr.HasRows Then
                Call setitemgrdproperty()
                While oDr.Read
                    With fspitem
                        addrowinitemgrd((1))
                        Call .SetText(enmforecastitem.ItemCode, .MaxRows, oDr("item_code"))
                        Call .SetText(enmforecastitem.Item_Description, .MaxRows, oDr("description"))
                        Call .SetText(enmforecastitem.Cust_DrgwNo, .MaxRows, oDr("cust_drgno"))
                    End With
                End While
                chkall.Enabled = True
                cmdgenerateformat.Enabled = True
            Else
                MsgBox("No Item Is Defined For The Specified Customer", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub frmMKTTRN0108_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Try

            mdifrmMain.CheckFormName = mintIndex
            frmModules.NodeFontBold(Tag) = True
            optGeneric.Checked = True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub frmMKTTRN0108_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        Try

            frmModules.NodeFontBold(Me.Tag) = False

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub frmMKTTRN0108_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Try

            mintIndex = mdifrmMain.AddFormNameToWindowList(ctlGenericSchedule.Text)
            FitToClient(Me, framain, ctlGenericSchedule, frabuttons)
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
            '' lblMsg.Text = ""
            chkPicklist.Checked = True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub frmMKTTRN0108_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try

            frmModules.NodeFontBold(Me.Tag) = False
            mdifrmMain.RemoveFormNameFromWindowList = mintIndex

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub cmdcustomercode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdcustomercode.Click
        Try

            Dim strCustHelp() As String

            chkPicklist.Checked = True

            If Len(Me.txtcustomercode.Text) = 0 Then
                strCustHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT ACCOUNT_CODE , CUST_NAME FROM VW_GENERATE_SCH_CUSTOMER WHERE UNIT_CODE ='" & gstrUNITID & "'", "Customer Codes List", 1)
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

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub SSTab1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SSTab1.SelectedIndexChanged
        Try
            If SSTab1.SelectedIndex = 0 Then
                cmdsave.Enabled = False
            Else
                chkPicklist.Checked = True
                cmdsave.Enabled = True
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub txtCustomerCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtcustomercode.TextChanged
        Try
            lblcustomername.Text = ""
            fspforecastdetail.MaxCols = 0
            fspforecastdetail.MaxRows = 0
            txtfilelocation.Text = ""

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub txtCustomerCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtcustomercode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Try
            If KeyCode = System.Windows.Forms.Keys.F1 Then
                Call cmdcustomercode_Click(cmdcustomercode, New System.EventArgs())
                chkPicklist.Checked = True
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub TxtCustomerCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtcustomercode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Try
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

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            eventArgs.KeyChar = Chr(KeyAscii)
            If KeyAscii = 0 Then
                eventArgs.Handled = True
            End If
        End Try
    End Sub

    Private Sub TxtCustomerCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtcustomercode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Try
            Dim strSQL As String
            Dim oDr As SqlDataReader

            If Len(txtcustomercode.Text) > 0 Then
                txtcustomercode.Text = Replace(txtcustomercode.Text, "'", "")
                strSQL = "select a.customer_Code, a.cust_name from Customer_Mst a where a.customer_code='" & Trim(txtcustomercode.Text) & "' and a.Unit_code = '" & gstrUNITID & "' and ((isnull(a.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= a.deactive_date))"
                oDr = SqlConnectionclass.ExecuteReader(strSQL)
                If oDr.HasRows Then
                    oDr.Read()
                    lblcustomername.Text = oDr("cust_name").ToString
                Else
                    MsgBox("Invalid Customer Code", MsgBoxStyle.Information, ResolveResString(100))
                    lblcustomername.Text = ""
                    txtcustomercode.Text = ""
                    txtcustomercode.Enabled = True
                    txtcustomercode.Focus()
                End If
            Else
                lblcustomername.Text = ""
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            eventArgs.Cancel = Cancel
        End Try
    End Sub

    Private Sub fillGrid()
        Try
            Dim strSQL As String
            Dim intRowCount As Long
            Dim intcounter As Long
            Dim intMaxCounter As Long

            Using sqlCmd As SqlCommand = New SqlCommand
                With sqlCmd
                    .Connection = SqlConnectionclass.GetConnection()
                    .CommandType = CommandType.Text
                    .CommandText = "select product_no,due_date,sum(quantity)as quantity from SCHEDULE_MST_TEMP where customer_code='" & Trim(txtcustomercode.Text) & "' and ipaddress='" & gstrIpaddressWinSck & "' and Unit_code = '" & gstrUNITID & "' and due_date>=convert(varchar(12),getdate(),106) group by product_no,due_date"
                    Using dt As DataTable = SqlConnectionclass.GetDataTable(sqlCmd)
                        If dt.Rows.Count > 0 Then
                            With fspforecastdetail
                                .MaxRows = 0
                                For Each row As DataRow In dt.Rows
                                    Call addRowAtEnterKeyPress(1)
                                    .Row = .MaxRows
                                    .Col = enmforecastdetail.ItemCode : Call .SetText(enmforecastdetail.ItemCode, .MaxRows, Convert.ToString(row("product_no")))
                                    .Col = enmforecastdetail.Item_Description
                                    Call .SetText(enmforecastdetail.Item_Description, .MaxRows, get_itemdescription(Convert.ToString(row("product_no"))))
                                    .Col = enmforecastdetail.Cust_DrgwNo
                                    Call .SetText(enmforecastdetail.Cust_DrgwNo, .MaxRows, get_customer_drgno(Convert.ToString(row("product_no"))))
                                    .Col = enmforecastdetail.UOM
                                    Call .SetText(enmforecastdetail.UOM, .MaxRows, get_item_uom(Convert.ToString(row("product_no"))))
                                    .Col = enmforecastdetail.DELIVERY_DATE : Call .SetText(enmforecastdetail.DELIVERY_DATE, .MaxRows, VB6.Format(Convert.ToString(row("due_date")), gstrDateFormat))
                                    .Col = enmforecastdetail.Quantity : Call .SetText(enmforecastdetail.Quantity, .MaxRows, Convert.ToString(row("quantity")))
                                Next
                            End With
                        Else
                            MsgBox("No record found !", MsgBoxStyle.Exclamation, ResolveResString(100))
                            Return
                        End If
                    End Using
                End With
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub SearchItem()
        Try
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

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub txtcustomerhelp_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtcustomerhelp.TextChanged
        Try
            lblcustname.Text = ""
            fspitem.MaxCols = 0
            fspitem.MaxRows = 0
            chkall.Enabled = False
            cmdgenerateformat.Enabled = False

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub txtcustomerhelp_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtcustomerhelp.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Try
            If KeyCode = System.Windows.Forms.Keys.F1 Then
                Call cmdCustHelp_Click(cmdcusthelp, New System.EventArgs())
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub txtcustomerhelp_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtcustomerhelp.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Try
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

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            eventArgs.KeyChar = Chr(KeyAscii)
            If KeyAscii = 0 Then
                eventArgs.Handled = True
            End If
        End Try
    End Sub

    Private Sub txtcustomerhelp_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtcustomerhelp.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Try
            '  Dim oRs As ADODB.Recordset
            Dim strSQL As String
            Dim oDr As SqlDataReader

            If Len(txtcustomerhelp.Text) > 0 Then
                txtcustomerhelp.Text = Replace(txtcustomerhelp.Text, "'", "")
                strSQL = "select a.customer_Code, a.cust_name from Customer_Mst a where a.customer_code='" & Trim(txtcustomerhelp.Text) & "' and a.Unit_code = '" & gstrUNITID & "' and ((isnull(a.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= a.deactive_date))"
                oDr = SqlConnectionclass.ExecuteReader(strSQL)

                If oDr.HasRows Then
                    lblcustname.Text = oDr("cust_name").Value
                Else
                    MsgBox("Invalid Customer Code", MsgBoxStyle.Information, ResolveResString(100))
                    lblcustname.Text = ""
                    txtcustomerhelp.Text = ""
                    txtcustomerhelp.Enabled = True
                    txtcustomerhelp.Focus()
                End If
            Else
                lblcustname.Text = ""
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            eventArgs.Cancel = Cancel
        End Try
    End Sub

    Private Sub TxtSearch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtSearch.TextChanged
        Try
            Call SearchItem()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub save_data()
        Dim strInsert As String
        Dim intcounter As Long
        Dim varduedate As Object
        Dim varproductno As Object
        Dim varquantity As Object
        Dim varDrgNo As Object
        Dim intDocNo As Integer
        Dim oTrans As SqlTransaction

        Try

            If chkPicklist.Checked = True And txtfilelocation.Text.Trim.Length = 0 Then
                If txtDocNo.Text.Length = 0 Then
                    MessageBox.Show("Upload schedue file or enter existing doc no to generate picklist.", ResolveResString(100), MessageBoxButtons.OK)
                    Exit Sub
                End If
                If txtDocNo.Text.Length > 0 Then
                    createPicklist(txtDocNo.Text)
                End If
            End If

            Using sqlcmd As New SqlCommand
                With sqlcmd
                    .CommandText = "GENERATE_GENERIC_DOC_NO"
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 0
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@DOC_TYPE", SqlDbType.Int).Value = 113
                    .Parameters.Add("@DOC_NO", SqlDbType.Int).Direction = ParameterDirection.Output
                    SqlConnectionclass.ExecuteNonQuery(sqlcmd)
                    intDocNo = .Parameters("@DOC_NO").Value
                End With
            End Using

            With fspforecastdetail
                For intcounter = 1 To fspforecastdetail.MaxRows
                    varproductno = Nothing
                    varduedate = Nothing
                    varquantity = Nothing
                    varDrgNo = Nothing
                    Call .GetText(enmforecastdetail.ItemCode, intcounter, varproductno)
                    Call .GetText(enmforecastdetail.DELIVERY_DATE, intcounter, varduedate)
                    Call .GetText(enmforecastdetail.Quantity, intcounter, varquantity)
                    Call .GetText(enmforecastdetail.Cust_DrgwNo, intcounter, varDrgNo)

                    Using sqlcmd As New SqlCommand
                        With sqlcmd
                            .CommandText = "USP_SAVE_MARKETING_SCHEDULE_UPLOAD"
                            .CommandType = CommandType.StoredProcedure
                            .CommandTimeout = 0
                            .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                            .Parameters.Add("@CUSTOMER_CODE", SqlDbType.VarChar, 20).Value = txtcustomercode.Text
                            .Parameters.Add("@ITEM_CODE", SqlDbType.VarChar, 16).Value = varproductno
                            .Parameters.Add("@DOC_NO", SqlDbType.Int).Value = intDocNo
                            .Parameters.Add("@CUST_DRGNO", SqlDbType.VarChar, 50).Value = varDrgNo
                            .Parameters.Add("@DUE_DATE", SqlDbType.Date).Value = varduedate
                            .Parameters.Add("@SCHEDULE_QTY", SqlDbType.Money).Value = varquantity
                            .Parameters.Add("@USER_ID", SqlDbType.VarChar, 20).Value = mP_User
                            SqlConnectionclass.ExecuteNonQuery(sqlcmd)
                        End With
                    End Using

                Next
            End With

            ''        lblMsg.Text = " S c h e d u l e    S a v e d "

            mflag = 1

            If chkPicklist.Checked = True Then
                Call createPicklist(intDocNo)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub Save_Data_Toyota(ByVal intDocNo As Integer, ByVal strPONO As String)
        Try

            Using sqlcmd As New SqlCommand
                With sqlcmd
                    .CommandText = "USP_SAVE_TOYOTA_MARKETING_SCHEDULE"
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 0

                    .Parameters.AddWithValue("@UNITCODE", gstrUNITID)
                    .Parameters.AddWithValue("@ipaddress", gstrIpaddressWinSck)
                    .Parameters.AddWithValue("@USERID", mP_User)
                    .Parameters.AddWithValue("@DocNo", intDocNo)
                    .Parameters.Add("@ErrMSg", SqlDbType.VarChar, 200).Direction = ParameterDirection.Output

                    SqlConnectionclass.ExecuteNonQuery(sqlcmd)
                End With
            End Using

            If chkPicklist.Checked = True Then
                Call createPicklist_TOYOTA(intDocNo)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub createPicklist(ByVal intDocNo As Integer)
        Dim strMSG As String

        Try

            If txtcustomercode.Text = "" Then
                MessageBox.Show("Select Customer Code", ResolveResString(100), MessageBoxButtons.OK)
                Exit Sub
            End If

            Using sqlcmd As New SqlCommand
                With sqlcmd
                    .CommandText = "USP_Picklist_For_Generic_Schedule"
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 0
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@Customer_Code", SqlDbType.VarChar).Value = txtcustomercode.Text
                    .Parameters.Add("@DOC_NO", SqlDbType.Int).Value = intDocNo

                    If optGeneric.Checked = True Then
                        .Parameters.AddWithValue("@SchType", SqlDbType.VarChar).Value = "GENERIC"
                    ElseIf optToyotaReleaseFile.Checked = True Then
                        .Parameters.AddWithValue("@SchType", SqlDbType.VarChar).Value = "TOYOTA"
                    End If

                    .Parameters.Add("@RETMSG", SqlDbType.VarChar, 2000).Direction = ParameterDirection.Output
                    SqlConnectionclass.ExecuteNonQuery(sqlcmd)

                    strMSG = .Parameters("@RETMSG").Value
                    If strMSG.ToString.Length > 0 Then
                        MessageBox.Show(strMSG, ResolveResString(100), MessageBoxButtons.OK)
                    Else
                        MessageBox.Show("Picklist Generated.", ResolveResString(100), MessageBoxButtons.OK)
                        ''      lblMsg.Text = lblMsg.Text + ",     " + "P i c k l i s t   G e n e r a t e d"
                    End If

                End With
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub createPicklist_TOYOTA(ByVal intDocNo As Integer)
        Dim strMSG As String

        Try

            If txtcustomercode.Text = "" Then
                MessageBox.Show("Select Customer Code", ResolveResString(100), MessageBoxButtons.OK)
                Exit Sub
            End If

            Using sqlcmd As New SqlCommand
                With sqlcmd
                    .CommandText = "USP_Picklist_For_TOYOTA_Schedule"
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 0
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@Customer_Code", SqlDbType.VarChar).Value = txtcustomercode.Text
                    .Parameters.Add("@DOC_NO", SqlDbType.Int).Value = intDocNo
                    .Parameters.Add("@RETMSG", SqlDbType.VarChar, 2000).Direction = ParameterDirection.Output
                    SqlConnectionclass.ExecuteNonQuery(sqlcmd)

                    strMSG = .Parameters("@RETMSG").Value
                    If strMSG.ToString.Length > 0 Then
                        MessageBox.Show(strMSG, ResolveResString(100), MessageBoxButtons.OK)
                    Else
                        MessageBox.Show("Picklist Generated.", ResolveResString(100), MessageBoxButtons.OK)
                        ''      lblMsg.Text = lblMsg.Text + ",     " + "P i c k l i s t   G e n e r a t e d"
                    End If

                End With
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Function get_itemdescription(ByVal stritemcode As String) As String
        Try
            Dim strSQL As String
            Dim oDr As SqlDataReader

            strSQL = "select description from item_mst where item_code='" & stritemcode & "' and Unit_code = '" & gstrUNITID & "'"
            oDr = SqlConnectionclass.ExecuteReader(strSQL)
            If oDr.HasRows Then
                oDr.Read()
                get_itemdescription = oDr("description")
            Else
                get_itemdescription = ""
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Function get_item_uom(ByRef stritemcode As String) As String
        Try
            Dim oDr As SqlDataReader
            Dim strSQL As String

            strSQL = "select cons_measure_code from item_mst where item_Code='" & stritemcode & "' and Unit_code = '" & gstrUNITID & "'"
            oDr = SqlConnectionclass.ExecuteReader(strSQL)
            If oDr.HasRows Then
                oDr.Read()
                get_item_uom = oDr("cons_measure_Code")
            Else
                get_item_uom = ""
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Function get_customer_drgno(ByRef stritemcode As String) As String
        Try
            Dim strSQL As String
            Dim oDr As SqlDataReader

            If SSTab1.SelectedIndex = 0 Then
                strSQL = "select top 1 cust_drgno from custitem_mst where account_code='" & Trim(txtcustomerhelp.Text) & "' and item_code='" & stritemcode & "' and active=1 and Unit_code = '" & gstrUNITID & "'"
            Else
                strSQL = "select top 1 cust_drgno from custitem_mst where account_code='" & Trim(txtcustomercode.Text) & "' and item_code='" & stritemcode & "' and active=1 and Unit_code = '" & gstrUNITID & "'"
            End If
            oDr = SqlConnectionclass.ExecuteReader(strSQL)
            If oDr.HasRows Then
                oDr.Read()
                get_customer_drgno = oDr("cust_drgno")
            Else
                get_customer_drgno = ""
            End If

            Exit Function

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Sub Upload_Generic_Forecast_Data()

        Dim objexl As New Excel.Application
        Dim objfso As New Scripting.FileSystemObject
        Dim row As Long
        Dim strcustomercode As String
        Dim stritemcode As String
        Dim strDate As String
        Dim strQuantity As Decimal
        Dim col As Long
        Dim row1 As Long
        Dim col1 As Long
        Dim row2 As Long
        Dim col2 As Long
        Dim strSQL As String

        Try

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
                        mP_Connection.Execute("delete from SCHEDULE_MST_TEMP where customer_code='" & txtcustomercode.Text & "' and ipaddress='" & gstrIpaddressWinSck & "' and Unit_code = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        row1 = 7 : row2 = 7
                        row = 6 : col = 2
                        stritemcode = objexl.Cells(row, col).Value
                        While (stritemcode <> "" And col <= objexl.Columns.Count)
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
                                        strSQL = "insert into SCHEDULE_MST_TEMP(Customer_code,product_no,Due_date,quantity,ent_userid,ent_dt,upd_userid,upd_dt,ipaddress,Unit_code)" & " values('" & txtcustomercode.Text & "','" & stritemcode.Trim & "','" & getDateForDB(VB6.Format(strDate, gstrDateFormat)) & "'," & CInt(strQuantity) & ",'" & mP_User & "',getdate(),'" & mP_User & "',getdate(),'" & gstrIpaddressWinSck & "', '" & gstrUNITID & "')"
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
                        If validate_itemcode() <> "" Then
                            MsgBox("Following Items Are Not Defined Or Inactive " & vbCrLf & " In Item Master Or Customer Item RelationShip " & vbCrLf & validate_itemcode(), MsgBoxStyle.Information, ResolveResString(100))
                            Exit Sub
                        End If

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

            objexl = Nothing
            objfso = Nothing

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            If Not objexl Is Nothing Then
                KillExcelProcess(objexl)
                objexl = Nothing
            End If
        End Try

    End Sub

    Private Sub Upload_Toyota_Forecast_Data()

        Dim objexl As New Excel.Application
        Dim objfso As New Scripting.FileSystemObject
        Dim row As Long
        Dim strcustomercode As String
        Dim stritemcode As String
        Dim strDate As String
        Dim strPONO As String
        Dim col As Long
        Dim intDocNo As Integer
        Dim strSQL As String

        Try

            If chkPicklist.Checked = True And txtfilelocation.Text.Trim.Length = 0 Then
                If txtDocNo.Text.Length = 0 Then
                    MessageBox.Show("Upload schedue file or enter existing doc no to generate picklist.", ResolveResString(100), MessageBoxButtons.OK)
                    Exit Sub
                End If
                If txtDocNo.Text.Length > 0 Then
                    createPicklist(txtDocNo.Text)
                End If
            End If

            If objfso.FileExists(txtfilelocation.Text) = False Then
                MsgBox("File Does Not Exists", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            Else

                'SqlConnectionclass.BeginTrans()

                objexl = New Excel.Application
                objexl.Workbooks.Open(Trim(Me.txtfilelocation.Text))
                If validate_Toyota_format(objexl) = True Then
                    strPONO = objexl.Cells(2, 5).Value
                    strSQL = "select TOP 1 1 from ToyotaSchedule_ReleaseNoWise where UNITCODE = '" + gStrUnitId + "' AND PONo = '" + strPONO + "'"

                    If IsRecordExists(strSQL) Then
                        strSQL = "update ToyotaSchedule_ReleaseNoWise set pono = pono + Doc_No where UNITCODE = '" + gstrUNITID + "' AND PONO = '" + strPONO + "'"
                        SqlConnectionclass.ExecuteNonQuery(strSQL)

                        strSQL = "update DAILYTOYOTASCHEDULE set pono = pono + cast(Doc_No as varchar(10)) where UNIT_CODE = '" + gstrUNITID + "' AND PONO = '" + strPONO + "'"
                        SqlConnectionclass.ExecuteNonQuery(strSQL)
                    End If

                    row = 2
                    strPONO = objexl.Cells(row, 5).Value

                    Using sqlcmd As New SqlCommand
                        With sqlcmd
                            .CommandText = "GENERATE_GENERIC_DOC_NO"
                            .CommandType = CommandType.StoredProcedure
                            .CommandTimeout = 0
                            .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                            .Parameters.Add("@DOC_TYPE", SqlDbType.Int).Value = 113
                            .Parameters.Add("@DOC_NO", SqlDbType.Int).Direction = ParameterDirection.Output
                            SqlConnectionclass.ExecuteNonQuery(sqlcmd)
                            intDocNo = .Parameters("@DOC_NO").Value
                        End With
                    End Using

                    strSQL = "DELETE FROM ToyotaSchedule_ReleaseNoWise WHERE UNITCODE='" & gStrUnitId & "' AND IP_Address='" & gstrIpaddressWinSck & "'"
                    SqlConnectionclass.ExecuteNonQuery(strSQL)
                    While (strPONO <> "" And col <= objexl.Columns.Count)
                        strSQL = "Set Dateformat 'dmy' insert into ToyotaSchedule_ReleaseNoWise(UnitCode,Doc_No,CustomerCode,DepotCode,SupplierCode,SupplierName,FactoryCode,"
                        strSQL = strSQL + " PONo,POItemNo,MakerCode,POPartNo,PartName,OrderType,SupplierPartNo,POQty,PODate,DeliveryDate,POTransportationCode,"
                        strSQL = strSQL + " DeliveryUnit,KanbanCycle,ReceivingAreaCode,ReceivingContainerBoxCode,LocationNo,PalletizeCode,SupplierInvoiceNo,"
                        strSQL = strSQL + " AdjustedDeliveryDate,AdjustedDeliveryQty,SupplierInvoiceIssueDate,ReceivingCaseNo,DealerCode,DistColumn,DeliveredPartNo,"
                        strSQL = strSQL + " Ent_Dt,Ent_UserID,IP_Address)"
                        strSQL = strSQL + " values('" & gstrUNITID & "','" & intDocNo & "','" & txtcustomercode.Text & "','" & objexl.Cells(row, 1).Value & "','" & objexl.Cells(row, 2).Value & "',"
                        strSQL = strSQL + " '" + objexl.Cells(row, 3).Value.ToString() + "','" + objexl.Cells(row, 4).Value.ToString() + "','" + objexl.Cells(row, 5).Value.ToString() + "',"
                        strSQL = strSQL + " '" + objexl.Cells(row, 6).Value.ToString() + "','" + objexl.Cells(row, 7).Value.ToString() + "','" + objexl.Cells(row, 8).Value.ToString() + "',"
                        strSQL = strSQL + " '" + objexl.Cells(row, 9).Value.ToString() + "','" + objexl.Cells(row, 10).Value.ToString() + "','" + objexl.Cells(row, 11).Value.ToString() + "',"
                        strSQL = strSQL + " '" + objexl.Cells(row, 12).Value.ToString() + "','" + objexl.Cells(row, 13).Value + "','" + objexl.Cells(row, 14).Value + "',"
                        strSQL = strSQL + " '" + objexl.Cells(row, 15).Value.ToString() + "','" + objexl.Cells(row, 16).Value.ToString() + "','" + objexl.Cells(row, 17).Value.ToString() + "',"
                        strSQL = strSQL + " '" + objexl.Cells(row, 18).Value.ToString() + "','" + objexl.Cells(row, 19).Value.ToString() + "','" + objexl.Cells(row, 20).Value.ToString() + "',"
                        strSQL = strSQL + " '" + objexl.Cells(row, 21).Value.ToString() + "','" + objexl.Cells(row, 22).Value.ToString() + "','" + objexl.Cells(row, 23).Value.ToString() + "',"
                        strSQL = strSQL + " '" + objexl.Cells(row, 24).Value.ToString() + "','" + objexl.Cells(row, 25).Value + "','" + objexl.Cells(row, 26).Value.ToString() + "',"
                        strSQL = strSQL + " '" + objexl.Cells(row, 27).Value.ToString() + "','" + objexl.Cells(row, 28).Value.ToString() + "','" + objexl.Cells(row, 29).Value.ToString() + "',"
                        strSQL = strSQL + " getdate(),'" + mP_User + "','" + gstrIpaddressWinSck + "')"
                        SqlConnectionclass.ExecuteNonQuery(strSQL)
                        row = row + 1
                        strPONO = objexl.Cells(row, 5).Value

                    End While

                    If validate_itemcode_ToyotaFile(strPONO) <> "" Then
                        MsgBox("Following Items Are Not Defined Or Inactive " & vbCrLf & " In Item Master Or Customer Item Master " & vbCrLf & validate_itemcode_ToyotaFile(strPONO), MsgBoxStyle.Information, ResolveResString(100))
                        Exit Sub
                    End If
                    If validate_holiday() <> "" Then
                        MsgBox("Following Is\Are Not Working Day(s) " & vbCrLf & validate_holiday() & vbCrLf, MsgBoxStyle.Information, ResolveResString(100))
                        Exit Sub
                    End If

                    Save_Data_Toyota(intDocNo, strPONO)

                    MessageBox.Show("Schedule Saved.", ResolveResString(100), MessageBoxButtons.OK)
                    ''    lblMsg.Text = " S c h e d u l e    S a v e d "
                    mflag = 1
                Else
                    MsgBox("Excel Format Is Not Correct", MsgBoxStyle.Information, ResolveResString(100))
                End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            If Not objexl Is Nothing Then
                KillExcelProcess(objexl)
                objexl = Nothing
            End If
        End Try

    End Sub

    Private Function validate_itemcode() As String
        Try
            Dim strSQL As String
            Dim oDr As SqlDataReader
            Dim oCmd As New SqlCommand
            Dim stritemlist As String
            strSQL = "select distinct product_no from SCHEDULE_MST_TEMP where customer_code='" & Trim(txtcustomercode.Text) & "' and ipaddress='" & gstrIpaddressWinSck & "' and Unit_code = '" & gstrUNITID & "'" & "and product_no not in(select cim.item_code from custitem_mst cim inner join item_mst im on cim.item_code=im.item_code and cim.Unit_Code=im.unit_code where cim.account_code='" & Trim(txtcustomercode.Text) & "' and active=1 and im.status='A' and cim.Unit_code = '" & gstrUNITID & "')"
            With oCmd
                .Connection = SqlConnectionclass.GetConnection()
                .CommandType = CommandType.Text
                .CommandText = strSQL
                oDr = .ExecuteReader
            End With
            If oDr.HasRows Then
                While oDr.Read
                    If stritemlist = "" Then
                        stritemlist = oDr("product_no")
                    Else
                        stritemlist = stritemlist & vbCrLf & oDr("product_no")
                    End If
                End While
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Function validate_itemcode_ToyotaFile(ByVal strPONO As String) As String
        Try
            Dim strSQL As String
            Dim oDr As SqlDataReader
            Dim oCmd As New SqlCommand
            Dim stritemlist As String

            strSQL = "select distinct t.POPartNo from ToyotaSchedule_ReleaseNoWise t where t.customercode= '" + txtcustomercode.Text + "'" & _
                " and t.ip_address = '" + gstrIpaddressWinSck + "' and t.Unitcode = '" + gstrUNITID + "' AND T.PONO = '" + strPONO + "'" & _
                " and not exists (select cim.item_code from custitem_mst cim inner join item_mst im on cim.item_code=im.item_code" & _
                " and cim.Unit_Code=im.unit_code where cim.account_code = t.customercode and cim.cust_drgno = t.POPartNo and cim.active=1 and im.status='A')"

            With oCmd
                .Connection = SqlConnectionclass.GetConnection()
                .CommandType = CommandType.Text
                .CommandText = strSQL
                oDr = .ExecuteReader
            End With

            If oDr.HasRows Then
                While oDr.Read
                    If stritemlist = "" Then
                        stritemlist = oDr("POPartNo")
                    Else
                        stritemlist = stritemlist & vbCrLf & oDr("POPartNo")
                    End If
                End While
            End If

            Return stritemlist

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Function validate_holiday() As String
        Try
            Dim strSQL As String
            Dim oDr As SqlDataReader
            Dim oCmd As New SqlCommand
            Dim strholidaylist As String

            If optGeneric.Checked = True Then
                strSQL = "select distinct due_date from SCHEDULE_MST_TEMP t where customer_code='" & Trim(txtcustomercode.Text) & "' and ipaddress='" & gstrIpaddressWinSck & "' and Unit_code = '" & gstrUNITID & "' and not exists(select dt from Calendar_MST where work_flg = 0 and Unit_code = t.unit_code and dt = t.due_date)"
            Else
                strSQL = "select distinct DeliveryDate as due_date from ToyotaSchedule_ReleaseNoWise t where customercode='" & Trim(txtcustomercode.Text) & "' and ip_address='" & gstrIpaddressWinSck & "' and UnitCode = '" & gstrUNITID & "' and not exists(select dt from Calendar_MST where work_flg = 0 and Unit_code = t.unitcode and dt = t.DeliveryDate)"
            End If

            With oCmd
                .Connection = SqlConnectionclass.GetConnection()
                .CommandType = CommandType.Text
                .CommandText = strSQL
                oDr = .ExecuteReader
            End With

            If oDr.HasRows Then
                While oDr.Read
                    If strholidaylist = "" Then
                        strholidaylist = oDr("due_Date")
                    Else
                        strholidaylist = strholidaylist & vbCrLf & oDr("due_Date")
                    End If
                End While
            End If

            Return strholidaylist

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Sub setitemgrdproperty()
        Try
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

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub setitemgrdcoltypes(ByRef pintRowNo As Long)
        Try
            Dim i As Long
            With Me.fspitem
                .Row = pintRowNo
                .Col = enmforecastitem.status : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                .Col = enmforecastitem.ItemCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                .Col = enmforecastitem.Item_Description : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                .Col = enmforecastitem.Cust_DrgwNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            End With
            fspitem.ReDraw = True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub addrowinitemgrd(ByRef pintRows As Long)
        Try
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

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Function getfilename(ByVal strcustname As String) As String
        Try
            Dim oDr As SqlDataReader
            Dim oCmd As New SqlCommand
            Dim strFileName As String

            With oCmd
                .Connection = SqlConnectionclass.GetConnection()
                .CommandType = CommandType.Text
                .CommandText = "select convert(varchar(11),getdate(),106) + '-' + convert(varchar(8),getdate(),108) as date1"
                oDr = .ExecuteReader
            End With
            If oDr.HasRows Then
                While oDr.Read
                    strFileName = oDr("date1")
                End While
            End If

            getfilename = "Firm Schedule_" & strcustname & "_" & strFileName & ".xls"

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Sub fillitems(ByRef objexl As Excel.Application)
        Try
            Dim intcounter As Long
            Dim varstatus As Object
            Dim varItemCode As Object
            Dim varCustdrgno As Object
            Dim itemcount As Long
            Dim intcounter1 As Long
            Dim interval As Long
            Dim dtDate As Date
            Dim datecounter As Long
            Dim oDr As SqlDataReader
            Dim oCmd As New SqlCommand

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
            With oCmd
                .Connection = SqlConnectionclass.GetConnection()
                .CommandType = CommandType.Text
                .CommandText = "select dt from Calendar_mst where dt between '" & getDateForDB(dtfromdate.Value) & "' and '" & getDateForDB(dttodate.Value) & "' and work_flg=0 and Unit_code = '" & gstrUNITID & "' order by dt"
                oDr = .ExecuteReader
            End With
            If oDr.HasRows Then
                While oDr.Read
                    datecounter = datecounter + 1
                    objexl.Cells._Default(6 + datecounter, 1) = VB6.Format(oDr("dt"), "DD MMM YYYY")
                    objexl.Cells._Default(6 + datecounter, 1).Font.Bold = True
                    objexl.Cells._Default(6 + datecounter, 1).ColumnWidth = 15
                End While
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Function validate_itemselected() As Boolean
        Try
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

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Sub clear_fields()
        Try
            fspitem.MaxCols = 0
            fspitem.MaxRows = 0
            lblcustname.Text = ""
            txtcustomerhelp.Text = ""
            dtfromdate.Value = GetServerDate()
            dttodate.Value = GetServerDate()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Function validate_calender() As Boolean
        Try
            Dim oDr As SqlDataReader
            Dim oCmd As New SqlCommand
            Dim noofdays As Integer

            noofdays = DateDiff(Microsoft.VisualBasic.DateInterval.Day, dtfromdate.Value, dttodate.Value) + 1
            validate_calender = True

            With oCmd
                .Connection = SqlConnectionclass.GetConnection()
                .CommandType = CommandType.Text
                .CommandText = "select count(*) as noofdays from Calendar_mst where dt between '" & getDateForDB(dtfromdate.Value) & "' and '" & getDateForDB(dttodate.Value) & "' and Unit_code = '" & gstrUNITID & "'"
                oDr = .ExecuteReader
            End With

            If oDr.HasRows Then
                While oDr.Read
                    If noofdays <> oDr("noofdays") Then
                        validate_calender = False
                        Exit Function
                    End If
                End While
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Function validate_daterange() As Boolean
        Try
            Dim intinterval As Integer
            validate_daterange = True
            intinterval = DateDiff(Microsoft.VisualBasic.DateInterval.Month, dtfromdate.Value, dttodate.Value)
            If intinterval > 24 Then
                validate_daterange = False
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Function validate_Toyota_format(ByRef objexl As Excel.Application) As Boolean
        Try
            Dim row As Long
            Dim strcustomercode As String
            Dim strPONO As String
            Dim strDate As String
            Dim lngquantity As Integer
            Dim col As Long
            Dim row1 As Long
            Dim col1 As Long
            Dim row2 As Long
            Dim col2 As Long

            validate_Toyota_format = True

            strPONO = objexl.Cells(1, 5).value.ToString()

            If objexl.Cells(1, 1).value.ToString() <> "Depot Code" Or objexl.Cells(1, 2).value.ToString() <> "Supplier Code" Or objexl.Cells(1, 3).value.ToString() <> "Supplier Name" Or objexl.Cells(1, 4).value.ToString() <> "Factory Code" Or objexl.Cells(1, 5).value.ToString() <> "PO No" Or objexl.Cells(1, 6).value.ToString() <> "PO Item No" Or objexl.Cells(1, 7).value.ToString() <> "Maker Code" Or objexl.Cells(1, 8).value.ToString() <> "PO Part No" Or objexl.Cells(1, 9).value.ToString() <> "Part Name" Or objexl.Cells(1, 10).value.ToString() <> "Order Type" Or objexl.Cells(1, 11).value.ToString() <> "Supplier Part No" Or objexl.Cells(1, 12).value.ToString() <> "PO Qty" Or objexl.Cells(1, 13).value.ToString() <> "PO Date" Or objexl.Cells(1, 14).value.ToString() <> "Delivery Date" Or objexl.Cells(1, 15).value.ToString() <> "PO Transportation Code" Or objexl.Cells(1, 16).value.ToString() <> "Delivery Unit" Or objexl.Cells(1, 17).value.ToString() <> "Kanban Cycle" Or objexl.Cells(1, 18).value.ToString() <> "Receiving Area Code" Or objexl.Cells(1, 19).value.ToString() <> "Receiving Container Box Code" Or objexl.Cells(1, 20).value.ToString() <> "Location No" Or objexl.Cells(1, 21).value.ToString() <> "Palletize Code" Or objexl.Cells(1, 22).value.ToString() <> "Supplier Invoice No" Or objexl.Cells(1, 23).value.ToString() <> "Adjusted Delivery Date" Or objexl.Cells(1, 24).value.ToString() <> "Adjusted Delivery Qty" Or objexl.Cells(1, 25).value.ToString() <> "Supplier Invoice Issue Date" Or objexl.Cells(1, 26).value.ToString() <> "Receiving Case No" Or objexl.Cells(1, 27).value.ToString() <> "Dealer Code" Or objexl.Cells(1, 28).value.ToString() <> "Dist Column" Or objexl.Cells(1, 29).value.ToString() <> "Delivered Part No" Then
                MessageBox.Show("Incorrect File Format.", ResolveResString(100), MessageBoxButtons.OK)
                ''       lblMsg.Text = "I n c o r r e c t   F i l e   F o r m a t"
                validate_Toyota_format = False
                Exit Function
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Function validate_excel_format(ByRef objexl As Excel.Application) As Boolean
        Try
            Dim row As Long
            Dim strcustomercode As String
            Dim stritemcode As String
            Dim strDate As String
            Dim lngquantity As Integer
            Dim col As Long
            Dim row1 As Long
            Dim col1 As Long
            Dim row2 As Long
            Dim col2 As Long
            validate_excel_format = True
            row1 = 7 : row2 = 7
            row = 6 : col = 2
            stritemcode = Nothing
            stritemcode = objexl.Cells(row, col).Value
            If stritemcode = "" Then
                validate_excel_format = False
                Exit Function
            End If
            While (stritemcode <> "" And col <= objexl.Columns.Count)
                row = 6 : col = col
                stritemcode = objexl.Cells(row, col).Value
                row1 = 7 : col1 = 1
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
                    End While
                End If
                col = col + 1
            End While

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Sub dtfromdate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtfromdate.ValueChanged
        Try

            If dtfromdate.Value < GetServerDate() Then
                dtfromdate.Value = GetServerDate()
            ElseIf dtfromdate.Value > dttodate.Value Then
                dtfromdate.Value = GetServerDate()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub dttodate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dttodate.ValueChanged
        Try
            If dttodate.Value < dtfromdate.Value Then
                dttodate.Value = dtfromdate.Value
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub cmdsave_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdsave.Click
        Try
            Dim strcustomercode As String
            Dim stritemlist As String
            Dim intvalue As Long
            Dim objfso As New Scripting.FileSystemObject

            If chkPicklist.Checked = True And txtfilelocation.Text.Trim.Length = 0 Then
                If txtDocNo.Text.Length = 0 Then
                    MessageBox.Show("Upload schedue file or enter existing doc no to generate picklist.", ResolveResString(100), MessageBoxButtons.OK)
                    Exit Sub
                End If
                If txtDocNo.Text.Length > 0 Then
                    createPicklist(txtDocNo.Text)
                    Exit Sub
                End If
            End If

            If optToyotaReleaseFile.Checked = True Then
                Upload_Toyota_Forecast_Data()
                Exit Sub
            End If

            If fspforecastdetail.MaxRows > 0 Then
                If txtcustomercode.Text <> "" Then
                    save_data()
                    If objfso.FolderExists(gstrLocalCDrive & "Firm Schedule\uploaded_files") = True Then
                        If objfso.FileExists(gstrLocalCDrive & "Firm Schedule\uploaded_files\" + strUploadedFileName) Then
                            objfso.DeleteFile(gstrLocalCDrive & "Firm Schedule\uploaded_files\" + strUploadedFileName)
                        End If
                        objfso.MoveFile(txtfilelocation.Text, gstrLocalCDrive & "Firm Schedule\uploaded_files\")
                    Else
                        objfso.CreateFolder(gstrLocalCDrive & "Firm Schedule\uploaded_files")
                        objfso.MoveFile(txtfilelocation.Text, gstrLocalCDrive & "Firm Schedule\uploaded_files\")
                    End If
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
                MsgBox("Please Upload Schedule First", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
            objfso = Nothing
            objfso = Nothing

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            SqlConnectionclass.RollbackTran()
        End Try
    End Sub

    Private Sub chkPicklist_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPicklist.CheckedChanged
        Try

            txtDocNo.Text = ""
            If chkPicklist.Checked = False Then
                txtDocNo.Enabled = False
            Else
                txtDocNo.Enabled = True
                txtDocNo.Focus()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub optGeneric_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optGeneric.CheckedChanged

        Try

            cmdupload.Visible = True
            fspforecastdetail.Visible = True
            txtcustomercode.Text = "" : txtDocNo.Text = "" : txtfilelocation.Text = "" : chkPicklist.Checked = True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub optToyotaReleaseFile_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optToyotaReleaseFile.CheckedChanged

        Try

            cmdupload.Visible = False
            fspforecastdetail.Visible = False
            txtcustomercode.Text = "" : txtDocNo.Text = "" : txtfilelocation.Text = "" : chkPicklist.Checked = True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

End Class