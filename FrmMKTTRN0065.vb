Option Strict Off
Option Explicit On
Friend Class FrmMKTTRN0065
	Inherits System.Windows.Forms.Form
	Dim mintIndex As Short
	
	' Developed By    Amal.L
	' Date  21 March 2007
	' For   Mate - Bangalore
	' Purpose
	' For Finished Goods Barcode Printing, they have purchased Software From BarCode India Pvt Ltd,
	' Their software needs invoice information in CSV file format from Empower
	'-----------------------------------------------------------------------------------------------
	'Tables Used
	'SalesChallan_Dtl
	'Sales_Dtl
	'Invoice_Bar_Code_File_Path - This New Table is used to store the Folder Path..
	'Revised By     : Manoj Kr. Vaish
	'Revised On     : 30 Jan 2009
	'Issue ID       : eMpro-20090130-26775
	'Reason         : Addition of New Field in Invoice Bar Code CSV File.
    '--------------------------------------------------------------------------------------
    'Revised By     : Manoj Kr. Vaish
    'Revised On     : 30 Mar 2009
    'Issue ID       : eMpro-20090330-29420
    'Reason         : Change in Date Format of the New Field in Invoice Bar Code CSV File.
    '--------------------------------------------------------------------------------------
    'Revised By     : Manoj Kr. Vaish
    'Revised On     : 24 Apr 2009
    'Issue ID       : eMpro-20090424-30591
    'Reason         : Change in header fields in Invoice Bar Code CSV File.
    '--------------------------------------------------------------------------------------
    'Revised By     : Manoj Kr. Vaish
    'Revised On     : 24 Apr 2009
    'Issue ID       : eMpro-20090424-30591
    'Reason         : Change in header fields in Invoice Bar Code CSV File.
    '--------------------------------------------------------------------------------------
    'Revised By     : Manoj Kr. Vaish
    'Revised On     : 19 May 2009
    'Issue ID       : eMpro-20090519-31503
    'Reason         : Changes in Bar Code Text File Generation of TKML
    '                 as extra commas are coming in text format file.
    'Modified By Sanchi on 20 May 2011
    '   Modified to support MultiUnit functionality

    'MODIFIED BY DEEPAK ON 11 OCT 2011 FOR MULTIUNIT CHANGE MANAGEMENT
    '--------------------------------------------------------------------------------------
    'Revised By     : prashant Rajpal
    'Revised On     :  23 Nov 2011
    'Issue ID       : 10162900
    'Reason         : Changes For Citrix issue 
    '--------------------------------------------------------------------------------------
    'Revised By     : prashant Rajpal
    'Revised On     : 06 Dec 2011
    'Issue ID       : 10168209 
    'Reason         : Changed for -> TOYOTA BarCoding , path should not be editable 
    '--------------------------------------------------------------------------------------
    'Modified By Roshan Singh on 20 Dec 2011 for multiUnit change management    

    Private Function FN_Display_Folder_Path() As Object
        On Error GoTo ErrHandler
        Dim adors As New ADODB.Recordset
        Dim sql As String
        sql = " Select FolderPath from Invoice_Bar_Code_File_Path where UNIT_CODE='" & gstrUNITID & "'"
        adors.Open(sql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic)
        If adors.EOF = False Then
            txtfilepath.Text = IIf(IsDBNull(adors.Fields("FolderPath").Value), "", adors.Fields("FolderPath").Value)
        Else
            txtfilepath.Text = ""
        End If
        If adors.State Then adors.Close()
        adors = Nothing
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Private Function FN_Generate_File(ByRef File_Name As String) As Boolean
        On Error GoTo ErrHandler
        Dim Obj_FSO As Scripting.FileSystemObject
        Dim strBarCodedata As String
        Dim adors As ADODB.Recordset
        Dim Ex_Amount, sql, Check_Sheet_No As String
        'Declared by SAURAV KUMAR
        'Issue Id 10119745
        Dim intSales As Integer = 0
        '-------------------------------------------------------
        FN_Generate_File = True
        Obj_FSO = New Scripting.FileSystemObject
        If Obj_FSO.FileExists(File_Name) = True Then
            If MsgBox("File Already Exists! Do you want Proceed.. ? ", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                Obj_FSO.DeleteFile((File_Name))
            Else
                FN_Generate_File = False
                Exit Function
            End If
        End If
        FileOpen(1, File_Name, OpenMode.Append)
        Obj_FSO = Nothing
        'Total Excise Amount
        '-----------------------------------------
        sql = " Select sum(Excise_Tax) as Total_Excise_Amt"
        sql = sql & " From Sales_Dtl"
        sql = sql & " Where Doc_No ='" & Trim(txtInvNo.Text) & "' and UNIT_CODE='" & gstrUNITID & "'"
        adors = New ADODB.Recordset
        adors.Open(sql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic)
        If adors.EOF = False Then
            Ex_Amount = IIf(IsDBNull(adors.Fields("Total_Excise_Amt").Value), "", VB6.Format(adors.Fields("Total_Excise_Amt").Value, "#0.00"))
        End If
        If adors.State Then adors.Close()
        '-----------------------------------------
        'Header Information
        '[Invoice No, Total Exise Amount , VAT/CST Amount ,Total Payable Invoice Amount]
        'It Should Be In First Line
        'New Format against Issue ID eMpro-20090130-26775
        '[Invoice No,Invoice Date,Total Basic Amount,Total Exise Amount,
        'Total Ecess Amount,Total Secess Amount,VAT/CST Amount ,Total Payable Invoice Amount]
        '-----------------------------------------
        sql = "Select a.Doc_No as Inv_No,a.LorryNo_Date as CheckSheetNo,a.Total_Amount as Invoice_Amount,a.Sales_Tax_Amount,"
        sql = sql & " a.Invoice_date,isnull(sum(b.basic_Amount),0) as Basic_Amount,isnull(a.ECESS_Amount,0) as ECESS_Amount,"
        sql = sql & " isnull(a.SECESS_Amount,0) as SECESS_Amount,isnull(sum(b.cvd_amount),0)as cvd_amount From Saleschallan_dtl a"
        sql = sql & " Inner Join Sales_dtl b on a.doc_no=b.doc_no and a.UNIT_CODE=b.UNIT_CODE"
        sql = sql & " Where a.account_code='" & Trim(txtcustomer.Text) & "'"
        sql = sql & " And a.Doc_No='" & Trim(txtInvNo.Text) & "'"
        sql = sql & " And a.Invoice_Date = '" & getDateForDB(DTPicker1.Value) & "'"
        sql = sql & " And a.bill_flag=1"
        sql = sql & " And a.cancel_flag=0 AND a.UNIT_CODE='" & gstrUNITID & "'"
        sql = sql & " Group by a.Doc_No,a.Invoice_date,a.LorryNo_Date,a.Total_Amount,a.Sales_Tax_Amount,a.ECESS_Amount,a.SECESS_Amount"
        sql = sql & " Order by a.doc_no desc"
        adors.Open(sql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic)
        If adors.EOF = False Then
            strBarCodedata = strBarCodedata & IIf(IsDBNull(adors.Fields("inv_no").Value), "", adors.Fields("inv_no").Value) & "," & VB6.Format(adors.Fields("Invoice_Date").Value, "ddmmyy") & ","
            strBarCodedata = strBarCodedata & VB6.Format(adors.Fields("Basic_Amount").Value, "#0.00") & "," & VB6.Format(Ex_Amount, "#0.00") & "," & VB6.Format(adors.Fields("CVD_Amount").Value, "#0.00") & ","
            strBarCodedata = strBarCodedata & VB6.Format(adors.Fields("ECESS_Amount").Value, "#0.00") & "," & VB6.Format(adors.Fields("SECESS_Amount").Value, "#0.00") & "," & IIf(IsDBNull(adors.Fields("Sales_Tax_Amount").Value), "", VB6.Format(adors.Fields("Sales_Tax_Amount").Value, "#0.00")) & vbCrLf
            Check_Sheet_No = IIf(IsDBNull(adors.Fields("CheckSheetNo").Value), "", adors.Fields("CheckSheetNo").Value)
        End If
        If adors.State Then adors.Close()
        '-----------------------------------------
        'Detailed Information
        'Check Sheet No / Delivery Note Number , Part Number , Quantity
        'Check Sheet No - LorryNo_Date
        'Part Number    - Customer Item Code
        'Qty            - Invoice Qty
        'It Should Start Second Line Onwards
        '------------------------------------------
        sql = " Select Distinct Cust_Item_Code,Sales_Quantity"
        sql = sql & " From Sales_Dtl"
        sql = sql & " Where Doc_No ='" & Trim(txtInvNo.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'"
        adors.Open(sql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic)
        While adors.EOF = False
            'Change By Deepak on 11-Oct-2011 For Support change Management----------
            '********************************************************************************
            'Modified By        :       SAURAV KUMAR
            'Modified Date      :       28 July 2011
            'Issue Id           :       10119745
            '********************************************************************************
            intSales = VB6.Format(adors.Fields("Sales_Quantity").Value, "#0")
            strBarCodedata = strBarCodedata & Check_Sheet_No & "," & IIf(IsDBNull(adors.Fields("Cust_Item_Code").Value), "", Trim(adors.Fields("Cust_Item_Code").Value)) & ","
            strBarCodedata = strBarCodedata & IIf(IsDBNull(adors.Fields("Sales_Quantity").Value), 0, intSales) & vbCrLf
            adors.MoveNext()
            intSales = 0
            '-------------------------------------------
        End While
        adors = Nothing
        PrintLine(1, strBarCodedata)
        FileClose(1)
        Exit Function
ErrHandler:
        FN_Generate_File = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Private Function FN_Save_Folder_Path() As Object
        On Error GoTo ErrHandler
        mP_Connection.Execute(" Delete from Invoice_Bar_Code_File_Path WHERE UNIT_CODE='" & gstrUNITID & "'")
        mP_Connection.Execute(" Insert Into Invoice_Bar_Code_File_Path(FolderPath,UNIT_CODE) Values('" & Trim(txtfilepath.Text) & "','" & gstrUNITID & "')")
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Private Sub CmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdCancel.Click
        On Error GoTo ErrHanler
        txtcustomer.Text = ""
        txtfilepath.Text = ""
        txtInvNo.Text = ""
        cmdfilepath.Text = ">>"
        CmdGenerate.Enabled = False
        FraFilePath.Enabled = False
        DTPicker1.Value = GetServerDate()
        On Error Resume Next
        txtcustomer.Focus()
        Exit Sub
ErrHanler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdClose.Click
        On Error GoTo ErrHandler
        Me.Close()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmdCustomer_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdcustomer.Click
        On Error GoTo ErrHandler
        Dim strCustCode() As String
        Dim strCust As String
        Dim strString As String
        strString = txtcustomer.Text & "%"
        With ctlEMPHelp1
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        If txtcustomer.Text <> "" Then strCust = "where Account_Code like '" & strString & "' and UNIT_CODE='" & gstrUNITID & "'" Else  : strCust = " WHERE UNIT_CODE='" & gstrUNITID & "'"
        strCustCode = Me.ctlEMPHelp1.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Distinct Account_Code,Cust_name from SalesChallan_Dtl " & strCust & "", "Customer List")
        If UBound(strCustCode) = -1 Then Exit Sub
        txtcustomer.Tag = ""
        If strCustCode(0) = "0" Then
            ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO) : txtcustomer.Text = "" : txtcustomer.Focus() : Exit Sub
        Else
            Me.txtcustomer.Text = strCustCode(0)
            txtcustomer.Tag = txtcustomer.Text
        End If
        Call txtcustomer_Validating(txtcustomer, New System.ComponentModel.CancelEventArgs(False))
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdfilepath_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdfilepath.Click
        On Error GoTo ErrHandler
        If cmdfilepath.Text = ">>" Then
            FraFilePath.Enabled = True
            cmdfilepath.Text = "<<"
        ElseIf cmdfilepath.Text = "<<" Then
            FraFilePath.Enabled = False
            cmdfilepath.Text = ">>"
        End If
        txtfilepath.Text = Dir1.Path
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Function FN_Required_Field() As Boolean
        On Error GoTo ErrHandler
        If Trim(txtcustomer.Text) = "" Then
            MsgBox(" Please Select the Customer ", MsgBoxStyle.Information, ResolveResString(100))
            txtcustomer.Focus()
            Exit Function
        End If
        If Trim(txtInvNo.Text) = "" Then
            MsgBox(" Please Select Invoice Number ", MsgBoxStyle.Information, ResolveResString(100))
            txtInvNo.Focus()
            Exit Function
        End If
        If Trim(txtfilepath.Text) = "" Then
            MsgBox(" Please Select Folder Path ", MsgBoxStyle.Information, ResolveResString(100))
            txtfilepath.Focus()
            Exit Function
        End If
        FN_Required_Field = True
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Private Sub cmdGenerate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdGenerate.Click
        On Error GoTo ErrHandler
        Dim File_Name As String
        If FN_Required_Field() = False Then
            Exit Sub
        End If
        Call FN_Save_Folder_Path()
        File_Name = Trim(txtfilepath.Text) & "\BC" & Trim(txtInvNo.Text) & VB6.Format(GetServerDate(), "ddmmyy") & ".CSV"
        If FN_Generate_File(File_Name) = True Then
            MsgBox(" File Path is  : " & File_Name.Replace("\\", "\"), MsgBoxStyle.Information, " File Generated Succesfully ! ")
        End If
        Exit Sub
ErrHandler:
        If Err.Number = 1004 Then
            MsgBox("Please select the folder by double clicking")
            Exit Sub
        End If
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub CmdInvoiceNo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdInvoiceNo.Click
        On Error GoTo ErrHandler
        Dim StrInvNo() As String
        Dim StrInv As String
        Dim strString As String
        Dim Cond As String
        If Trim(txtcustomer.Text) = "" Then
            MsgBox(" Please Select Customer first, then Proceed..", MsgBoxStyle.Information, ResolveResString(100))
            txtcustomer.Focus()
            Exit Sub
        End If
        strString = txtInvNo.Text & "%"
        With ctlEMPHelp1
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        Cond = " Where Bill_Flag = 1 And Cancel_Flag =0 And Account_code ='" & Trim(txtcustomer.Text) & "' And Invoice_Date='" & getDateForDB(DTPicker1.Value) & "' AND UNIT_CODE='" & gstrUNITID & "'"
        If txtInvNo.Text <> "" Then StrInv = " And Doc_No like '" & strString & "' " Else StrInv = ""
        StrInv = Cond & StrInv
        StrInvNo = Me.ctlEMPHelp1.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Distinct Doc_No,Cust_Ref from SalesChallan_Dtl " & StrInv & " ", "Invoice List")
        If UBound(StrInvNo) = -1 Then Exit Sub
        txtInvNo.Tag = ""
        If StrInvNo(0) = "0" Then
            ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO) : txtInvNo.Text = "" : txtInvNo.Focus() : Exit Sub
        Else
            Me.txtInvNo.Text = StrInvNo(0)
            txtInvNo.Tag = txtInvNo.Text
        End If
        Call txtInvNo_Validating(txtInvNo, New System.ComponentModel.CancelEventArgs(False))
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub Dir1_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Dir1.Change
        On Error GoTo ErrHandler
        txtfilepath.Text = Dir1.Path
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub Drive1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Drive1.SelectedIndexChanged
        On Error GoTo ErrHandler
        Dir1.Path = Drive1.Drive
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub FrmMKTTRN0065_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub FrmMKTTRN0065_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub FrmMKTTRN0065_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlFormHeader1_Click(ctlFormHeader1, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub FrmMKTTRN0065_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        Call Initialize_controls()
        Call FN_Display_Folder_Path()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub Initialize_controls()
        On Error GoTo ErrHandler
        Call FillLabelFromResFile(Me)
        Call FitToClient(Me, fracontrols, ctlFormHeader1, FraButton)
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        txtfilepath.Text = ""
        DTPicker1.Format = DateTimePickerFormat.Custom
        DTPicker1.CustomFormat = gstrDateFormat
        FraFilePath.Enabled = False
        CmdGenerate.Enabled = False
        'Issue : 10168209
        FraFilePath.Visible = False
        cmdfilepath.Visible = False
        'Issue : 10168209 
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtcustomer_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtcustomer.TextChanged
        On Error GoTo ERR_Renamed
        txtcustomerName.Text = ""
        txtInvNo.Text = ""
        TxtCustRef.Text = ""
        'txtfilepath.Text = ""
        DTPicker1.Value = GetServerDate()
        Exit Sub
ERR_Renamed:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtcustomer_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtcustomer.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                If Len(txtcustomer.Text) > 0 Then
                    Call txtcustomer_Validating(txtcustomer, New System.ComponentModel.CancelEventArgs(False))
                End If
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCustomer_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtcustomer.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            If cmdcustomer.Enabled Then Call cmdCustomer_Click(cmdcustomer, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtcustomer_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtcustomer.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ERR_Renamed
        Dim sql As String
        Dim adors As New ADODB.Recordset
        If Trim(txtcustomer.Text) = "" Then GoTo EventExitSub
        If Len(txtcustomer.Text.Trim) > 0 Then
            mP_Connection.Execute("set dateformat  'dmy'")
            sql = " Select Cust_Name From SalesChallan_Dtl Where Account_Code='" & Trim(txtcustomer.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'"
            adors.Open(sql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic)
            If adors.EOF = False Then
                txtcustomerName.Text = IIf(IsDBNull(adors.Fields("Cust_Name").Value), "", adors.Fields("Cust_Name").Value)
            Else
                MsgBox("Invalid Customer Code", MsgBoxStyle.Information, ResolveResString(100))
                txtcustomer.Text = ""
                txtcustomerName.Text = ""
                DTPicker1.Value = GetServerDate()
                txtInvNo.Text = ""
                TxtCustRef.Text = ""
            End If
        End If
        If adors.State Then adors.Close()
        adors = Nothing
        GoTo EventExitSub
ERR_Renamed:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtInvNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtInvNo.TextChanged
        On Error GoTo ErrHandler
        TxtCustRef.Text = ""
        'txtfilepath.Text = ""
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtInvNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtInvNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Dim aa As String = Chr(KeyAscii)
        Dim ff As String = Convert.ToInt16(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Dim IntVal As Integer = 0
        If Integer.TryParse(Chr(KeyAscii).ToString, IntVal) = True Then
            If (Convert.ToInt16(Chr(KeyAscii).ToString) < 0 Or Convert.ToInt16(Chr(KeyAscii).ToString) > 9) And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
        Else
            If KeyAscii <> 8 Then KeyAscii = 0
        End If
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                If Len(txtInvNo.Text) > 0 Then
                    Call txtInvNo_Validating(txtInvNo, New System.ComponentModel.CancelEventArgs(False))
                End If
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtInvNo_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtInvNo.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            If CmdInvoiceNo.Enabled Then Call CmdInvoiceNo_Click(CmdInvoiceNo, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtInvNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtInvNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ERR_Renamed
        Dim sql As String
        Dim adors As New ADODB.Recordset
        If Trim(txtcustomer.Text) = "" Then
            MsgBox("Please Select Customer,then you can Proceed...", MsgBoxStyle.Information, ResolveResString(100))
            txtcustomer.Focus()
            GoTo EventExitSub
        End If
        If Trim(txtInvNo.Text) = "" Then GoTo EventExitSub
        If Len(txtInvNo.Text) > 0 Then
            mP_Connection.Execute("set dateformat  'dmy'")
            sql = " Select cust_ref From SalesChallan_Dtl Where Account_Code='" & Trim(txtcustomer.Text) & "'"
            sql = sql & " And Doc_No='" & Trim(txtInvNo.Text) & "'"
            sql = sql & " And Invoice_Date='" & getDateForDB(DTPicker1.Value) & "'"
            sql = sql & " And Bill_Flag = 1 And Cancel_Flag =0 AND UNIT_CODE='" & gstrUNITID & "'"
            adors.Open(sql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic)
            If adors.EOF = False Then
                TxtCustRef.Text = IIf(IsDBNull(adors.Fields("Cust_Ref").Value), "", adors.Fields("Cust_Ref").Value)
                CmdGenerate.Enabled = True
            Else
                MsgBox("Invalid Invoice Number.", MsgBoxStyle.Information, ResolveResString(100))
                txtInvNo.Text = ""
                CmdGenerate.Enabled = False
                TxtCustRef.Text = ""
                'txtfilepath.Text = ""
                txtInvNo.Focus()
            End If
            If adors.State Then adors.Close()
            adors = Nothing
            GoTo EventExitSub
        End If
ERR_Renamed:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTPicker1_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DTPicker1.ValueChanged
        On Error GoTo errorhandler
        txtInvNo.Text = ""
        TxtCustRef.Text = ""
        'txtfilepath.Text = ""
        Exit Sub
errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub CmdGenerate_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdGenerate.Click
    End Sub
End Class