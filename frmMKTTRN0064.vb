Option Strict Off
Option Explicit On
Imports System.Data.SqlClient
Imports System.IO
Friend Class frmMKTTRN0064
    Inherits System.Windows.Forms.Form
    '---------------------------------------------------------------------------
    'Copyright          :   MIND Ltd.
    'Form Name          :   frmMKTTRN0064
    'Created By         :   Manoj Kr Vaish
    'Created on         :   03 Feb 2009 Issue ID eMpro-20090204-27027
    'Description        :   ASN File Generation for Mahindra & Mahindra
    '---------------------------------------------------------------------------
    'Revised By         :   Manoj Vaish
    'Revision On        :   22 Jun 2009
    'Issue ID           :   eMpro-20090610-32326
    'History            :   ASN File Generation for Nissan
    '----------------------------------------------------
    'Modified by    :   Virendra Gupta
    'Modified ON    :   26/05/2011
    'Modified to support MultiUnit functionality
    '-----------------------------------------------------------------------
    'Modified By        :   Prashant Rajpal
    'Revision On        :   27 May 2012-04 june 2012
    'Issue ID           :   10229992 
    'History            :   ASN File Generation for HUUNDAI 
    '----------------------------------------------------
    'issue id : 10240780 changed by prashant rajpal for formattting issue 
    '***********************************************************************************
    'Modified By        :   Prashant Rajpal
    'Revision On        :   09 SEP 2013
    'Issue ID           :   10229989
    'History            :   ASN CHANGES FORM MULTIPLE SALES ORDER FUNCTIONLITY
    '----------------------------------------------------

    Dim mlngCounter As Integer
    Dim msqlcon As SqlConnection
    Dim msqlcmd As SqlCommand
    Dim msqldr As SqlDataReader
    Dim mintIndex As Short
    Private Enum InvoiceGrid
        invSel = 1
        InvNo = 2
        invDate = 3
        invTypeDesc = 4
        invSubTypeDesc = 5
    End Enum
    Private Sub cmdCustHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCustHelp.Click
        '-------------------------------------------------------------------------------------
        'Revised By     : Manoj Kr. Vaish
        'Revised On     : 04 Feb 2009
        'Arguments      : NIL
        'Return Value   : NIL
        'Issue ID       : eMpro-20090204-27027
        'Reason         : To Show the customer code help
        '--------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strHelp() As String
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        If gstrUNITID = "STH" Then
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select CUSTOMER_CODE,CUST_NAME  FROM  VW_ASN_CUSTOMERHELP_Gen_Motors where unit_code='" & gstrUNITID & "'", "Help", 2)
        Else
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select CUSTOMER_CODE,CUST_NAME  FROM  VW_ASN_CUSTOMERHELP where unit_code='" & gstrUNITID & "'", "Help", 2)
        End If


        If UBound(strHelp) <> -1 Then
            If strHelp(0) <> "0" Then
                Me.txtCustomerCode.Text = strHelp(0)
                Me.LblCustomerName.Text = strHelp(1)
                Me.txtCustomerCode.Enabled = False
                spgrid.MaxRows = 0
            Else
                Me.txtCustomerCode.Text = ""
                Me.LblCustomerName.Text = ""
                Me.txtCustomerCode.Enabled = False
                spgrid.MaxRows = 0
                MsgBox(" No record available", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdShowInvoices_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdShowInvoices.Click
        On Error GoTo Errorhandler
        If Len(Trim(txtCustomerCode.Text)) = 0 Then
            MsgBox(" Select Customer Code.", MsgBoxStyle.Information, ResolveResString(100))
            cmdCustHelp.Focus()
            Exit Sub
        ElseIf dtFromDate.Value > dtToDate.Value Then
            MsgBox("[From date] should be less than or equal to [To date].", MsgBoxStyle.Information, ResolveResString(100))
            dtFromDate.Focus()
            Exit Sub
        Else
            Call SetGridCells()
            Call ShowPendingInvoices()
        End If
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ctlFormHeader1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        On Error GoTo ErrHandler
        Call ShowHelp("HLPMKTTRN0064.HTM")
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0064_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0064_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0064_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        'Add Form Name To Window List
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        'Fill Lebels From Resource File
        'Call FillLabelFromResFile(Me) 'Fill Labels >From Resource File
        Me.ctlFormHeader1.HeaderString = Mid(Me.ctlFormHeader1.HeaderString, InStr(1, Me.ctlFormHeader1.HeaderString(), "-") + 1, Len(Me.ctlFormHeader1.HeaderString()))
        Call FitToClient(Me, frmMain, ctlFormHeader1, cmdLockInvoice)
        SetGridCells()
        Me.dtFromDate.Format = DateTimePickerFormat.Custom
        Me.dtFromDate.CustomFormat = gstrDateFormat
        Me.dtFromDate.Value = GetServerDate()
        Me.dtToDate.Format = DateTimePickerFormat.Custom
        Me.dtToDate.CustomFormat = gstrDateFormat
        Me.dtToDate.Value = GetServerDate()
        optCheckAll.Checked = True
        msqlcon = SqlConnectionclass.GetConnection(gstrConnectSQLClient)
        Call DisableControls()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SetGridCells()
        On Error GoTo Errorhandler
        With Me.spgrid
            .MaxRows = 0 : .MaxCols = 5
            .Row = 0 : .set_RowHeight(0, 300)
            .Col = InvoiceGrid.invSel : .Text = "Select Invoice" : .set_ColWidth(InvoiceGrid.invSel, 1500)
            .Col = InvoiceGrid.InvNo : .Text = "Invoice No" : .set_ColWidth(InvoiceGrid.InvNo, 1600)
            .Col = InvoiceGrid.invDate : .Text = "Invoice Date" : .set_ColWidth(InvoiceGrid.invDate, 1600)
            .Col = InvoiceGrid.invTypeDesc : .Text = "Invoice Type" : .set_ColWidth(InvoiceGrid.invTypeDesc, 1800)
            .Col = InvoiceGrid.invSubTypeDesc : .Text = "Invoice Category" : .set_ColWidth(InvoiceGrid.invSubTypeDesc, 1800)
        End With
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub ShowPendingInvoices()
        'Revised By     : Manoj Kr. Vaish
        'Revised On     : 04 Feb 2009
        'Arguments      : NIL
        'Return Value   : NIL
        'Issue ID       : eMpro-20090204-27027
        'Reason         : Show Pending Invoices for ASN Generation
        '--------------------------------------------------------------------------------------
        On Error GoTo Errorhandler
        Dim rsobject As New ADODB.Recordset
        Dim cmdObject As New ADODB.Command
        Dim strsql As String
        rsobject.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
        With cmdObject
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandText = "GETINVOICE_ASN"
            .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
            .Parameters.Append(.CreateParameter("@CUSTOMER_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(txtCustomerCode.Text)))
            .Parameters.Append(.CreateParameter("@FROM_DATE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 11, getDateForDB(dtFromDate.Value)))
            .Parameters.Append(.CreateParameter("@TO_DATE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 11, getDateForDB(dtToDate.Value)))
            .Parameters.Append(.CreateParameter("@ERR", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 100))
            .let_ActiveConnection(mP_Connection)
            rsobject = .Execute
        End With
        If Len(cmdObject.Parameters(4).Value) > 0 Then
            MsgBox(cmdObject.Parameters(4).Value, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            cmdObject = Nothing
            rsobject = Nothing
            Exit Sub
        End If
        cmdObject = Nothing
        If Not rsobject.EOF Then
            mlngCounter = 1
            With spgrid
                Do While Not rsobject.EOF
                    AddNewRow()
                    .Row = mlngCounter
                    .Col = InvoiceGrid.InvNo : .Text = rsobject.Fields("doc_No").Value
                    .Col = InvoiceGrid.invDate : .Text = VB6.Format(rsobject.Fields("Invoice_Date").Value, gstrDateFormat)
                    .Col = InvoiceGrid.invTypeDesc : .Text = rsobject.Fields("description").Value
                    .Col = InvoiceGrid.invSubTypeDesc : .Text = rsobject.Fields("Sub_Type_Description").Value
                    rsobject.MoveNext() : mlngCounter = mlngCounter + 1
                Loop
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                Me.optCheckAll.Checked = True
            End With
        Else
            Call MsgBox("No data found between selected dates", MsgBoxStyle.OkOnly, ResolveResString(100))
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        End If
        rsobject = Nothing
        Exit Sub
Errorhandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub AddNewRow()
        On Error GoTo Errorhandler
        With spgrid
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows : .set_RowHeight(.Row, 300)
            .Col = InvoiceGrid.invSel : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .Value = CStr(1) : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeCheckCenter = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
            .Col = InvoiceGrid.InvNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = InvoiceGrid.invDate : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = InvoiceGrid.invTypeDesc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = InvoiceGrid.invSubTypeDesc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Row = .MaxRows : .Row2 = .MaxRows : .Col = InvoiceGrid.InvNo
        End With
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0064_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub optCheckAll_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optCheckAll.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            With Me.spgrid
                If .MaxRows > 0 Then
                    For mlngCounter = 1 To .MaxRows
                        .Row = mlngCounter : .Col = InvoiceGrid.invSel : .Value = CStr(1)
                    Next
                End If
            End With
            Exit Sub
ErrHandler:  'The Error Handling Code Starts here
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        End If
    End Sub
    Private Sub optUncheckAll_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optunCheckAll.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            With Me.spgrid
                If .MaxRows > 0 Then
                    For mlngCounter = 1 To .MaxRows
                        .Row = mlngCounter : .Col = InvoiceGrid.invSel : .Value = CStr(0)
                    Next
                End If
            End With
            Exit Sub
ErrHandler:  'The Error Handling Code Starts here
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        End If
    End Sub
    Private Function GenerateASNFileForMahindra(ByVal pstrFileLocation As String, ByVal pstrInvoiceNo As String) As Boolean
        'Revised By     : Manoj Kr. Vaish
        'Revised On     : 04 Feb 2009
        'Arguments      : File Location,Invoice No
        'Return Value   : True/False
        'Issue ID       : eMpro-20090204-27027
        'Reason         : To Generate ASN File for selected Invoice
        '--------------------------------------------------------------------------------------
        Dim strsql As String
        Dim rsGetASNData As ClsResultSetDB
        Dim Obj_FSO As Scripting.FileSystemObject
        Dim strLocation As String
        Dim strFileName As String
        Dim strRecord As String
        Dim intLineNo As Short
        On Error GoTo Err_Handler
        GenerateASNFileForMahindra = True
        '----------------------------------------
        Obj_FSO = New Scripting.FileSystemObject
        If Not Obj_FSO.FolderExists(pstrFileLocation) Then
            Obj_FSO.CreateFolder(pstrFileLocation)
        End If
        If Mid(Trim(pstrFileLocation), Len(Trim(pstrFileLocation))) <> "\" Then
            strLocation = pstrFileLocation & "\"
        End If
        strFileName = "ASN" & VB6.Format(GetServerDate(), "ddMMyy") & ".csv"
        strFileName = strLocation & strFileName
        Kill(strLocation & "*.csv")
        FileClose(1)
        FileOpen(1, strFileName, OpenMode.Append)
        Obj_FSO = Nothing
        rsGetASNData = New ClsResultSetDB
        strsql = "select a.Invoice_Date,CASE WHEN LEN(B.EXTERNAL_SALESORDER_NO)>0 THEN B.EXTERNAL_SALESORDER_NO ELSE   A.CUST_REF END AS CUST_REF,A.LorryNo_Date,A.Doc_no,A.Vehicle_No,A.Carriage_Name,"
        strsql = strsql & " B.Sales_Quantity,A.Total_Amount,B.Excise_Tax,'10' as Item_Srno from Saleschallan_dtl a"
        strsql = strsql & " Inner join Sales_Dtl b on a.doc_no=b.doc_no AND a.UNIT_CODE=b.UNIT_CODE where a.bill_flag=1 and a.cancel_flag=0 AND a.UNIT_CODE = '" & gstrUNITID & "'"
        strsql = strsql & " and a.doc_no in(" & pstrInvoiceNo & ")"
        rsGetASNData.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsGetASNData.GetNoRows > 0 Then
            rsGetASNData.MoveFirst()
            Do While Not rsGetASNData.EOFRecord
                strRecord = ""
                strRecord = IIf(IsDBNull(rsGetASNData.GetValue("Cust_Ref")), "", rsGetASNData.GetValue("Cust_Ref"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("Item_Srno")), "", rsGetASNData.GetValue("Item_Srno"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("Sales_Quantity")), "", rsGetASNData.GetValue("Sales_Quantity"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("Doc_no")), "", rsGetASNData.GetValue("Doc_no"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("Invoice_Date")), "", VB6.Format(rsGetASNData.GetValue("Invoice_Date"), "dd.mm.yyyy"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("Total_Amount")), "", VB6.Format(rsGetASNData.GetValue("Total_Amount"), "#0.00"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("Excise_Tax")), "", VB6.Format(rsGetASNData.GetValue("Excise_Tax"), "#0.00"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("LorryNo_Date")), "", rsGetASNData.GetValue("LorryNo_Date"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("Invoice_Date")), "", VB6.Format(rsGetASNData.GetValue("Invoice_Date"), "dd.mm.yyyy"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("Vehicle_No")), "", rsGetASNData.GetValue("Vehicle_No"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("Carriage_Name")), "", rsGetASNData.GetValue("Carriage_Name"))
                PrintLine(1, strRecord) : intLineNo = intLineNo + 1
                rsGetASNData.MoveNext()
            Loop
            rsGetASNData.ResultSetClose()
            rsGetASNData = Nothing
        Else
            MsgBox("No Invoice Records found to generate the File.", MsgBoxStyle.Information, ResolveResString(100))
            FileClose(1)
            Kill(strFileName)
            rsGetASNData.ResultSetClose()
            rsGetASNData = Nothing
            GenerateASNFileForMahindra = False
            Exit Function
        End If
        FileClose(1)
        Exit Function
Err_Handler:
        If Err.Number = 55 Then
            MsgBox("File Already Open, Cann't Generate the ASN File.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            GenerateASNFileForMahindra = False
            Exit Function
        End If
        GenerateASNFileForMahindra = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    'Created By Tek on 04_03_25 for General_Motors

    Private Function GenerateASNFileForGen_Motors(ByVal pstrFileLocation As String, ByVal pstrInvoiceNo As String) As Boolean
        'Created By     : Tek Chand
        'Created On     : 04 Mar 2025
        'Reason         : To Generate ASN File for selected Invoices
        '--------------------------------------------------------------------------------------
        Dim strsql As String
        Dim CUMMULATIVE_QTY As Integer
        Dim Obj_FSO As Scripting.FileSystemObject
        Dim strLocation As String
        Dim strRecord As String
        Dim strINvoiceNo As String
        Dim fs As FileStream
        Dim sw As StreamWriter
        Dim strASNFilepath As String

        GenerateASNFileForGen_Motors = True
        '----------------------------------------
        Obj_FSO = New Scripting.FileSystemObject
        If Not Obj_FSO.FolderExists(pstrFileLocation) Then
            Obj_FSO.CreateFolder(pstrFileLocation)
        End If
        If Mid(Trim(pstrFileLocation), Len(Trim(pstrFileLocation))) <> "\" Then
            strLocation = pstrFileLocation & "\"
        End If
        'FileClose(1)
        'FileOpen(1, strFileName, OpenMode.Append)
        'Obj_FSO = Nothing
        Try
            GenerateASNFileForGen_Motors = True
            With spgrid
                For mlngCounter = 1 To .MaxRows
                    .Row = mlngCounter : .Col = InvoiceGrid.invSel
                    If CDbl(.Value) = 1 Then
                        .Col = InvoiceGrid.InvNo
                        strINvoiceNo = Trim(.Text)
                        strsql = "Select * From dbo.FN_GETASNDETAIL_GENMOTORS('" & strINvoiceNo & "','" & gstrUNITID & "')"
                        msqlcmd = New SqlCommand(strsql, msqlcon)
                        msqldr = msqlcmd.ExecuteReader()
                        strRecord = ""
                        While msqldr.Read()
                            strRecord = strRecord & IIf(IsDBNull(msqldr("SUPLIEREDICODE")), "", msqldr("SUPLIEREDICODE"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("CustomerEDICode")), "", msqldr("CustomerEDICode"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("DOC_NO")), "", msqldr("DOC_NO"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("ASNDate")), "", msqldr("ASNDate"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("DESPDATE")), "", msqldr("DESPDATE"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("ASNDateAddMonth")), "", msqldr("ASNDateAddMonth"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("Gross_wt")), "", msqldr("Gross_wt"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("Net_Wt")), "", msqldr("Net_Wt"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("Measure_code")), "", msqldr("Measure_code"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("LANDING_NO")), "", msqldr("LANDING_NO"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("STPC")), "", msqldr("STPC"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("SPC")), "", msqldr("SPC"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("PD")), "", msqldr("PD"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("MRI")), "", msqldr("MRI"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("MT")), "", msqldr("MT"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("CSC")), "", msqldr("CSC"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("EQ")), "", msqldr("EQ"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("CONVEYANCE_NUMBER")), "", msqldr("CONVEYANCE_NUMBER"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("PACK_SEQ")), "", msqldr("PACK_SEQ"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("PACKING_CODE")), "", msqldr("PACKING_CODE"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("PKG_QTY")), "", msqldr("PKG_QTY"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("cust_Item_Code")), "", msqldr("cust_Item_Code"))

                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("MODELYEAR")), "", msqldr("MODELYEAR"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("Sales_Quantity")), "", msqldr("Sales_Quantity"))
                            CUMMULATIVE_QTY = Find_Value("SELECT DBO.UDF_GET_CUMMULATIVEQTY_THAI('" & gstrUNITID & "','" & msqldr("cust_Item_Code").ToString() & "','" & txtCustomerCode.Text.ToString() & "','" & strINvoiceNo.ToString() & "')")
                            strRecord = strRecord & "," & IIf(IsDBNull(CUMMULATIVE_QTY), "", CUMMULATIVE_QTY)
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("cons_measure_code")), "", msqldr("cons_measure_code"))

                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("SONumber")), "", msqldr("SONumber")) & vbCrLf
                        End While
                        'Check Directory to create ASN File
                        If Directory.Exists(strLocation) = False Then
                            Directory.CreateDirectory(strLocation)
                        End If

                        strASNFilepath = strLocation & strINvoiceNo & ".txt"
                        fs = File.Create(strASNFilepath)
                        sw = New StreamWriter(fs)
                        sw.WriteLine(strRecord)
                        sw.Close()
                        fs.Close()

                    End If
                    msqldr.Close()
                    msqlcmd.Dispose()
                Next
            End With
            Exit Function
        Catch Ex As Exception
            GenerateASNFileForGen_Motors = False
            MessageBox.Show(Ex.Message.ToString(), ResolveResString(100), MessageBoxButtons.OK)
        Finally
            msqldr.Close()
            msqlcmd.Dispose()
        End Try
    End Function
    Public Sub DisableControls()
        On Error GoTo Errorhandler
        Me.cmdLockInvoice.Revert()
        Me.cmdLockInvoice.Caption(0) = "Create"
        Me.optunCheckAll.Checked = True
        Exit Sub
Errorhandler:
        If Err.Number = 5 Then Resume Next
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function FN_Get_Folder_Path() As String
        Dim rsGetASNPath As ClsResultSetDB
        On Error GoTo Errorhandler
        rsGetASNPath = New ClsResultSetDB
        rsGetASNPath.GetResult("select isnull(ASN_HMIL_FilePath,'')as ASN_HMIL_FilePath from sales_parameter WHERE UNIT_CODE = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsGetASNPath.GetNoRows > 0 Then
            If Len(rsGetASNPath.GetValue("ASN_HMIL_FilePath")) = 0 Then
                FN_Get_Folder_Path = "False"
                rsGetASNPath.ResultSetClose()
                Exit Function
            Else
                FN_Get_Folder_Path = rsGetASNPath.GetValue("ASN_HMIL_FilePath")
            End If
        End If
        rsGetASNPath.ResultSetClose()
        rsGetASNPath = Nothing
        Exit Function
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub cmdLockInvoice_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtngrptwo.ButtonClickEventArgs) Handles cmdLockInvoice.ButtonClick
        '-------------------------------------------------------------------------------------
        'Revised By     : Manoj Kr. Vaish
        'Revised On     : 04 Feb 2009
        'Arguments      : NIL
        'Return Value   : NIL
        'Issue ID       : eMpro-20090204-27027
        'Reason         : To Create the ASN File
        '--------------------------------------------------------------------------------------
        Dim strInvoiceNo As String
        Dim strFilePath As String
        Dim blnflag As Boolean
        Dim strASNFunctionCode As String
        Dim strSupplierEDICode As String
        On Error GoTo Errorhandler
        Select Case e.ControlIndex
            Case 0
                Me.cmdLockInvoice.Caption(0) = "Create"
                If Me.spgrid.MaxRows <= 0 Then
                    Call MsgBox("Grid does not contain any invoice", MsgBoxStyle.OkOnly, ResolveResString(100))
                    Call DisableControls()
                    Exit Sub
                Else
                    With Me.spgrid
                        blnflag = False
                        For mlngCounter = 1 To .MaxRows
                            .Row = mlngCounter : .Col = InvoiceGrid.invSel
                            If CDbl(.Value) = 1 Then
                                blnflag = True
                                Exit For
                            End If
                        Next
                        If blnflag = False Then
                            Call MsgBox("Select at least one invoice for file generation", MsgBoxStyle.OkOnly, ResolveResString(100))
                            Call DisableControls()
                            Exit Sub
                        End If
                        For mlngCounter = 1 To .MaxRows
                            .Row = mlngCounter : .Col = InvoiceGrid.invSel
                            If CDbl(.Value) = 1 Then
                                .Col = InvoiceGrid.InvNo
                                strInvoiceNo = strInvoiceNo & "'" & Trim(.Text) & "',"
                            End If
                        Next
                        strInvoiceNo = Mid(strInvoiceNo, 1, Len(strInvoiceNo) - 1)
                        strFilePath = FN_Get_Folder_Path()

                        If gstrUNITID = "STH" Then
                            strFilePath = Find_Value("Select ASNFILEPATH from Customer_Mst where customer_code='" & txtCustomerCode.Text.Trim() & "' AND UNIT_CODE = '" & gstrUNITID & "'")
                            If GenerateASNFileForGen_Motors(strFilePath, strInvoiceNo) = True Then
                                MsgBox(" ASN File Generated Succesfully.File Path is - " & strFilePath, MsgBoxStyle.Information, ResolveResString(100))
                            End If
                            Call DisableControls()
                            Exit Sub
                        End If
                        strASNFunctionCode = Find_Value("Select ASNFunctionCode from Customer_Mst where customer_code='" & txtCustomerCode.Text.Trim() & "' AND UNIT_CODE = '" & gstrUNITID & "'")
                        If strFilePath = "False" And strASNFunctionCode = "A02" Then
                            MsgBox("Default Location is not defined in Sales Parameter.", MsgBoxStyle.Information, ResolveResString(100))
                            Exit Sub
                        Else
                            Select Case strASNFunctionCode
                                Case "A01"
                                    If GenerateASNForNissan(gstrASNPath, gstrASNPathForEDI) = True Then
                                        MsgBox(" ASN File Generated Succesfully.File Path is - " & gstrASNPath, MsgBoxStyle.Information, ResolveResString(100))
                                    End If
                                Case "A02"
                                    If GenerateASNFileForMahindra(strFilePath, strInvoiceNo) = True Then
                                        MsgBox(" ASN File Generated Succesfully.File Path is - " & strFilePath, MsgBoxStyle.Information, ResolveResString(100))
                                    End If
                                Case "A03"
                                    If GenerateASNFileForHyundai(strFilePath, strInvoiceNo) = True Then

                                    End If
                            End Select
                        End If
                        Call DisableControls()
                    End With
                End If
            Case 1
                Call DisableControls()
            Case 2
                Me.Close()
        End Select
        Exit Sub
Errorhandler:
        Me.cmdLockInvoice.Revert() : Me.cmdLockInvoice.Caption(0) = "POST"
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub dtToDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtToDate.ValueChanged
        On Error GoTo ErrHandler
        Me.spgrid.MaxRows = 0
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub dtFromDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtFromDate.ValueChanged
        On Error GoTo ErrHandler
        Me.spgrid.MaxRows = 0
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs)
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm") '("HLPCSTMS0001.htm")
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Function Find_Value(ByVal pstrquery As String) As String
        '----------------------------------------------------------------------------
        'Author         :   Manoj Vaish
        'Argument       :   Sql query string as strField
        'Return Value   :   selected table field value as String
        'Function       :   Return a field value from a table
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim Rs As New ADODB.Recordset
        Rs = New ADODB.Recordset
        Rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        Rs.Open(pstrquery, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
        If Rs.RecordCount > 0 Then
            If IsDBNull(Rs.Fields(0).Value) = False Then
                Find_Value = Rs.Fields(0).Value
            Else
                Find_Value = ""
            End If
        Else
            Find_Value = ""
        End If
        Rs.Close()
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function


    Private Function GenerateASNForNissan(ByVal pstrASNLoc As String, ByVal pstrASNForEDI As String) As Boolean
        'Revised By     : Manoj Kr. Vaish
        'Revised On     : 23 Jun 2009
        'Arguments      : File Location,EDI File Location,Invoice No
        'Return Value   : True/False
        'Issue ID       : eMpro-20090610-32326
        'Reason         : To Generate NISSAN ASN File for selected Invoice
        '--------------------------------------------------------------------------------------
        Dim strsql As String
        Dim strRecord As String
        Dim strINvoiceNo As String
        Dim fs As FileStream
        Dim sw As StreamWriter
        Dim strASNFilepath As String
        Dim strASNFilepathforEDI As String
        Try
            GenerateASNForNissan = True
            With spgrid
                For mlngCounter = 1 To .MaxRows
                    .Row = mlngCounter : .Col = InvoiceGrid.invSel
                    If CDbl(.Value) = 1 Then
                        .Col = InvoiceGrid.InvNo
                        strINvoiceNo = Trim(.Text)
                        strsql = "SELECT SP.SUPPLIEREDICODE,CM.Customer_EDICode,SC.DOC_NO,DBO.UFN_GET_YYYYMMDDHHMM(INVOICE_DATE,INVOICE_TIME) AS INVOICE_TIME,"
                        strsql = strsql & " CM.DOCK_CODE,CM.CUST_VENDOR_CODE,SC.CARRIAGE_NAME,SC.PORT_OF_DISCHARGE,SC.TRANSPORT_TYPE,"
                        strsql = strsql & " SC.VEHICLE_NO,CIM.PACKING_CODE,CIM.CONTAINER,(SD.TO_BOX-SD.FROM_BOX)+1 AS NO_PACKAGE,"
                        strsql = strsql & " SD.BINQUANTITY,SD.CUST_ITEM_CODE,SD.SALES_QUANTITY,SD.MEASURE_CODE,SC.CUST_REF,MD.DSNO AS RAN_NO FROM SALESCHALLAN_DTL SC"
                        strsql = strsql & " INNER JOIN SALES_DTL SD ON SC.DOC_NO=SD.DOC_NO AND SC.LOCATION_CODE=SD.LOCATION_CODE AND SC.UNIT_CODE=SD.UNIT_CODE"
                        strsql = strsql & " INNER JOIN SALES_PARAMETER SP ON SC.LOCATION_CODE=SP.COMPANY_CODE AND SC.UNIT_CODE=SP.UNIT_CODE"
                        strsql = strsql & " INNER JOIN CUSTOMER_MST CM ON SC.ACCOUNT_CODE=CM.CUSTOMER_CODE  AND SC.UNIT_CODE=CM.UNIT_CODE"
                        strsql = strsql & " LEFT OUTER JOIN CUSTITEM_MST CIM ON SD.ITEM_CODE=CIM.ITEM_CODE AND SD.CUST_ITEM_CODE=CIM.CUST_DRGNO AND SD.UNIT_CODE=CIM.UNIT_CODE"
                        strsql = strsql & " AND SC.ACCOUNT_CODE=CIM.ACCOUNT_CODE LEFT OUTER JOIN MKT_INVDSHISTORY MD ON SC.DOC_NO=MD.DOC_NO AND SC.UNIT_CODE=MD.UNIT_CODE"
                        strsql = strsql & " AND SC.LOCATION_CODE=MD.LOCATION_CODE AND SD.ITEM_CODE=MD.ITEM_CODE AND SD.CUST_ITEM_CODE=MD.CUST_PART_CODE"
                        strsql = strsql & " WHERE SC.BILL_FLAG=1 AND SC.CANCEL_FLAG=0 AND SC.DOC_NO='" & strINvoiceNo & "' AND SC.UNIT_CODE = '" & gstrUNITID & "'"
                        msqlcmd = New SqlCommand(strsql, msqlcon)
                        msqldr = msqlcmd.ExecuteReader()
                        strRecord = ""
                        While msqldr.Read()
                            strRecord = strRecord & IIf(IsDBNull(msqldr("SUPPLIEREDICODE")), "", msqldr("SUPPLIEREDICODE"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("Customer_EDICode")), "", msqldr("Customer_EDICode"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("DOC_NO")), "", msqldr("DOC_NO"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("INVOICE_TIME")), "", msqldr("INVOICE_TIME"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("INVOICE_TIME")), "", msqldr("INVOICE_TIME"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("DOCK_CODE")), "", msqldr("DOCK_CODE"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("CUST_VENDOR_CODE")), "", msqldr("CUST_VENDOR_CODE"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("CARRIAGE_NAME")), "", msqldr("CARRIAGE_NAME"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("PORT_OF_DISCHARGE")), "", msqldr("PORT_OF_DISCHARGE"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("TRANSPORT_TYPE")), "", msqldr("TRANSPORT_TYPE"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("VEHICLE_NO")), "", msqldr("VEHICLE_NO"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("PACKING_CODE")), "", msqldr("PACKING_CODE"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("CONTAINER")), "", msqldr("CONTAINER"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("NO_PACKAGE")), "", msqldr("NO_PACKAGE"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("BINQUANTITY")), "", msqldr("BINQUANTITY"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("CUST_ITEM_CODE")), "", msqldr("CUST_ITEM_CODE"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("SALES_QUANTITY")), "", msqldr("SALES_QUANTITY"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("MEASURE_CODE")), "", msqldr("MEASURE_CODE"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("CUST_REF")), "", msqldr("CUST_REF"))
                            strRecord = strRecord & "," & IIf(IsDBNull(msqldr("RAN_NO")), "", msqldr("RAN_NO")) & vbCrLf
                        End While
                        'Check Directory to create ASN File
                        If Directory.Exists(pstrASNLoc) = False Then
                            Directory.CreateDirectory(pstrASNLoc)
                        End If
                        'Check Directory to create ASN File
                        If Directory.Exists(pstrASNForEDI) = False Then
                            Directory.CreateDirectory(pstrASNForEDI)
                        End If
                        strASNFilepath = pstrASNLoc & "\ASN" & strINvoiceNo & ".csv"
                        strASNFilepathforEDI = pstrASNForEDI & "\ASN" & strINvoiceNo & ".csv"
                        fs = File.Create(strASNFilepath)
                        sw = New StreamWriter(fs)
                        sw.WriteLine(strRecord)
                        sw.Close()
                        fs.Close()
                        If File.Exists(strASNFilepathforEDI) = False Then
                            File.Copy(strASNFilepath, strASNFilepathforEDI)
                        End If
                    End If
                    msqldr.Close()
                    msqlcmd.Dispose()
                Next
            End With
            Exit Function
        Catch Ex As Exception
            GenerateASNForNissan = False
            MessageBox.Show(Ex.Message.ToString(), ResolveResString(100), MessageBoxButtons.OK)
        Finally
            msqldr.Close()
            msqlcmd.Dispose()
        End Try
    End Function
    Private Function GenerateASNFileForHyundai(ByVal pstrFileLocation As String, ByVal pstrInvoiceNo As String) As Boolean
        'Revised By     : prashant Rajpal
        'Revised On     : 01 june 2012
        'Issue ID       : 10229992 
        'Reason         : To Generate ASN File for selected Invoice for HUNDAI
        '--------------------------------------------------------------------------------------
        Dim strsql As String
        Dim rsGetASNData As ClsResultSetDB
        Dim Obj_FSO As Scripting.FileSystemObject
        Dim strLocation As String
        Dim strFileName As String
        Dim strRecord As String
        Dim intLineNo As Short
        Dim BalanceQty As Integer
        On Error GoTo Err_Handler
        GenerateASNFileForHyundai = True
        '----------------------------------------
        Obj_FSO = New Scripting.FileSystemObject
        If Not Obj_FSO.FolderExists(pstrFileLocation) Then
            Obj_FSO.CreateFolder(pstrFileLocation)
        End If
        If Mid(Trim(pstrFileLocation), Len(Trim(pstrFileLocation))) <> "\" Then
            strLocation = pstrFileLocation & "\"
        End If
        strFileName = "ASN" & VB6.Format(GetServerDateTime(), "ddMMyyyyhhmmss") & ".txt"
        strFileName = strLocation & strFileName
        'Kill(strLocation & "*.csv")
        FileClose(1)
        If Dir(strFileName) <> "" Then
            Kill(strFileName)
        End If

        FileOpen(1, strFileName, OpenMode.Append)
        Obj_FSO = Nothing
        rsGetASNData = New ClsResultSetDB
        strsql = "select distinct c.Cust_Vendor_Code,PLANT_CODE =(SELECT TOP 1 PLANT_C FROM SO_UPLD_HDR H ,SO_UPLD_DTL D WHERE H.UNIT_CODE = D.UNIT_CODE "
        strsql = strsql & " AND H.DOC_NO=D.DOC_NO AND H.CUST_CODE=A.ACCOUNT_CODE AND D.SALESORDER="
        strsql = strsql & " CASE WHEN LEN(B.EXTERNAL_SALESORDER_NO)>0  THEN B.EXTERNAL_SALESORDER_NO ELSE A.CUST_REF END) ,CI.Gate_no,a.Invoice_Date,convert(varchar(8),a.ent_dt,108) as invoice_time ,convert(varchar(8),dateadd(hh,2,a.ent_dt),108) Estimatedtime ,A.LorryNo_Date,A.Doc_no,A.Vehicle_No,A.Carriage_Name,"
        strsql = strsql & " B.Sales_Quantity,A.Total_Amount,B.Excise_Tax ,CASE WHEN LEN(B.EXTERNAL_SALESORDER_NO)>0 THEN B.EXTERNAL_SALESORDER_NO ELSE   A.CUST_REF END AS CUST_REF ,c.office_phone ,b.cust_item_code ,b.item_code ,sales_quantity from Saleschallan_dtl a"
        strsql = strsql & " Inner join Sales_Dtl b on a.doc_no=b.doc_no AND a.UNIT_CODE=b.UNIT_CODE "
        strsql = strsql & " Inner join customer_mst C on C.customer_code=a.account_code AND a.UNIT_CODE=C.UNIT_CODE "
        strsql = strsql & " Inner join custitem_mst CI on CI.UNIT_CODE=b.UNIT_CODE and CI.account_code=a.account_code and CI.item_code=b.item_code and b.cust_item_code=CI.Cust_Drgno"
        strsql = strsql & " where a.bill_flag=1 and a.cancel_flag=0 and CI.ACTIVE=1 AND a.UNIT_CODE = '" & gstrUNITID & "'"
        strsql = strsql & " and a.doc_no in(" & pstrInvoiceNo & ") order by invoice_date "
        rsGetASNData.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        'strRecord = "Vendor   Invoiceno. Plant Gate        Departure date   Departure time Arrive plandate   Estimated time of arrival  Car Number Delivery Person      Phone Number      PO Number    Po Item  Material  Remain Qty  Deliver Qty"
        strRecord = "Vendor" & vbTab & "Invoiceno." & vbTab & "Plant " & vbTab & "Gate " & vbTab & "Departure date" & vbTab & "Departure time" & vbTab & "Arrive plandate" & vbTab & "Estimated time of arrival" & vbTab & "Car Number" & vbTab & "Delivery Person" & vbTab & "Phone Number" & vbTab & "PO Number" & vbTab & "Po Item " & vbTab & "Material " & vbTab & "Remain Qty" & vbTab & "Delivery Qty" & vbTab & "Manufacturing Date"
        PrintLine(1, strRecord) : intLineNo = intLineNo + 1
        If rsGetASNData.GetNoRows > 0 Then
            rsGetASNData.MoveFirst()
            Do While Not rsGetASNData.EOFRecord
                strRecord = ""
                '10236562
                strRecord = IIf(IsDBNull(rsGetASNData.GetValue("Cust_Vendor_Code")), "", rsGetASNData.GetValue("Cust_Vendor_Code")) & vbTab
                strRecord = strRecord & "" & IIf(IsDBNull(rsGetASNData.GetValue("Doc_no")), "", rsGetASNData.GetValue("Doc_no")) & vbTab
                strRecord = strRecord & "" & IIf(Len(rsGetASNData.GetValue("plant_code")) = 0, "-", rsGetASNData.GetValue("plant_code")) & vbTab
                strRecord = strRecord & "" & IIf(Len(rsGetASNData.GetValue("Gate_no")) = 0, "", rsGetASNData.GetValue("Gate_no")) & vbTab
                'strRecord = Set_Line_Width(strRecord, 40) & rsGetASNData.GetValue("Gate_no")
                strRecord = strRecord & "" & IIf(IsDBNull(rsGetASNData.GetValue("Invoice_Date")), "", VB6.Format(rsGetASNData.GetValue("Invoice_Date"), "ddmmyyyy")) & vbTab
                strRecord = strRecord & "" & IIf(IsDBNull(rsGetASNData.GetValue("invoice_time")), "", VB6.Format(rsGetASNData.GetValue("invoice_time"), "hhmmss")) + CStr("   ") & vbTab
                strRecord = strRecord & "" & IIf(IsDBNull(rsGetASNData.GetValue("Invoice_Date")), "", VB6.Format(rsGetASNData.GetValue("Invoice_Date"), "ddmmyyyy")) & vbTab
                strRecord = strRecord & "" & IIf(IsDBNull(rsGetASNData.GetValue("Estimatedtime")), "", Replace(rsGetASNData.GetValue("Estimatedtime"), ":", "")) + CStr("   ") + CStr("   ") + CStr("   ") + CStr("   ") + CStr("   ") + CStr("   ") & vbTab
                'strRecord = strRecord & "" & IIf(IsDBNull(rsGetASNData.GetValue("Invoice_Date")), "", VB6.Format(rsGetASNData.GetValue("Invoice_Date"), "ddmmyyyy"))& vbTab
                strRecord = strRecord & "" & IIf(IsDBNull(rsGetASNData.GetValue("Vehicle_No")), "", rsGetASNData.GetValue("Vehicle_No")) + CStr("   ") + CStr("   ") & vbTab
                strRecord = strRecord & "" & IIf(IsDBNull(rsGetASNData.GetValue("Carriage_Name")), "", rsGetASNData.GetValue("Carriage_Name")) + CStr("   ") + CStr("   ") & vbTab
                strRecord = strRecord & "" & IIf(Len(rsGetASNData.GetValue("Office_phone")) = 0, "", rsGetASNData.GetValue("Office_phone")) & vbTab
                strRecord = strRecord & "" & IIf(IsDBNull(rsGetASNData.GetValue("cust_ref")), "", rsGetASNData.GetValue("Cust_ref")) & vbTab
                strRecord = strRecord & "" & "1" + CStr("   ") + CStr("   ") + CStr("   ") & vbTab
                strRecord = strRecord & "" & IIf(IsDBNull(rsGetASNData.GetValue("cust_item_code")), "", rsGetASNData.GetValue("Cust_item_code")) + CStr("   ") & vbTab
                BalanceQty = Find_Value("SELECT DBO.UDF_PENDINGSCHQTY_HYUNDAI_ASN('" & gstrUNITID & "','" & rsGetASNData.GetValue("Cust_ref").ToString() & "','" & rsGetASNData.GetValue("Cust_item_code").ToString() & "','" & rsGetASNData.GetValue("item_code").ToString() & "')")
                strRecord = strRecord & "" & CStr(BalanceQty) + CStr("   ") + CStr("   ") & vbTab
                strRecord = strRecord & "" & IIf(IsDBNull(rsGetASNData.GetValue("sales_quantity")), "", CInt(rsGetASNData.GetValue("sales_quantity"))) & vbTab
                strRecord = strRecord & "" & CStr("") & vbTab
                'strRecord = Set_Line_Width(strRecord, 3) & rsGetASNData.GetValue("Cust_Vendor_Code") & vbTab
                'strRecord = Set_Line_Width(strRecord, 7) & rsGetASNData.GetValue("Doc_no") & vbTab
                'strRecord = Set_Line_Width(strRecord, 18) & rsGetASNData.GetValue("plant_code") & vbTab
                'strRecord = Set_Line_Width(strRecord, 25) & rsGetASNData.GetValue("Gate_no").ToString & vbTab
                'strRecord = Set_Line_Width(strRecord, 30) & VB6.Format(rsGetASNData.GetValue("Invoice_Date"), "ddmmyyyy") & vbTab
                'strRecord = Set_Line_Width(strRecord, 39) & Replace(rsGetASNData.GetValue("Invoice_time").ToString, ":", "") & vbTab
                'strRecord = Set_Line_Width(strRecord, 46) & VB6.Format(rsGetASNData.GetValue("Invoice_Date"), "ddmmyyyy") & vbTab
                'strRecord = Set_Line_Width(strRecord, 58) & Replace(rsGetASNData.GetValue("Estimatedtime").ToString, ":", "") & vbTab
                'strRecord = Set_Line_Width(strRecord, 81) & rsGetASNData.GetValue("Vehicle_No") & vbTab
                'strRecord = Set_Line_Width(strRecord, 91) & rsGetASNData.GetValue("Carriage_Name") & vbTab
                'strRecord = Set_Line_Width(strRecord, 103) & rsGetASNData.GetValue("Office_phone") & vbTab
                'strRecord = Set_Line_Width(strRecord, 116) & rsGetASNData.GetValue("Cust_ref") & vbTab
                'strRecord = Set_Line_Width(strRecord, 131) & "1" & vbTab
                'strRecord = Set_Line_Width(strRecord, 137) & rsGetASNData.GetValue("Cust_item_code") & vbTab
                'strRecord = Set_Line_Width(strRecord, 151) & BalanceQty & vbTab
                'strRecord = Set_Line_Width(strRecord, 163) & CInt(rsGetASNData.GetValue("sales_quantity")) & vbTab
                '10236562 DONE 
                PrintLine(1, strRecord) : intLineNo = intLineNo + 1
                rsGetASNData.MoveNext()
            Loop
            rsGetASNData.ResultSetClose()
            rsGetASNData = Nothing
        Else
            MsgBox("No Invoice Records found to generate the File.", MsgBoxStyle.Information, ResolveResString(100))
            FileClose(1)
            Kill(strFileName)
            rsGetASNData.ResultSetClose()
            rsGetASNData = Nothing
            GenerateASNFileForHyundai = False
            Exit Function
        End If
        FileClose(1)
        If GenerateASNFileForHyundai = True Then
            MsgBox(" ASN File Generated Succesfully.File Path is - " & strFileName, MsgBoxStyle.Information, ResolveResString(100))
        End If
        Exit Function
Err_Handler:
        If Err.Number = 55 Then
            MsgBox("File Already Open, Cann't Generate the ASN File.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            GenerateASNFileForHyundai = False
            Exit Function
        End If
        GenerateASNFileForHyundai = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Function Set_Line_Width(ByRef strval As String, ByRef intLen As Short, Optional ByRef dblLastNo As Double = 0, Optional ByRef blnPrint_Zero_Amt As Boolean = False) As String
        '--------------------------------------------------------------------------------------
        'Revised By     : Prashant Rajpal
        'Revised On     : 01 june 2012
        'Issue ID       : 10229992 
        'Reason         : To Generate ASN File for SPACE ALLOCATION
        '--------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strval1 As String
        strval1 = strval & IIf(dblLastNo > 0 Or blnPrint_Zero_Amt = True, VB6.Format(dblLastNo, "####0.00"), "")
        If Len(strval1) < intLen Then
            strval1 = strval & Space(intLen - Len(strval1)) & IIf(dblLastNo > 0 Or blnPrint_Zero_Amt = True, VB6.Format(dblLastNo, "####0.00"), "")
        Else
            If dblLastNo > 0 Then
                strval1 = Mid(strval, 1, intLen - Len(VB6.Format(dblLastNo, "####0.00"))) & " " & IIf(dblLastNo > 0 Or blnPrint_Zero_Amt = True, VB6.Format(dblLastNo, "###,##,##0.00"), "")
            Else
                strval1 = Mid(strval, 1, intLen)
            End If
        End If
        Set_Line_Width = strval1
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
End Class