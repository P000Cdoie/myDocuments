Option Strict Off
Option Explicit On
Imports System.Data.SqlClient
Imports System.IO
'Imports System.Collections.Generic
Friend Class frmMKTTRN0096
    Inherits System.Windows.Forms.Form
    '---------------------------------------------------------------------------
    'Copyright          :   MIND Ltd.
    'Form Name          :   frmMKTTRN0096
    'Created By         :   ABHIJIT KUMAR SINGH
    'Created on         :   20 SEP 2017
    'Description        :   ASN Text File Generation for MP2
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
        CarrierIdentificationCode = 6
        EquipmentDescriptionCode = 7
        EquipmentNumber = 8
        AirbillNumber = 9
        Billoflading = 10

    End Enum
    Private Sub cmdCustHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCustHelp.Click
        On Error GoTo ErrHandler
        Dim strHelp() As String
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select CUSTOMER_CODE,CUST_NAME  FROM  Customer_Mst where unit_code='" & gstrUNITID & "'", "Help", 2)
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
        txtInvoice_no.Text = ""
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
        Call ShowHelp("underconstruction.htm")
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SetGridCells()
        On Error GoTo Errorhandler
        With Me.spgrid
            .MaxRows = 0 : .MaxCols = 5
            .Row = 0 : .set_RowHeight(0, 300)
            .Col = InvoiceGrid.invSel : .Text = "Select Invoice" : .set_ColWidth(InvoiceGrid.invSel, 1500) : .ColHidden = True
            .Col = InvoiceGrid.InvNo : .Text = "Invoice No" : .set_ColWidth(InvoiceGrid.InvNo, 1600)
            .Col = InvoiceGrid.invDate : .Text = "Invoice Date" : .set_ColWidth(InvoiceGrid.invDate, 1600)
            .Col = InvoiceGrid.invTypeDesc : .Text = "Invoice Type" : .set_ColWidth(InvoiceGrid.invTypeDesc, 1800)
            .Col = InvoiceGrid.invSubTypeDesc : .Text = "Invoice Category" : .set_ColWidth(InvoiceGrid.invSubTypeDesc, 1800)
            If LblCustomerName.Text.Contains("IAC ") Then
                .MaxRows = 0 : .MaxCols = 10
                .Col = InvoiceGrid.CarrierIdentificationCode : .Text = "Carrier Id Code" : .set_ColWidth(InvoiceGrid.CarrierIdentificationCode, 1600)
                .Col = InvoiceGrid.EquipmentDescriptionCode : .Text = "Equipment Desc Code" : .set_ColWidth(InvoiceGrid.EquipmentDescriptionCode, 1600)
                .Col = InvoiceGrid.EquipmentNumber : .Text = "Equipment Number" : .set_ColWidth(InvoiceGrid.EquipmentNumber, 1600)
                .Col = InvoiceGrid.AirbillNumber : .Text = "Air bill Number" : .set_ColWidth(InvoiceGrid.AirbillNumber, 1600)
                .Col = InvoiceGrid.Billoflading : .Text = "Bill of lading" : .set_ColWidth(InvoiceGrid.Billoflading, 1600)
            End If
        End With
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub ShowPendingInvoices()
        On Error GoTo Errorhandler
        Dim rsobject As New ADODB.Recordset
        Dim cmdObject As New ADODB.Command
        Dim strsql As String
        rsobject.CursorType = ADODB.CursorTypeEnum.adOpenForwardOnly
        With cmdObject
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandText = "GETINVOICE_FOR_ASN_MP2"
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
                    If LblCustomerName.Text.Contains("IAC ") Then
                        .Col = InvoiceGrid.CarrierIdentificationCode : .Text = ""
                        .Col = InvoiceGrid.EquipmentDescriptionCode : .Text = ""
                        .Col = InvoiceGrid.EquipmentNumber : .Text = ""
                        .Col = InvoiceGrid.AirbillNumber : .Text = ""
                        .Col = InvoiceGrid.Billoflading : .Text = ""
                    End If
                    rsobject.MoveNext() : mlngCounter = mlngCounter + 1
                Loop
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
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
            If LblCustomerName.Text.Contains("IAC ") Then
                .Col = InvoiceGrid.CarrierIdentificationCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 4
                .Col = InvoiceGrid.EquipmentDescriptionCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 2
                .Col = InvoiceGrid.EquipmentNumber : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 7
                .Col = InvoiceGrid.AirbillNumber : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 30
                .Col = InvoiceGrid.Billoflading : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 30
            End If
            .Row = .MaxRows : .Row2 = .MaxRows : .Col = InvoiceGrid.InvNo
        End With
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Public Sub DisableControls()
        On Error GoTo Errorhandler
        Me.cmdLockInvoice.Revert()
        Me.cmdLockInvoice.Caption(0) = "Create"
        Exit Sub
Errorhandler:
        If Err.Number = 5 Then Resume Next
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub cmdLockInvoice_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtngrptwo.ButtonClickEventArgs) Handles cmdLockInvoice.ButtonClick
        Dim strInvoiceNo As String
        Dim strFilePath As String
        Dim blnflag As Boolean
        Dim strASNFunctionCode As String
        Dim strSupplierEDICode As String

        Dim strCarrierIdentificationCode As Object
        Dim strEquipmentDescriptionCode As Object
        Dim strEquipmentNumber As Object
        Dim strAirBillNumber As Object
        Dim strBillOfLading As Object
        On Error GoTo Errorhandler
        Select Case e.ControlIndex
            Case 0
                Me.cmdLockInvoice.Caption(0) = "Create"

                If txtInvoice_no.Text = "" Then
                    MessageBox.Show("Invoice no. can't be blank", "eMPRO", MessageBoxButtons.OK)
                    Exit Sub
                End If
                Dim strFormatType = SqlConnectionclass.ExecuteScalar("Select ASNFORMAT_TYPE	FROM AUTOASN_MAPPING WHERE CUSTOMER_CODE='" & txtCustomerCode.Text & "'  and unit_code='" & gstrUNITID & "'")


                If Me.spgrid.MaxRows <= 0 Then
                    Call MsgBox("Grid does not contain any invoice", MsgBoxStyle.OkOnly, ResolveResString(100))
                    Call DisableControls()
                    Exit Sub
                Else
                    With Me.spgrid '
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

                        Dim strSRV As String
                        Dim strmessage As String

                        Dim result As Integer = MessageBox.Show("Do you want to generate ASN Text File..!!", "eMPRO", MessageBoxButtons.YesNo)
                        If result = DialogResult.No Then
                            MessageBox.Show("You have cancelled ASN Text File Generation..!!", "eMPRO", MessageBoxButtons.OK)
                            txtCustomerCode.Text = ""
                            txtInvoice_no.Text = ""
                            LblCustomerName.Text = ""
                            spgrid.MaxRows = 0
                            Exit Sub
                        End If
                        If LblCustomerName.Text.Contains("IAC ") Then
                            strCarrierIdentificationCode = Nothing
                            strEquipmentDescriptionCode = Nothing
                            strEquipmentNumber = Nothing
                            strAirBillNumber = Nothing
                            strBillOfLading = Nothing
                            With spgrid
                                .Row = .ActiveRow
                                '.Col = .ActiveCol
                                'If .Col = 6 Then
                                '    strCarrierIdentificationCode = .Text
                                'Else
                                '    strCarrierIdentificationCode = ""
                                'End If
                                Call .GetText(InvoiceGrid.CarrierIdentificationCode, .Row, strCarrierIdentificationCode)
                                Call .GetText(InvoiceGrid.EquipmentDescriptionCode, .Row, strEquipmentDescriptionCode)
                                Call .GetText(InvoiceGrid.EquipmentNumber, .Row, strEquipmentNumber)
                                Call .GetText(InvoiceGrid.AirbillNumber, .Row, strAirBillNumber)
                                Call .GetText(InvoiceGrid.Billoflading, .Row, strBillOfLading)

                            End With
                            If CStr(strCarrierIdentificationCode) = "" Then
                                MessageBox.Show("Please Enter Carrier Identification Code", "eMPRO", MessageBoxButtons.OK)
                                Exit Sub
                            End If
                            If CStr(strEquipmentDescriptionCode) = "" Then
                                MessageBox.Show("Please Enter Equipment Description Code", "eMPRO", MessageBoxButtons.OK)
                                Exit Sub
                            End If
                            If CStr(strEquipmentNumber) = "" Then
                                MessageBox.Show("Please Enter Equipment Number", "eMPRO", MessageBoxButtons.OK)
                                Exit Sub
                            End If
                            If CStr(strAirBillNumber) = "" Then
                                MessageBox.Show("Please Enter Air Bill Number", "eMPRO", MessageBoxButtons.OK)
                                Exit Sub
                            End If
                            If CStr(strBillOfLading) = "" Then
                                MessageBox.Show("Please Enter Bill Of Lading", "eMPRO", MessageBoxButtons.OK)
                                Exit Sub
                            End If
                            strmessage = ASN_856_MP2(gstrUNITID, txtCustomerCode.Text, txtInvoice_no.Text, CStr(strCarrierIdentificationCode), CStr(strEquipmentDescriptionCode), CStr(strEquipmentNumber), CStr(strAirBillNumber), CStr(strBillOfLading))
                        ElseIf isGalliaASN() = True Then
                            strmessage = ASN_Gallia_MP2(gstrUNITID, txtCustomerCode.Text, txtInvoice_no.Text, strSRV)
                        ElseIf isYazakiASN() = True Then
                            strmessage = GenerateYazakiASN(gstrUNITID, txtCustomerCode.Text, txtInvoice_no.Text, strSRV)
                        ElseIf Len(strFormatType) > 0 Then
                            strmessage = ASN_DESADV_96A(gstrUNITID, txtCustomerCode.Text, txtInvoice_no.Text, strFormatType)
                        ElseIf isBoschASN() = True Then
                            strmessage = ASN_Bosch_MS1(gstrUNITID, txtCustomerCode.Text, txtInvoice_no.Text, strSRV)
                        Else
                            strmessage = ASN_4913_MP2(gstrUNITID, txtCustomerCode.Text, txtInvoice_no.Text, strSRV)
                        End If

                        MessageBox.Show(strmessage, "eMPRO", MessageBoxButtons.OK)
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

    Public Function ASN_856_MP2(ByVal UnitCode As String, ByVal pstrAccountCode As String, ByVal pstrInvoice As String, ByVal strCarrierIdentificationCode As String, ByVal strEquipmentDescriptionCode As String, ByVal strEquipmentNumber As String, ByVal strAirBillNumber As String, ByVal strBillOfLading As String) As String
        Try
            Dim strLocation As String
            Dim strFileName As String
            Dim intLineCount As Short
            Dim intLineWidth As Short
            Dim intLineNo As Short
            Dim strSql As String
            Dim strsql2 As String

            Dim strRecord_H As String
            Dim strRecord_P As String
            Dim strRecord_D As String
            Dim cnt As Integer = 0
            Dim cnt_T As Integer = 0
            Dim cnt_T1 As Integer = 0
            Dim cnt_T2 As Integer = 0

            Dim oSqlCmdLG As New SqlCommand
            Dim adap As SqlDataAdapter
            Dim ds As DataSet

            Dim ds_T As DataSet
            Dim adap_T As SqlDataAdapter

            Dim Item_Code As String = ""
            Dim IP_Address As String = ""

            Dim strSql_Pallet As String = ""
            Dim ds_Pallet As DataSet
            Dim adap_Pallet As SqlDataAdapter
            Dim cnt_Pallet As Integer

            Dim strSql_BOX As String = "'"
            Dim adap_BOX As SqlDataAdapter
            Dim ds_BOX As DataSet
            Dim cnt_box As Integer

            Dim min1, max1, no_of_package As Integer
            min1 = 0
            max1 = 0
            no_of_package = 0

            Dim cnt_pallet_no As Integer


            Dim sqlcmd_T As New SqlCommand
            Dim sqlcmd_ds As New SqlCommand
            Dim sqlcmd_Pallet As New SqlCommand
            Dim sqlcmd_BOX As New SqlCommand

            IP_Address = gstrIpaddressWinSck
            ' IP_Address = "1.01.12.13"
            strLocation = Trim(Find_Value("SELECT ISNULL(TextFileDefaultLocation,'') FROM SALES_PARAMETER_ASN_MP2 WHERE UNIT_CODE='" & UnitCode & "'"))
            ''strLocation = "C:\MUL_Txtfile"

            If Len(strLocation) = 0 Then
                ASN_856_MP2 = "FALSE|Default location not defined in sales_parameter."
                Exit Function
            Else
                If Mid(Trim(strLocation), Len(Trim(strLocation))) <> "\" Then
                    strLocation = strLocation & "\"
                End If
                If Directory.Exists(strLocation) = False Then
                    Directory.CreateDirectory(strLocation)
                End If
                'strFileName = strLocation & pstrInvoice & "_" & VB6.Format(Now, "dd-MMM-yyyy") & ".txt"
                strFileName = strLocation & UnitCode & "_" & pstrInvoice & ".txt"
                'On Error Resume Next
                'Kill(strLocation & "*.txt")
                'If System.IO.File.Exists(strFileName) Then
                '    System.IO.File.Delete(strFileName)
                If System.IO.File.Exists(strFileName) Then
                    Kill(strFileName)
                    FileClose(1)
                End If

                'On Error GoTo ErrHandler
                FileOpen(1, strFileName, OpenMode.Append)

            End If
            intLineCount = 60
            intLineWidth = 180
            Dim rsTextfileNonDI As ADODB.Recordset
            Dim strSalesTextFile As String

            ''pstrInvoice = Mid(Trim(pstrInvoice), 2, Len(Trim(pstrInvoice)) - 2)
            'If Len(pstrInvoice) > 0 Then

            strSql = ""
            strSql = "SELECT sales_dtl.Item_Code FROM sales_dtl INNER JOIN SalesChallan_Dtl ON sales_dtl.UNIT_CODE = "
            strSql = strSql & "SalesChallan_Dtl.UNIT_CODE AND sales_dtl.Location_Code = SalesChallan_Dtl.Location_Code "
            strSql = strSql & "AND    sales_dtl.Doc_No = SalesChallan_Dtl.Doc_No WHERE (sales_dtl.Doc_No = '" & pstrInvoice & "')"
            strSql = strSql & "AND (sales_dtl.UNIT_CODE = '" & UnitCode & "') AND (saleschallan_dtl.Account_Code='" & pstrAccountCode & "')"

            'adap_T = New SqlDataAdapter(strSql, SqlConnectionclass.GetConnection)
            '    ds_T = New DataSet
            '    adap_T.Fill(ds_T)

            ds_T = New DataSet
            ' Dim sqlcmd_T As New SqlCommand
            With sqlcmd_T
                .CommandText = strSql
                .CommandType = CommandType.Text
            End With
            ds_T = SqlConnectionclass.GetDataSet(sqlcmd_T)


            If ds_T.Tables(0).Rows.Count >= 1 Then
                'If ds_T.Tables(0).Rows.Count > 1 Then
                'HEADER'

                ''''
                '/////// Pallete sequence update

                'Dim STR_PALLET_DELETE As String
                'Dim SQLCMD_PALLET_DELETE As New SqlCommand

                'STR_PALLET_DELETE = "DELETE FROM MULTIPLE_PALLET_STATUS WHERE UNIT_CODE= '" & UnitCode & "' AND DOC_NO='" & pstrInvoice & "' "

                'With SQLCMD_PALLET_DELETE
                '    .CommandText = STR_PALLET_DELETE
                '    .CommandType = CommandType.Text
                'End With
                'SqlConnectionclass.ExecuteNonQuery(SQLCMD_PALLET_DELETE)


                'Dim STR_PALLET_INSERT As String
                'Dim SQLCMD_PALLET_INSERT As New SqlCommand


                'STR_PALLET_INSERT = "INSERT INTO MULTIPLE_PALLET_STATUS(UNIT_CODE,DOC_NO,PALLET_NO,UPDATED_STATUS) " & _
                '" SELECT DISTINCT UNIT_CODE,INVOICENO,PALLETNO,0 FROM VDA_ASN_INVLABELS WHERE INVOICENO='" & pstrInvoice & "' AND UNIT_CODE='" & UnitCode & "'"

                'With SQLCMD_PALLET_INSERT
                '    .CommandText = STR_PALLET_INSERT
                '    .CommandType = CommandType.Text
                'End With
                'SqlConnectionclass.ExecuteNonQuery(SQLCMD_PALLET_INSERT)

                '/////// Pallete sequence update
                ''''
                ds = New DataSet
                'For cnt_T = 0 To ds_T.Tables(0).Rows.Count - 1
                'Item_Code = ds_T.Tables(0).Rows(cnt_T).Item("Item_Code")
                With oSqlCmdLG
                    .Parameters.Clear()
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 0
                    .CommandText = "USP_IAC_SAVE_ASNFILEDATA"
                    .Parameters.AddWithValue("@UNIT_CODE", UnitCode.Trim)
                    '.Parameters.AddWithValue("@CUSTOMERCODE", pstrAccountCode.Trim)
                    .Parameters.AddWithValue("@invoice_no", pstrInvoice.Trim)
                    .Parameters.AddWithValue("@Carrier_Identification_Code", strCarrierIdentificationCode)
                    .Parameters.AddWithValue("@Equipment_Description_Code", strEquipmentDescriptionCode)
                    .Parameters.AddWithValue("@Equipment_Number", strEquipmentNumber)
                    .Parameters.AddWithValue("@Air_bill_number", strAirBillNumber)
                    .Parameters.AddWithValue("@Bill_of_lading", strBillOfLading)
                    '.Parameters.AddWithValue("@ITEM_CODE", Item_Code.Trim)
                    .Parameters.AddWithValue("@IP_ADDRESS", IP_Address.Trim)
                    ds = SqlConnectionclass.GetDataSet(oSqlCmdLG)
                End With
                'Next

                If ds.Tables(0).Rows.Count = 0 Then
                    ASN_856_MP2 = "FALSE|No Invoice Records were Found."
                    FileClose(1)
                    Exit Function
                End If
                FileClose(1)
                FileOpen(1, strFileName, OpenMode.Append)
                strRecord_H = ""

                If ds.Tables(0).Rows.Count > 0 And ds.Tables.Count = 4 Then
                    For cnt = 0 To ds.Tables(0).Rows.Count - 1
                        strRecord_H = ds.Tables(0).Rows(cnt).Item("Header")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("Supplier_Code")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("Customer_Code")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("Purpose")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("Invoice_ASN_Number")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("ASN_Date")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("ASN_Time")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("Shipped_Date")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("Shipped_Time")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("Shipment_Hierarchy_Level")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("GrossWt_of_Shipment")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("NetWt_of_Shipment")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("Weight_UOM")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("Carrier_Identification_Code")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("Mode_of_Transport")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("Equipment_Description_Code")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("Equipment_Number")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("Air_bill_number")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("Bill_of_lading")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("Packing_List_Number")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("Receiving_plant_code")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("Ship_from_location_Code")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("Manufacturer_Code")
                        strRecord_H = strRecord_H & "|" & ds.Tables(0).Rows(cnt).Item("Total_Qty_Shipped")
                        PrintLine(1, strRecord_H & "|") : intLineNo = intLineNo + 1
                        strRecord_H = ""
                        'Loop for item code
                        For cnt_T = 0 To ds.Tables(1).Rows.Count - 1
                            strRecord_H = ds.Tables(1).Rows(cnt_T).Item("Detail")
                            strRecord_H = strRecord_H & "|" & ds.Tables(1).Rows(cnt_T).Item("Part_Number")
                            strRecord_H = strRecord_H & "|" & ds.Tables(1).Rows(cnt_T).Item("Despatch_Qty")
                            strRecord_H = strRecord_H & "|" & ds.Tables(1).Rows(cnt_T).Item("Measurement_Unit")
                            strRecord_H = strRecord_H & "|" & ds.Tables(1).Rows(cnt_T).Item("Cumulative_Qty")
                            strRecord_H = strRecord_H & "|" & ds.Tables(1).Rows(cnt_T).Item("Purchase_Order_Reference")
                            strRecord_H = strRecord_H & "|" & ds.Tables(1).Rows(cnt_T).Item("Number_of_containers")
                            strRecord_H = strRecord_H & "|" & ds.Tables(1).Rows(cnt_T).Item("number_of_units_per_container")
                            PrintLine(1, strRecord_H & "|") : intLineNo = intLineNo + 1
                            strRecord_H = ""
                            For cnt_T1 = 0 To ds.Tables(2).Rows.Count - 1
                                If ds.Tables(1).Rows(cnt_T).Item("Part_Number") = ds.Tables(2).Rows(cnt_T1).Item("Part_Number") Then
                                    strRecord_H = ds.Tables(2).Rows(cnt_T1).Item("Detail1")
                                    strRecord_H = strRecord_H & "|" & ds.Tables(2).Rows(cnt_T1).Item("Shipping_Label")
                                    PrintLine(1, strRecord_H & "|") : intLineNo = intLineNo + 1
                                    strRecord_H = ""
                                End If

                            Next
                            For cnt_T2 = 0 To ds.Tables(3).Rows.Count - 1
                                If ds.Tables(1).Rows(cnt_T).Item("Part_Number") = ds.Tables(3).Rows(cnt_T2).Item("Part_Number") Then
                                    strRecord_H = ds.Tables(3).Rows(cnt_T2).Item("Detail2")
                                    strRecord_H = strRecord_H & "|" & ds.Tables(3).Rows(cnt_T2).Item("Packaging_Specification_Number")
                                    PrintLine(1, strRecord_H & "|") : intLineNo = intLineNo + 1
                                    strRecord_H = ""
                                End If
                            Next
                        Next

                    Next


                    FileClose(1)
                    ASN_856_MP2 = "TRUE|Invoice Text File Generated Successfully."
                    Exit Function
                Else
                    ASN_856_MP2 = "FALSE|Data is not Complete."
                    FileClose(1)
                    Exit Function
                End If
            Else
                ASN_856_MP2 = "FALSE|Invoice Not Found."
                FileClose(1)
                Exit Function
            End If

            ''
        Catch Ex As Exception
            ASN_856_MP2 = Ex.Message.ToString()
            Exit Function
        Finally
            'Kill(strFileName)
            'FileClose(1)
        End Try
    End Function
    Public Function ASN_DESADV_96A(ByVal UnitCode As String, ByVal strCustomer As String, ByVal strInvoiceNo As String, ByVal strFormatType As String) As String
        Try
            Dim strLogReasons As String = ""
            Dim strLocation As String = ""
            Dim strFileName As String = ""
            Dim intLineNo As Short
            Dim strHdr, strdtl As Object
            Dim strSql As String = ""
            Dim ASNSupplier_Code As String = ""
            Dim ASNCustomer_Code As String = ""
            Dim ASN_SUPPLIER_PLANTCODE As String = ""
            Dim strBarCodeLabel As String = ""
            Dim strRecord_H As String = ""
            Dim dt As DataTable
            Dim dtInvoice As DataTable
            Dim dtASN As DataTable
            Dim dsASN As DataSet
            Dim strInvoiceDate As Date
            Dim strAllowedDate As Date
            strAllowedDate = DateTime.Now.AddDays(-2)
            strInvoiceDate = SqlConnectionclass.ExecuteScalar("Select ent_dt from saleschallan_dtl where doc_no='" & txtInvoice_no.Text & "' and unit_code='" & gstrUNITID & "'")
            If strInvoiceDate > strAllowedDate Then
                MsgBox("MANUAL ASN - ASN can be generated only after two days of Invoice generation.  " & strInvoiceNo & "", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            End If
            'strLocation = Trim(Find_Value("SELECT ISNULL(AUTOASN_PATH  ,'') FROM SALES_PARAMETER WHERE UNIT_CODE='" & UnitCode & "'"))
            ''strLocation = "D:\AUTOASN\"
            'If Len(strLocation) = 0 Then
            '    ASN_DESADV_96A = "FALSE|Default location not defined in sales_parameter."
            '    Exit Function
            'Else
            '    If Mid(Trim(strLocation), Len(Trim(strLocation))) <> "\" Then
            '        strLocation = strLocation & "\"
            '    End If
            '    If Directory.Exists(strLocation) = False Then
            '        Directory.CreateDirectory(strLocation)
            '    End If
            'End If
            Dim ExistBarCode As Boolean = False
            strSql = "Select * FROM AUTOASN_MAPPING WHERE CUSTOMER_CODE='" & strCustomer & "'  and unit_code='" & gstrUNITID & "'"
            dt = SqlConnectionclass.GetDataTable(strSql)
            If dt.Rows.Count > 0 Then
                ASNSupplier_Code = ""
                ASNCustomer_Code = ""
                ASN_SUPPLIER_PLANTCODE = ""
                strBarCodeLabel = ""
                strFormatType = Convert.ToString(dt.Rows(0)("ASNFORMAT_TYPE"))
                ASNSupplier_Code = Convert.ToString(dt.Rows(0)("ASN_SUPPLIERCODE"))
                ASNCustomer_Code = Convert.ToString(dt.Rows(0)("ASN_CUSTOMERCODE"))
                ASN_SUPPLIER_PLANTCODE = Convert.ToString(dt.Rows(0)("ASN_SUPPLIER_PLANTCODE"))
                strBarCodeLabel = Convert.ToString(dt.Rows(0)("CUSTOMER_LABEL"))
                strLocation = Convert.ToString(dt.Rows(0)("AUTOASN_PATH"))
            End If
            If strBarCodeLabel = "E" Then
                Dim strVDAexist = SqlConnectionclass.ExecuteScalar("SELECT INVOICENO FROM VDA_ASN_INVLABELS  where INVOICENO='" & strInvoiceNo & "' and unit_code='" & UnitCode & "'")
                If Len(strVDAexist) = 0 Then
                    ExistBarCode = False
                Else
                    ExistBarCode = True
                End If
            Else
                ExistBarCode = True
            End If

            'Path Code start here
            'strLocation = "D:\AUTOASN\"
            If Len(strLocation) = 0 Then
                ASN_DESADV_96A = "FALSE|Default location not defined in sales_parameter."
                Exit Function
            Else
                If Mid(Trim(strLocation), Len(Trim(strLocation))) <> "\" Then
                    strLocation = strLocation & "\"
                End If
                If Directory.Exists(strLocation) = False Then
                    Directory.CreateDirectory(strLocation)
                End If
            End If
            'Path Code ends here
            If ExistBarCode = True Then
                '---------------LOG DATA ---------------------------' Invoice wise start
                strLogReasons = "[INVOICE START : " & strInvoiceNo & "  FOR CUSTOMER - " & strCustomer & "   ASN FORMAT - " & strFormatType & " ] "
                strSql = "INSERT INTO AUTOASN_LOG (DOC_NO,SLNO,CUSTOMER_CODE,LOG_REASON,PROCESS_TYPE,PROCESS_NAME ,ASN_HEADER_STRING ,LOT_DETAIL_STRING,RESULT,ENT_DT,UNIT_CODE,ENT_USERID) " &
                "SELECT " & strInvoiceNo & " ,(SELECT  ISNULL(MAX(SLNO),0)+1 FROM AUTOASN_LOG) SLNO,'" & strCustomer & "','" & strLogReasons & "','GETMAPPINGDATA','TABLE-AUTOASN_MAPPING' ,'' ,'','SUCCESS', getdate(),'" & gstrUNITID & "','MANASN'"
                SqlConnectionclass.ExecuteNonQuery(strSql)
                '---------------LOG DATA ---------------------------'

                ''location code start here
                strFileName = strLocation & gstrUNITID & "_" & strInvoiceNo & ".txt"
                If System.IO.File.Exists(strFileName) Then
                    Kill(strFileName)
                    FileClose(1)
                End If
                FileOpen(1, strFileName, OpenMode.Append)
                strRecord_H = ""
                dsASN = New DataSet
                ''Location code ends here
                If strFormatType = "DESADV_97A" Then
                    Using sqlcmd1 As SqlCommand = New SqlCommand
                        With sqlcmd1
                            .CommandText = "USP_ASN_DESADEV_96A"
                            .CommandTimeout = 0
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                            .Parameters.Add("@invoice_no", SqlDbType.VarChar, 20).Value = strInvoiceNo
                            .Parameters.Add("@ASNSupplier_Code", SqlDbType.VarChar, 35).Value = ASNSupplier_Code
                            .Parameters.Add("@ASNCustomer_Code", SqlDbType.VarChar, 35).Value = ASNCustomer_Code
                            .Parameters.Add("@ASN_SUPPLIER_PLANTCODE", SqlDbType.VarChar, 35).Value = ASN_SUPPLIER_PLANTCODE
                            .Parameters.Add("@IP_Address", SqlDbType.VarChar, 30).Value = ""
                            .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = "AUTOASN"
                            dsASN = SqlConnectionclass.GetDataSet(sqlcmd1)
                        End With
                    End Using
                    If dsASN.Tables(0).Rows.Count = 0 Then
                        '---------------LOG DATA ---------------------------' ASN Data found
                        strLogReasons = "[INVOICE DATA NOT FOUND FOR : " & strInvoiceNo & "  FOR CUSTOMER - " & strCustomer & "   ASN FORMAT - " & strFormatType & " ] "
                        strSql = "INSERT INTO AUTOASN_LOG (DOC_NO,SLNO,CUSTOMER_CODE,LOG_REASON,PROCESS_TYPE,PROCESS_NAME ,ASN_HEADER_STRING ,LOT_DETAIL_STRING,RESULT,ENT_DT,UNIT_CODE,ENT_USERID) " &
                        "SELECT " & strInvoiceNo & ", (SELECT ISNULL(MAX(SLNO),0)+1 FROM AUTOASN_LOG) SLNO,'" & strCustomer & "','" & strLogReasons & "','GETASNDATA' ,'PROC - USP_ASN_DESADEV_96A','' ,'','SUCCESS', getdate(),'" & gstrUNITID & "','MANASN'"
                        SqlConnectionclass.ExecuteNonQuery(strSql)
                        '---------------LOG DATA ---------------------------'
                        FileClose(1)
                        Exit Function
                    End If
                    Dim strItemCode As String = ""
                    Dim min1, max1 As Integer
                    Dim no_of_Box As String = ""
                    If dsASN.Tables(0).Rows.Count > 0 And dsASN.Tables.Count = 2 Then
                        For cnt As Integer = 0 To dsASN.Tables(0).Rows.Count - 1
                            strRecord_H = dsASN.Tables(0).Rows(cnt).Item("Header")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Supplier_Code")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Customer_Code")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Purpose")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Invoice_ASN_Number")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("ASN_Date")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("ASN_Time")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Shipped_Date")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Shipped_Time")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Shipment_Hierarchy_Level")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("GrossWt_of_Shipment")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("NetWt_of_Shipment")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Carrier_Identification_Code")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Mode_of_Transport")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Equipment_identification_number")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Bill_of_landing")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Packing_List_Number")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Supplier_plant_code")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Customer_plant_code")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Ship_from_location_Code")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Total_Line_Items")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Total_Qty_Shipped")
                            PrintLine(1, strRecord_H) : intLineNo = intLineNo + 1
                            strRecord_H = ""
                            For cnt2 As Integer = 0 To dsASN.Tables(1).Rows.Count - 1
                                If dsASN.Tables(1).Rows(cnt2).Item("Item_Code") = strItemCode Then
                                    min1 = 1 + max1
                                    max1 = max1 + Convert.ToInt32(dsASN.Tables(1).Rows(cnt2).Item("Barcode_Serial_No").ToString.Trim)
                                    no_of_Box = Convert.ToString(min1) + "-" + Convert.ToString(max1)
                                Else
                                    min1 = 0
                                    max1 = 0
                                    no_of_Box = ""
                                    max1 = Convert.ToInt32(dsASN.Tables(1).Rows(cnt2).Item("Barcode_Serial_No").ToString.Trim)
                                    no_of_Box = Convert.ToString("1") + "-" + Convert.ToString(max1)
                                End If
                                strItemCode = dsASN.Tables(1).Rows(cnt2).Item("Item_Code")
                                strRecord_H = dsASN.Tables(1).Rows(cnt2).Item("Detail")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Pallet_No")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Pallet_Qty")
                                PrintLine(1, strRecord_H) : intLineNo = intLineNo + 1

                                strRecord_H = dsASN.Tables(1).Rows(cnt2).Item("Detail1")
                                strRecord_H = strRecord_H & "," & no_of_Box
                                strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Part_Number")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Despatch_Qty")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Measurement_Unit")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Cumulative_Qty")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Purchase_Order_Reference")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Number_of_containers")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("number_of_units_per_container")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Packing_Code")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Sale_Order_Reference")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Batch_No")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Box_Desc")

                                PrintLine(1, strRecord_H) : intLineNo = intLineNo + 1
                                strRecord_H = ""

                                strRecord_H = dsASN.Tables(0).Rows(cnt).Item("Detail2")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("container_number")
                                PrintLine(1, strRecord_H) : intLineNo = intLineNo + 1

                                strRecord_H = ""

                            Next


                        Next
                        strSql = "Insert into AUTOASN_DETAILS (DOC_NO,CUSTOMER_CODE,INVOICE_DATE,ENT_DT,UNIT_CODE,ASNFORMAT_YPE,ENT_USERID)" &
                        "values(" & strInvoiceNo & ",'" & strCustomer & "',Convert(DateTime,'" & strInvoiceDate & "', 103),getdate(),'" & gstrUNITID & "','" & strFormatType & "','MANASN')"
                        SqlConnectionclass.ExecuteNonQuery(strSql)


                        '---------------LOG DATA ---------------------------'
                        strLogReasons = "[INVOICE COMPLETED : " & strInvoiceNo & "  FOR CUSTOMER - " & strCustomer & "   ASN FORMAT - " & strFormatType & " ] "
                        strSql = "INSERT INTO AUTOASN_LOG (DOC_NO,SLNO,CUSTOMER_CODE,LOG_REASON,PROCESS_TYPE,PROCESS_NAME ,ASN_HEADER_STRING ,LOT_DETAIL_STRING,RESULT,ENT_DT,UNIT_CODE,ENT_USERID) " &
                        "SELECT " & strInvoiceNo & " ,(SELECT  ISNULL(MAX(SLNO),0)+1 FROM AUTOASN_LOG) SLNO,'" & strCustomer & "','" & strLogReasons & "','GETINVOICEDATA','USP_GET_PENDING_INVOICES_AUTOASN' ,'' ,'','SUCCESS', getdate(),'" & gstrUNITID & "','MANASN'"
                        SqlConnectionclass.ExecuteNonQuery(strSql)
                        '---------------LOG DATA ---------------------------'
                    End If

                ElseIf strFormatType = "DESADV_96A" Then
                    Using sqlcmd1 As SqlCommand = New SqlCommand
                        With sqlcmd1
                            .CommandText = "USP_ASN_DESADEV_PKC"
                            .CommandTimeout = 0
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = UnitCode
                            .Parameters.Add("@DOC_NO", SqlDbType.BigInt).Value = strInvoiceNo
                            .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = "AUTOASN"
                            SqlConnectionclass.ExecuteNonQuery(sqlcmd1)

                            strSql = "SELECT  FormatType,Doc_No,InvoiceDate,ItemCode,Cust_Item_Code,Cust_Item_Desc,Gross_Weight,Net_Weight,Sales_Qty, " &
               "Cust_vendor_code,ReferenceNo, No_of_Box,BoxNo_Identity, ISNULL( Pkg_style_c,'BOX') as Pkg_style_c, Box_Qty, ItemWeight, PalletWeight, boxweight " &
               " FROM TEMP_DESADEV_96A WHERE UNIT_CODE='" & UnitCode & "' and Doc_No='" & strInvoiceNo & "'"
                            dtASN = SqlConnectionclass.GetDataTable(strSql)
                            If dtASN.Rows.Count > 0 Then

                                For Each drASN As DataRow In dtASN.Rows
                                    strHdr = "HDR"
                                    strHdr = strHdr & "|" & drASN("Doc_No")
                                    strHdr = strHdr & "|" & drASN("InvoiceDate")
                                    strHdr = strHdr & "|" & drASN("Gross_Weight")
                                    strHdr = strHdr & "|" & drASN("Net_Weight")
                                    strHdr = strHdr & "|" & drASN("Sales_Qty")
                                    strHdr = strHdr & "|" & drASN("ReferenceNo")
                                    strHdr = strHdr & "|" & drASN("Cust_vendor_code")
                                    strHdr = strHdr & "|" & drASN("Doc_No")
                                    PrintLine(1, strHdr & "|") : intLineNo = intLineNo + 1


                                    strdtl = "DTL"
                                    strdtl = strdtl & "|" & drASN("No_of_Box")
                                    strdtl = strdtl & "|" & drASN("BoxNo_Identity")
                                    strdtl = strdtl & "|" & drASN("Pkg_style_c")
                                    strdtl = strdtl & "|" & drASN("Cust_Item_Code")
                                    strdtl = strdtl & "|" & drASN("ItemCode")
                                    strdtl = strdtl & "|" & drASN("Sales_Qty")
                                    strdtl = strdtl & "|" & drASN("ReferenceNo")
                                    PrintLine(1, strdtl) : intLineNo = intLineNo + 1

                                    strHdr = ""
                                    strdtl = ""
                                Next
                                strSql = "Insert into AUTOASN_DETAILS (DOC_NO,CUSTOMER_CODE,ENT_DT,UNIT_CODE,ASNFORMAT_YPE,ENT_USERID)" &
                                        "values(" & strInvoiceNo & ",'" & txtCustomerCode.Text & "',getdate(),'" & UnitCode & "','DESADV_96A','MANASN')"
                                SqlConnectionclass.ExecuteNonQuery(strSql)
                            Else
                                MsgBox("MANUAL ASN - ASN DATA NOT FOUND FOR INVOICE " & strInvoiceNo & "", MsgBoxStyle.Information, ResolveResString(100))
                            End If

                        End With
                    End Using
                ElseIf strFormatType = "DUMAREY_97A" Then '' Added ny priti on 21 Jan 2026 for new customer Dumrey
                    Using sqlcmd1 As SqlCommand = New SqlCommand
                            With sqlcmd1
                            .CommandText = "USP_ASN_DUMAREY_97A"
                            .CommandTimeout = 0
                                .CommandType = CommandType.StoredProcedure
                                .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                                .Parameters.Add("@invoice_no", SqlDbType.VarChar, 20).Value = strInvoiceNo
                                .Parameters.Add("@ASNSupplier_Code", SqlDbType.VarChar, 35).Value = ASNSupplier_Code
                                .Parameters.Add("@ASNCustomer_Code", SqlDbType.VarChar, 35).Value = ASNCustomer_Code
                                .Parameters.Add("@ASN_SUPPLIER_PLANTCODE", SqlDbType.VarChar, 35).Value = ASN_SUPPLIER_PLANTCODE
                                .Parameters.Add("@IP_Address", SqlDbType.VarChar, 30).Value = ""
                                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = "AUTOASN"
                                dsASN = SqlConnectionclass.GetDataSet(sqlcmd1)
                            End With
                        End Using
                        If dsASN.Tables(0).Rows.Count = 0 Then
                            '---------------LOG DATA ---------------------------' ASN Data found
                            strLogReasons = "[INVOICE DATA NOT FOUND FOR : " & strInvoiceNo & "  FOR CUSTOMER - " & strCustomer & "   ASN FORMAT - " & strFormatType & " ] "
                            strSql = "INSERT INTO AUTOASN_LOG (DOC_NO,SLNO,CUSTOMER_CODE,LOG_REASON,PROCESS_TYPE,PROCESS_NAME ,ASN_HEADER_STRING ,LOT_DETAIL_STRING,RESULT,ENT_DT,UNIT_CODE,ENT_USERID) " &
                            "SELECT " & strInvoiceNo & ", (SELECT ISNULL(MAX(SLNO),0)+1 FROM AUTOASN_LOG) SLNO,'" & strCustomer & "','" & strLogReasons & "','GETASNDATA' ,'PROC - USP_ASN_DESADEV_96A','' ,'','SUCCESS', getdate(),'" & gstrUNITID & "','MANASN'"
                            SqlConnectionclass.ExecuteNonQuery(strSql)
                            '---------------LOG DATA ---------------------------'
                            FileClose(1)
                            Exit Function
                        End If
                        Dim strItemCode As String = ""
                        Dim min1, max1 As Integer
                        Dim no_of_Box As String = ""
                        If dsASN.Tables(0).Rows.Count > 0 And dsASN.Tables.Count = 2 Then
                            For cnt As Integer = 0 To dsASN.Tables(0).Rows.Count - 1
                                strRecord_H = dsASN.Tables(0).Rows(cnt).Item("Header")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Supplier_Code")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Customer_Code")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Purpose")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Invoice_ASN_Number")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("ASN_Date")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("ASN_Time")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Shipped_Date")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Shipped_Time")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Shipment_Hierarchy_Level")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("GrossWt_of_Shipment")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("NetWt_of_Shipment")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Carrier_Identification_Code")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Mode_of_Transport")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Equipment_identification_number")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Bill_of_landing")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Packing_List_Number")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Supplier_plant_code")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Customer_plant_code")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Ship_from_location_Code")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Total_Line_Items")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Total_Qty_Shipped")
                                PrintLine(1, strRecord_H) : intLineNo = intLineNo + 1
                                strRecord_H = ""
                                For cnt2 As Integer = 0 To dsASN.Tables(1).Rows.Count - 1
                                    If dsASN.Tables(1).Rows(cnt2).Item("Item_Code") = strItemCode Then
                                        min1 = 1 + max1
                                        max1 = max1 + Convert.ToInt32(dsASN.Tables(1).Rows(cnt2).Item("Barcode_Serial_No").ToString.Trim)
                                        no_of_Box = Convert.ToString(min1) + "-" + Convert.ToString(max1)
                                    Else
                                        min1 = 0
                                        max1 = 0
                                        no_of_Box = ""
                                        max1 = Convert.ToInt32(dsASN.Tables(1).Rows(cnt2).Item("Barcode_Serial_No").ToString.Trim)
                                        no_of_Box = Convert.ToString("1") + "-" + Convert.ToString(max1)
                                    End If
                                    strItemCode = dsASN.Tables(1).Rows(cnt2).Item("Item_Code")
                                    strRecord_H = dsASN.Tables(1).Rows(cnt2).Item("Detail")
                                    strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Pallet_No")
                                    strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Pallet_Qty")
                                    PrintLine(1, strRecord_H) : intLineNo = intLineNo + 1

                                    strRecord_H = dsASN.Tables(1).Rows(cnt2).Item("Detail1")
                                    strRecord_H = strRecord_H & "," & no_of_Box
                                    strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Part_Number")
                                    strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Despatch_Qty")
                                    strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Measurement_Unit")
                                    strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Cumulative_Qty")
                                    strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Purchase_Order_Reference")
                                    strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Number_of_containers")
                                    strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("number_of_units_per_container")
                                    strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Packing_Code")
                                    strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Sale_Order_Reference")
                                    strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Batch_No")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(1).Rows(cnt2).Item("Box_Desc")
                                strRecord_H = strRecord_H & "," & strItemCode

                                PrintLine(1, strRecord_H) : intLineNo = intLineNo + 1
                                    strRecord_H = ""

                                    strRecord_H = dsASN.Tables(0).Rows(cnt).Item("Detail2")
                                    strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("container_number")
                                    PrintLine(1, strRecord_H) : intLineNo = intLineNo + 1

                                    strRecord_H = ""

                                Next


                            Next
                            strSql = "Insert into AUTOASN_DETAILS (DOC_NO,CUSTOMER_CODE,INVOICE_DATE,ENT_DT,UNIT_CODE,ASNFORMAT_YPE,ENT_USERID)" &
                            "values(" & strInvoiceNo & ",'" & strCustomer & "',Convert(DateTime,'" & strInvoiceDate & "', 103),getdate(),'" & gstrUNITID & "','" & strFormatType & "','MANASN')"
                            SqlConnectionclass.ExecuteNonQuery(strSql)


                            '---------------LOG DATA ---------------------------'
                            strLogReasons = "[INVOICE COMPLETED : " & strInvoiceNo & "  FOR CUSTOMER - " & strCustomer & "   ASN FORMAT - " & strFormatType & " ] "
                            strSql = "INSERT INTO AUTOASN_LOG (DOC_NO,SLNO,CUSTOMER_CODE,LOG_REASON,PROCESS_TYPE,PROCESS_NAME ,ASN_HEADER_STRING ,LOT_DETAIL_STRING,RESULT,ENT_DT,UNIT_CODE,ENT_USERID) " &
                            "SELECT " & strInvoiceNo & " ,(SELECT  ISNULL(MAX(SLNO),0)+1 FROM AUTOASN_LOG) SLNO,'" & strCustomer & "','" & strLogReasons & "','GETINVOICEDATA','USP_GET_PENDING_INVOICES_AUTOASN' ,'' ,'','SUCCESS', getdate(),'" & gstrUNITID & "','MANASN'"
                            SqlConnectionclass.ExecuteNonQuery(strSql)
                            '---------------LOG DATA ---------------------------'
                        End If
                    ElseIf strFormatType = "ANSIX12856" Then
                    Using sqlcmd1 As SqlCommand = New SqlCommand
                        With sqlcmd1
                            .CommandText = "USP_ASN_ANSIX12856Lear"
                            .CommandTimeout = 0
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                            .Parameters.Add("@DOC_NO", SqlDbType.VarChar, 20).Value = strInvoiceNo
                            .Parameters.Add("@ASNSupplier_Code", SqlDbType.VarChar, 35).Value = ASNSupplier_Code
                            .Parameters.Add("@ASNCustomerEDI_Code", SqlDbType.VarChar, 35).Value = ASNCustomer_Code
                            .Parameters.Add("@ASN_SUPPLIER_PLANTCODE", SqlDbType.VarChar, 35).Value = ASN_SUPPLIER_PLANTCODE
                            .Parameters.Add("@IP_Address", SqlDbType.VarChar, 30).Value = ""
                            .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = "AUTOASN"
                            dsASN = SqlConnectionclass.GetDataSet(sqlcmd1)
                        End With
                    End Using
                    If dsASN.Tables(0).Rows.Count = 0 Then
                        '---------------LOG DATA ---------------------------' ASN Data found
                        strLogReasons = "[INVOICE DATA NOT FOUND FOR : " & strInvoiceNo & "  FOR CUSTOMER - " & strCustomer & "   ASN FORMAT - " & strFormatType & " ] "
                        strSql = "INSERT INTO AUTOASN_LOG (DOC_NO,SLNO,CUSTOMER_CODE,LOG_REASON,PROCESS_TYPE,PROCESS_NAME ,ASN_HEADER_STRING ,LOT_DETAIL_STRING,RESULT,ENT_DT,UNIT_CODE,ENT_USERID) " &
                        "SELECT " & strInvoiceNo & ", (SELECT ISNULL(MAX(SLNO),0)+1 FROM AUTOASN_LOG) SLNO,'" & strCustomer & "','" & strLogReasons & "','GETASNDATA' ,'PROC - USP_ASN_ANSIX12856Lear','' ,'','SUCCESS', getdate(),'" & gstrUNITID & "','MANASN'"
                        SqlConnectionclass.ExecuteNonQuery(strSql)
                        '---------------LOG DATA ---------------------------'
                        FileClose(1)
                        Exit Function
                    End If
                    If dsASN.Tables(0).Rows.Count > 0 Then
                        For cnt As Integer = 0 To dsASN.Tables(0).Rows.Count - 1
                            If cnt = 0 Then
                                strRecord_H = "HDR"
                                strRecord_H = strRecord_H & "," & ASNSupplier_Code
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Customer_EDICode")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Transaction_PurposeCode")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Doc_No")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("ASN_Date")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("ASN_Time")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Shipped_Date")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Shipped_Time")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("TotalGross_Weight")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("TotalNet_Weight")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("CarrirerSec_Code")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Mode_Of_Shipment")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("ConveyanceNo")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Bill_of_landing")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Packing_List_Number")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Supplier_plant_code")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Customer_plant_code")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Total_Line_Items")
                                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Total_Qty_Shipped")

                                PrintLine(1, strRecord_H) : intLineNo = intLineNo + 1
                                strRecord_H = ""
                            End If

                            strRecord_H = dsASN.Tables(0).Rows(cnt).Item("Detail")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Cust_Item_Code")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Sales_Qty")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("UOM")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("CumsQty")
                            strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(cnt).Item("Purchase_Order_Reference")

                            PrintLine(1, strRecord_H) : intLineNo = intLineNo + 1
                            strRecord_H = ""
                        Next
                        strSql = "Insert into AUTOASN_DETAILS (DOC_NO,CUSTOMER_CODE,INVOICE_DATE,ENT_DT,UNIT_CODE,ASNFORMAT_YPE,ENT_USERID)" &
                       "values(" & strInvoiceNo & ",'" & strCustomer & "',Convert(DateTime,'" & strInvoiceDate & "', 103),getdate(),'" & gstrUNITID & "','" & strFormatType & "','MANASN')"
                        SqlConnectionclass.ExecuteNonQuery(strSql)


                        '---------------LOG DATA ---------------------------'
                        strLogReasons = "[INVOICE COMPLETED : " & strInvoiceNo & "  FOR CUSTOMER - " & strCustomer & "   ASN FORMAT - " & strFormatType & " ] "
                        strSql = "INSERT INTO AUTOASN_LOG (DOC_NO,SLNO,CUSTOMER_CODE,LOG_REASON,PROCESS_TYPE,PROCESS_NAME ,ASN_HEADER_STRING ,LOT_DETAIL_STRING,RESULT,ENT_DT,UNIT_CODE,ENT_USERID) " &
                        "SELECT " & strInvoiceNo & " ,(SELECT  ISNULL(MAX(SLNO),0)+1 FROM AUTOASN_LOG) SLNO,'" & strCustomer & "','" & strLogReasons & "','GETINVOICEDATA','USP_GET_PENDING_INVOICES_AUTOASN' ,'' ,'','SUCCESS', getdate(),'" & gstrUNITID & "','MANASN'"
                        SqlConnectionclass.ExecuteNonQuery(strSql)
                        '---------------LOG DATA ---------------------------'
                    Else
                        MsgBox("MANUAL ASN - ASN DATA NOT FOUND FOR INVOICE " & strInvoiceNo & "", MsgBoxStyle.Information, ResolveResString(100))
                    End If

                End If
            Else
                MsgBox("MANUAL ASN - VDA BarCode Data not found", MsgBoxStyle.Information, ResolveResString(100))
            End If
            FileClose(1)
            ASN_DESADV_96A = "TRUE|Invoice Text File Generated Successfully."
            Exit Function
        Catch ex As Exception
            ASN_DESADV_96A = ex.Message.ToString()
            Exit Function
        End Try
    End Function

    Public Function ASN_Bosch_MS1(ByVal UnitCode As String, ByVal strCustomer As String, ByVal strInvoiceNo As String, ByVal strFormatType As String) As String
        Try
            Dim strLogReasons As String = ""
            Dim strLocation As String = ""
            Dim strFileName As String = ""
            Dim intLineNo As Short

            Dim strSql As String = ""
            Dim ASNSupplier_Code As String = ""
            Dim ASNCustomer_Code As String = ""
            Dim ASN_SUPPLIER_PLANTCODE As String = ""
            Dim strBarCodeLabel As String = ""
            Dim strRecord_H As String = ""
            Dim dt As DataTable
            Dim dtInvoice As DataTable
            Dim dtASN As DataTable
            Dim dsASN As DataSet
            Dim strInvoiceDate As Date
            'Dim strAllowedDate As Date
            'strAllowedDate = DateTime.Now.AddDays(-2)
            'strInvoiceDate = SqlConnectionclass.ExecuteScalar("Select ent_dt from saleschallan_dtl where doc_no='" & txtInvoice_no.Text & "' and unit_code='" & gstrUNITID & "'")
            'If strInvoiceDate > strAllowedDate Then
            '    MsgBox("MANUAL ASN - ASN can be generated only after two days of Invoice generation.  " & strInvoiceNo & "", MsgBoxStyle.Information, ResolveResString(100))
            '    Exit Function
            'End If
            ''location code start here
            strLocation = Trim(Find_Value("SELECT ISNULL(BoschASNLoc  ,'') FROM SALES_PARAMETER WHERE UNIT_CODE='" & UnitCode & "'"))
            'strLocation = "D:\BoschASN\"
            If Len(strLocation) = 0 Then
                ASN_Bosch_MS1 = "FALSE|Default location not defined in sales_parameter."
                Exit Function
            Else
                If Mid(Trim(strLocation), Len(Trim(strLocation))) <> "\" Then
                    strLocation = strLocation & "\"
                End If
                If Directory.Exists(strLocation) = False Then
                    Directory.CreateDirectory(strLocation)
                End If
            End If


            strFileName = strLocation & strInvoiceNo & ".txt"
            If System.IO.File.Exists(strFileName) Then
                Kill(strFileName)
                FileClose(1)
            End If
            FileOpen(1, strFileName, OpenMode.Append)
            strRecord_H = ""
            dsASN = New DataSet
            ''Location code ends here
            Using sqlcmd1 As SqlCommand = New SqlCommand
                With sqlcmd1
                    .CommandText = "Proc_ASN_Data_Bosch_VV"
                    .CommandTimeout = 0
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@invoice_no", SqlDbType.VarChar, 20).Value = strInvoiceNo
                    '.Parameters.Add("@ASNSupplier_Code", SqlDbType.VarChar, 35).Value = ASNSupplier_Code
                    '.Parameters.Add("@ASNCustomer_Code", SqlDbType.VarChar, 35).Value = ASNCustomer_Code
                    '.Parameters.Add("@ASN_SUPPLIER_PLANTCODE", SqlDbType.VarChar, 35).Value = ASN_SUPPLIER_PLANTCODE
                    '.Parameters.Add("@IP_Address", SqlDbType.VarChar, 30).Value = ""
                    '.Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = "AUTOASN"
                    dsASN = SqlConnectionclass.GetDataSet(sqlcmd1)
                End With
            End Using
            If dsASN.Tables(0).Rows.Count = 0 Then
                '---------------LOG DATA ---------------------------' ASN Data found
                strLogReasons = "[INVOICE DATA NOT FOUND FOR : " & strInvoiceNo & "  FOR CUSTOMER - " & strCustomer & "   ASN FORMAT - " & strFormatType & " ] "
                strSql = "INSERT INTO AUTOASN_LOG (DOC_NO,SLNO,CUSTOMER_CODE,LOG_REASON,PROCESS_TYPE,PROCESS_NAME ,ASN_HEADER_STRING ,LOT_DETAIL_STRING,RESULT,ENT_DT,UNIT_CODE,ENT_USERID) " &
                "SELECT " & strInvoiceNo & ", (SELECT ISNULL(MAX(SLNO),0)+1 FROM AUTOASN_LOG) SLNO,'" & strCustomer & "','" & strLogReasons & "','GETBoschASNDATA' ,'PROC - Proc_ASN_Data_Bosch_VV','' ,'','SUCCESS', getdate(),'" & gstrUNITID & "','MANASN'"
                SqlConnectionclass.ExecuteNonQuery(strSql)
                '---------------LOG DATA ---------------------------'
                FileClose(1)
                Exit Function
            End If
            Dim strItemCode As String = ""
            Dim strPallet As String = ""
            Dim no_of_Box As String = ""
            If dsASN.Tables(0).Rows.Count > 0 And dsASN.Tables.Count = 5 Then

                strRecord_H = "HDR"
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("MSSLSenderID")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("BoschReceiverID")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("ProductionCode")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("ASNNo")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("ASNDate")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("TransportContarctIdentifier")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("TransportContarctRefNo")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("DocumnetDt")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("TransportDt")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("ShipmentDt")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("Gross_Weight")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("Net_Weight")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("TotalInvoiceQty")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("BuyerID")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("SellerID")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("ShipFromPartyID")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("ShiptoPartyID")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("BuyerAdd1")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("BuyerAdd2")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("BuyerAdd3")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("BuyerAdd4")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("BuyerAdd5")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("BuyerAdd6")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("SellerAdd1")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("SellerAdd2")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("SellerAdd3")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("SellerAdd4")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("SellerAdd5")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("SellerAdd6")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("plantCode")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("Port_Of_Discharge")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("Supplier_code")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("ShipFrom")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("EquipmentType")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("EquipmentIdentification")
                strRecord_H = strRecord_H & "," & dsASN.Tables(0).Rows(0).Item("additionalDest")
                PrintLine(1, strRecord_H) : intLineNo = intLineNo + 1
                strRecord_H = ""
                Dim dtDistinct As DataTable = dsASN.Tables(1).DefaultView.ToTable(True, "Item_Code")
                Dim Itemindex As Integer = 0
                Dim palletNoArray() As String =
                 dtDistinct.AsEnumerable().
                 Select(Function(r) r.Field(Of String)("Item_Code")).
                 Distinct().
                 ToArray() 'add by tek
                For Each ItemCode As String In palletNoArray
                    For Each cnt2 As DataRow In dsASN.Tables(1).Select("Type='2' and item_code='" & palletNoArray(Itemindex) & "'")
                        strPallet = cnt2("palletno")

                        strRecord_H = "DTL1"
                        strRecord_H = strRecord_H & "," & "1"
                        strRecord_H = strRecord_H & "," & cnt2("PalletID")
                        strRecord_H = strRecord_H & "," & cnt2("NoOfBoxes")
                        strRecord_H = strRecord_H & "," & cnt2("QtyPerBox")
                        strRecord_H = strRecord_H & "," & cnt2("UOM")
                        strRecord_H = strRecord_H & "," & cnt2("PackType")
                        strRecord_H = strRecord_H & "," & cnt2("AddPartInfo")
                        strRecord_H = strRecord_H & "," & cnt2("GrossMeasure_Inner")
                        strRecord_H = strRecord_H & "," & cnt2("gross_Wt_Inner")
                        strRecord_H = strRecord_H & "," & cnt2("NetWt_Inner")
                        strRecord_H = strRecord_H & "," & cnt2("LenMeasure_Inner")
                        strRecord_H = strRecord_H & "," & cnt2("PalLen_Inner")
                        strRecord_H = strRecord_H & "," & cnt2("PalWid_Inner")
                        strRecord_H = strRecord_H & "," & cnt2("PalHgt_Inner")
                        strRecord_H = strRecord_H & "," & cnt2("StackAbility")
                        PrintLine(1, strRecord_H) : intLineNo = intLineNo + 1
                        strRecord_H = ""
                        'End If
                        For Each cntBox As DataRow In dsASN.Tables(1).Select("Type='1' and palletno='" & strPallet & "' and item_code='" & palletNoArray(Itemindex) & "' ")

                            strRecord_H = "DTL2"
                            strRecord_H = strRecord_H & "," & cntBox("PalletID")
                            PrintLine(1, strRecord_H) : intLineNo = intLineNo + 1
                            strRecord_H = ""

                        Next

                    Next
                    For Each cntSummary As DataRow In dsASN.Tables(2).Select("item_code='" & palletNoArray(Itemindex) & "' ") '"palletno='" & strPallet & "' and
                        strRecord_H = "DTL3"
                        strRecord_H = strRecord_H & "," & cntSummary("BuyerPartNo")
                        strRecord_H = strRecord_H & "," & cntSummary("DispatchQty")
                        strRecord_H = strRecord_H & "," & cntSummary("OrderNo")
                        strRecord_H = strRecord_H & "," & cntSummary("DischargePort")
                        strRecord_H = strRecord_H & "," & cntSummary("Item_Code")
                        strRecord_H = strRecord_H & "," & cntSummary("REVISION_NO")
                        strRecord_H = strRecord_H & "," & cntSummary("BatchNo_NO")
                        strRecord_H = strRecord_H & "," & cntSummary("ManFPart_NO")
                        strRecord_H = strRecord_H & "," & cntSummary("ROHS")
                        strRecord_H = strRecord_H & "," & cntSummary("ExpiryDt")
                        strRecord_H = strRecord_H & "," & cntSummary("ProductionDt")
                        strRecord_H = strRecord_H & "," & cntSummary("DUNSNo")
                        PrintLine(1, strRecord_H) : intLineNo = intLineNo + 1
                        strRecord_H = ""

                    Next
                    Itemindex = Itemindex + 1
                Next

            End If
            If dsASN.Tables(3).Rows.Count > 0 Then
                strRecord_H = "DTL4"
                strRecord_H = strRecord_H & "," & dsASN.Tables(3).Rows(0).Item("ConSqNo")
                strRecord_H = strRecord_H & "," & dsASN.Tables(3).Rows(0).Item("SupArtNo")
                strRecord_H = strRecord_H & "," & dsASN.Tables(3).Rows(0).Item("TotalPallets")
                strRecord_H = strRecord_H & "," & dsASN.Tables(3).Rows(0).Item("TotalBoxes")
                strRecord_H = strRecord_H & "," & dsASN.Tables(3).Rows(0).Item("GrossWTMeasure_Outer")
                strRecord_H = strRecord_H & "," & dsASN.Tables(3).Rows(0).Item("GrossWt_Outer")
                strRecord_H = strRecord_H & "," & dsASN.Tables(3).Rows(0).Item("NetWt_Outer")
                strRecord_H = strRecord_H & "," & dsASN.Tables(3).Rows(0).Item("LenMeasure_Outer")
                strRecord_H = strRecord_H & "," & dsASN.Tables(3).Rows(0).Item("LenVal_Outer")
                strRecord_H = strRecord_H & "," & dsASN.Tables(3).Rows(0).Item("WidthDem_Outer")
                strRecord_H = strRecord_H & "," & dsASN.Tables(3).Rows(0).Item("HeightVal_Outer")
                strRecord_H = strRecord_H & "," & dsASN.Tables(3).Rows(0).Item("MaxStack")
                PrintLine(1, strRecord_H) : intLineNo = intLineNo + 1
                strRecord_H = ""
            End If
            For Each cntPB As DataRow In dsASN.Tables(4).Rows
                strRecord_H = "DTL5"
                strRecord_H = strRecord_H & "," & cntPB("PalletID")
                strRecord_H = strRecord_H & "," & cntPB("PalletBoxID")

                PrintLine(1, strRecord_H) : intLineNo = intLineNo + 1
                strRecord_H = ""
            Next
            FileClose(1)
            ASN_Bosch_MS1 = "TRUE|Invoice Text File Generated Successfully."
            Exit Function

        Catch Ex As Exception
            ASN_Bosch_MS1 = Ex.Message.ToString()
            Exit Function
        Finally
            'Kill(strFileName)
            'FileClose(1)
        End Try
    End Function
    Public Function ASN_4913_MP2(ByVal UnitCode As String, ByVal pstrAccountCode As String, ByVal pstrInvoice As String, ByVal pstrSRV As String) As String

        Try
            Dim strLocation As String
            Dim strFileName As String
            Dim intLineCount As Short
            Dim intLineWidth As Short
            Dim intLineNo As Short
            Dim strSql As String
            Dim strsql2 As String

            Dim strRecord_H As String
            Dim strRecord_P As String
            Dim strRecord_D As String
            Dim cnt As Integer = 0
            Dim cnt_T As Integer = 0

            Dim oSqlCmdLG As New SqlCommand
            Dim adap As SqlDataAdapter
            Dim ds As DataSet

            Dim ds_T As DataSet
            Dim adap_T As SqlDataAdapter

            Dim Item_Code As String = ""
            Dim IP_Address As String = ""

            Dim strSql_Pallet As String = ""
            Dim ds_Pallet As DataSet
            Dim adap_Pallet As SqlDataAdapter
            Dim cnt_Pallet As Integer

            Dim strSql_BOX As String = "'"
            Dim adap_BOX As SqlDataAdapter
            Dim ds_BOX As DataSet
            Dim cnt_box As Integer

            Dim min1, max1, no_of_package As Integer
            min1 = 0
            max1 = 0
            no_of_package = 0

            Dim cnt_pallet_no As Integer


            Dim sqlcmd_T As New SqlCommand
            Dim sqlcmd_ds As New SqlCommand
            Dim sqlcmd_Pallet As New SqlCommand
            Dim sqlcmd_BOX As New SqlCommand

            IP_Address = gstrIpaddressWinSck
            ' IP_Address = "1.01.12.13"
            strLocation = Trim(Find_Value("SELECT ISNULL(TextFileDefaultLocation,'') FROM SALES_PARAMETER_ASN_MP2 WHERE UNIT_CODE='" & UnitCode & "'"))

            If Len(strLocation) = 0 Then
                ASN_4913_MP2 = "FALSE|Default location not defined in sales_parameter."
                Exit Function
            Else
                If Mid(Trim(strLocation), Len(Trim(strLocation))) <> "\" Then
                    strLocation = strLocation & "\"
                End If
                If Directory.Exists(strLocation) = False Then
                    Directory.CreateDirectory(strLocation)
                End If

                'strFileName = strLocation & pstrInvoice & "_" & VB6.Format(Now, "dd-MMM-yyyy") & ".txt"
                strFileName = strLocation & UnitCode & "_" & pstrInvoice & ".txt"
                'On Error Resume Next
                'Kill(strLocation & "*.txt")
                'If System.IO.File.Exists(strFileName) Then
                '    System.IO.File.Delete(strFileName)
                If System.IO.File.Exists(strFileName) Then
                    Kill(strFileName)
                    FileClose(1)
                End If

                'On Error GoTo ErrHandler
                FileOpen(1, strFileName, OpenMode.Append)

            End If
            intLineCount = 60
            intLineWidth = 180
            Dim rsTextfileNonDI As ADODB.Recordset
            Dim strSalesTextFile As String

            ''pstrInvoice = Mid(Trim(pstrInvoice), 2, Len(Trim(pstrInvoice)) - 2)
            'If Len(pstrInvoice) > 0 Then

            strSql = ""
            strSql = "SELECT sales_dtl.Item_Code FROM sales_dtl INNER JOIN SalesChallan_Dtl ON sales_dtl.UNIT_CODE = "
            strSql = strSql & "SalesChallan_Dtl.UNIT_CODE AND sales_dtl.Location_Code = SalesChallan_Dtl.Location_Code "
            strSql = strSql & "AND    sales_dtl.Doc_No = SalesChallan_Dtl.Doc_No WHERE (sales_dtl.Doc_No = '" & pstrInvoice & "')"
            strSql = strSql & "AND (sales_dtl.UNIT_CODE = '" & UnitCode & "') AND (saleschallan_dtl.Account_Code='" & pstrAccountCode & "')"

            'adap_T = New SqlDataAdapter(strSql, SqlConnectionclass.GetConnection)
            '    ds_T = New DataSet
            '    adap_T.Fill(ds_T)

            ds_T = New DataSet
            ' Dim sqlcmd_T As New SqlCommand
            With sqlcmd_T
                .CommandText = strSql
                .CommandType = CommandType.Text
            End With
            ds_T = SqlConnectionclass.GetDataSet(sqlcmd_T)


            If ds_T.Tables(0).Rows.Count >= 1 Then
                'If ds_T.Tables(0).Rows.Count > 1 Then
                'HEADER'

                ''''
                '/////// Pallete sequence update

                Dim STR_PALLET_DELETE As String
                Dim SQLCMD_PALLET_DELETE As New SqlCommand

                STR_PALLET_DELETE = "DELETE FROM MULTIPLE_PALLET_STATUS WHERE UNIT_CODE= '" & UnitCode & "' AND DOC_NO='" & pstrInvoice & "' "

                With SQLCMD_PALLET_DELETE
                    .CommandText = STR_PALLET_DELETE
                    .CommandType = CommandType.Text
                End With
                SqlConnectionclass.ExecuteNonQuery(SQLCMD_PALLET_DELETE)


                Dim STR_PALLET_INSERT As String
                Dim SQLCMD_PALLET_INSERT As New SqlCommand


                STR_PALLET_INSERT = "INSERT INTO MULTIPLE_PALLET_STATUS(UNIT_CODE,DOC_NO,PALLET_NO,UPDATED_STATUS) " &
                " SELECT DISTINCT UNIT_CODE,INVOICENO,PALLETNO,0 FROM VDA_ASN_INVLABELS WHERE INVOICENO='" & pstrInvoice & "' AND UNIT_CODE='" & UnitCode & "'"

                With SQLCMD_PALLET_INSERT
                    .CommandText = STR_PALLET_INSERT
                    .CommandType = CommandType.Text
                End With
                SqlConnectionclass.ExecuteNonQuery(SQLCMD_PALLET_INSERT)

                '/////// Pallete sequence update
                ''''

                For cnt_T = 0 To ds_T.Tables(0).Rows.Count - 1
                    Item_Code = ds_T.Tables(0).Rows(cnt_T).Item("Item_Code")
                    With oSqlCmdLG
                        .Parameters.Clear()
                        .CommandType = CommandType.StoredProcedure
                        .CommandTimeout = 0
                        .CommandText = "USP_ASN_TEXT_M_01"
                        .Parameters.AddWithValue("@UNIT_CODE", UnitCode.Trim)
                        .Parameters.AddWithValue("@CUSTOMERCODE", pstrAccountCode.Trim)
                        .Parameters.AddWithValue("@DOC_NO", pstrInvoice.Trim)
                        .Parameters.AddWithValue("@ITEM_CODE", Item_Code.Trim)
                        .Parameters.AddWithValue("@IP_ADDRESS", IP_Address.Trim)
                        SqlConnectionclass.ExecuteNonQuery(oSqlCmdLG)
                    End With
                    strSql = "select *,CONVERT(VARCHAR,EXPIRY_DT,104) AS EXPIRYDATE from  TBL_ASN_TEXT_M_01 WHERE S_FREIGHT_REF_NO=Right('" & pstrInvoice & "',8) AND "
                    strSql = strSql & "UNIT_CODE='" & UnitCode & "' AND Article_Code_Supplier='" & Item_Code & "'"
                    strSql = strSql & "AND ACCOUNT_CODE = '" & pstrAccountCode & "'"


                    'adap = New SqlDataAdapter(strSql, SqlConnectionclass.GetConnection)
                    'ds = New DataSet
                    'adap.Fill(ds)

                    ds = New DataSet
                    ''Dim sqlcmd_ds As New SqlCommand
                    With sqlcmd_ds
                        .CommandText = strSql
                        .CommandType = CommandType.Text
                    End With
                    ds = SqlConnectionclass.GetDataSet(sqlcmd_ds)

                    If ds.Tables(0).Rows.Count = 0 Then
                        ASN_4913_MP2 = "FALSE|No Invoice Records were Found."
                        FileClose(1)
                        Exit Function
                    End If
                    FileClose(1)
                    FileOpen(1, strFileName, OpenMode.Append)
                    strRecord_H = ""

                    'HEADER'

                    ''Pallete Details
                    If Len(pstrInvoice) > 0 Then
                        Dim seq_p As String = ""
                        strSql = ""
                        If cnt_T = 0 Then
                            seq_p = "1"
                        Else
                            seq_p = "0"
                        End If

                        With oSqlCmdLG
                            .Parameters.Clear()
                            .CommandType = CommandType.StoredProcedure
                            .CommandTimeout = 0
                            .CommandText = "USP_ASN_PALLETE_MP2"
                            .Parameters.AddWithValue("@UNIT_CODE", UnitCode.Trim)
                            .Parameters.AddWithValue("@ACCOUNT_NO", pstrAccountCode.Trim)
                            .Parameters.AddWithValue("@DOC_NO", pstrInvoice.Trim)
                            .Parameters.AddWithValue("@ITEM_CODE", Item_Code.Trim)
                            .Parameters.AddWithValue("@P_SEQ", seq_p.Trim)
                            .Parameters.AddWithValue("@IP_ADDRESS", IP_Address.Trim)
                            SqlConnectionclass.ExecuteNonQuery(oSqlCmdLG)
                        End With

                        strSql_Pallet = ""
                        strSql_Pallet = "select * from  ASN_PALLETE_MP2_TEMP where UNIT_CODE='" & UnitCode & "' and ACCOUNT_CODE='" & pstrAccountCode & "'"
                        strSql_Pallet = strSql_Pallet & " and DOC_NO='" & pstrInvoice & "' order by pallet_no"


                        'adap_Pallet = New SqlDataAdapter(strSql_Pallet, SqlConnectionclass.GetConnection)
                        'ds_Pallet = New DataSet
                        'adap_Pallet.Fill(ds_Pallet)

                        ds_Pallet = New DataSet
                        ''Dim sqlcmd_Pallet As New SqlCommand
                        With sqlcmd_Pallet
                            .CommandText = strSql_Pallet
                            .CommandType = CommandType.Text
                        End With
                        ds_Pallet = SqlConnectionclass.GetDataSet(sqlcmd_Pallet)



                        If ds_Pallet.Tables(0).Rows.Count = 0 Then
                            ASN_4913_MP2 = "FALSE|No Item's Pallet not Found."
                            FileClose(1)
                            Exit Function
                        End If

                        '''' '''For BOX UNDER PALLET
                        With oSqlCmdLG
                            .Parameters.Clear()
                            .CommandType = CommandType.StoredProcedure
                            .CommandTimeout = 0
                            .CommandText = "USP_ASN_BOX_MP2"
                            .Parameters.AddWithValue("@UNIT_CODE", UnitCode.Trim)
                            .Parameters.AddWithValue("@ACCOUNT_NO", pstrAccountCode.Trim)
                            .Parameters.AddWithValue("@DOC_NO", pstrInvoice.Trim)
                            .Parameters.AddWithValue("@ITEM_CODE", Item_Code.Trim)
                            .Parameters.AddWithValue("@IP_ADDRESS", IP_Address.Trim)
                            SqlConnectionclass.ExecuteNonQuery(oSqlCmdLG)
                        End With


                        ' For cnt_pallet_no = 0 To ds_Pallet.Tables(0).Rows.Count - 1
                        Dim Invoice_String_O As String = ""
                        Dim Invoice_String_A As String = ""

                        If ds_Pallet.Tables(0).Rows.Count > 0 And ds.Tables(0).Rows.Count > 0 Then
                            For cnt = 0 To ds.Tables(0).Rows.Count - 1
                                strRecord_H = ds.Tables(0).Rows(cnt).Item("HDR")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Cust_Plant_Code")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Supplier_Plant_Code")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Transm_No_O")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Transm_No_N")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Transm_Date")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Subcontratcor_No")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Freight_Carrier_No")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("STOCK_KEEPER_CODE")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Delivery_Identification")

                                '//strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("S_Freight_Ref_No")

                                Invoice_String_O = ds.Tables(0).Rows(cnt).Item("S_Freight_Ref_No").ToString.Trim
                                If Len(Invoice_String_O) >= 8 Then
                                    Invoice_String_A = Invoice_String_O.Substring(Len(Invoice_String_O) - 8)
                                    strRecord_H = strRecord_H & "," & Invoice_String_A
                                Else
                                    strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("S_Freight_Ref_No")
                                End If

                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Plant_Supplier")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Freight_Carrier")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("FREIGHT_CARR_TRANS_DT")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("FREIGHT_CARR_TRANS_TM")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Gr_Wt")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Nt_Wt")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Freight_Code")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Carrier_EDI_Code")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("No_Of_Pkg_Pcs")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Trnsprt_Partnr_No")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Means_Of_Trnsprt_Code")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Means_Of_Trnsprt_No")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Code_POS17")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Content_Conform_POS16")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Arrival_Date_Target")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Arrival_Time_Target")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Loading_Meter")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Code_Type_Truck")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Shipment_Inv_No")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Delivery_Date")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Unload_Point")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Type_Of_Dispatch")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Client_Re_LAB")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Closing_Ord_No")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Remark1")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Transaction_Code")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Client_Plant")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Consignment")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("NO_OF_RCVR")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Remark2")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Client_Storage_Location")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Supplier_No")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Point_of_consumption")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Call_Off_Number")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Client_Ref")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Client_Doc_No")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Article_Code_Client")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Article_Code_Supplier")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Country_Of_Origin")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Delivery_Quantity1")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Quantity_Unit1")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Delivery_Quantity2")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Quantity_Unit2")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Vat_rate")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Position_No_Ship_Inv")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Batch_No")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Use_Code")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Preference_Status")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Customs_Goods")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Stock_Status")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Changed_Dispatch_Code")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("BLANK")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("EXPIRYDATE")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("INDEX_VALUE")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Remark3")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Remark4")
                                PrintLine(1, strRecord_H) : intLineNo = intLineNo + 1
                            Next
                            For cnt_Pallet = 0 To ds_Pallet.Tables(0).Rows.Count - 1
                                strSql_BOX = ""
                                strSql_BOX = "SELECT DTL,CUSTOMER_PKG_NO,ITEM_CODE,SUPPLIER_PKG_NO,NO_OF_PKGS,POSTION_NO_SHIPMENT_INV," &
                                       " PRODUCT_PER_PACKING,MIN(PACKAGE_NO_FROM) AS PACKAGE_NO_FROM," &
                                       " LABEL_IDENTIFICATION,MAX(PACKAGE_NO_FROM) AS PACKAGE_NO_TO " &
                                       " FROM  ASN_BOX_MP2_TEMP WHERE UNIT_CODE='" & UnitCode & "' AND ACCOUNT_CODE='" & pstrAccountCode & "'"
                                strSql_BOX = strSql_BOX & " AND DOC_NO='" & pstrInvoice & "' AND ITEM_CODE='" & Item_Code & "' " &
                                " AND PALLET_NO='" & ds_Pallet.Tables(0).Rows(cnt_Pallet).Item("PALLET_NO").ToString.Trim & "' "
                                strSql_BOX = strSql_BOX & " GROUP BY CUSTOMER_PKG_NO,SUPPLIER_PKG_NO,NO_OF_PKGS,POSTION_NO_SHIPMENT_INV,PRODUCT_PER_PACKING, LABEL_IDENTIFICATION,ITEM_CODE,DTL"

                                ds_BOX = New DataSet
                                ''Dim sqlcmd_BOX As New SqlCommand
                                With sqlcmd_BOX
                                    .CommandText = strSql_BOX
                                    .CommandType = CommandType.Text
                                End With
                                ds_BOX = SqlConnectionclass.GetDataSet(sqlcmd_BOX)


                                ''
                                If ds_Pallet.Tables(0).Rows.Count > 0 And ds_BOX.Tables(0).Rows.Count > 0 Then
                                    strRecord_P = ds_Pallet.Tables(0).Rows(cnt_Pallet).Item("DTL").ToString.Trim
                                    strRecord_P = strRecord_P & "," & ds_Pallet.Tables(0).Rows(cnt_Pallet).Item("Customer_Pkg_No").ToString.Trim
                                    strRecord_P = strRecord_P & "," & ds_Pallet.Tables(0).Rows(cnt_Pallet).Item("Supplier_Pkg_No").ToString.Trim

                                    Dim PALLET_01 As Integer
                                    Dim STR_PALLET_SEARCH As String
                                    Dim DS_PALLET_SEARCH As DataSet
                                    Dim SQLCMD_PALLET_SEARCH As New SqlCommand

                                    PALLET_01 = ds_Pallet.Tables(0).Rows(cnt_Pallet).Item("PALLET_NO").ToString.Trim

                                    STR_PALLET_SEARCH = "SELECT PALLET_NO,UPDATED_STATUS FROM MULTIPLE_PALLET_STATUS WHERE UNIT_CODE='" & UnitCode & "' "
                                    STR_PALLET_SEARCH = STR_PALLET_SEARCH & " AND DOC_NO='" & pstrInvoice & "' AND  PALLET_NO='" & PALLET_01 & "'"
                                    DS_PALLET_SEARCH = New DataSet

                                    With SQLCMD_PALLET_SEARCH
                                        .CommandText = STR_PALLET_SEARCH
                                        .CommandType = CommandType.Text
                                    End With
                                    DS_PALLET_SEARCH = SqlConnectionclass.GetDataSet(SQLCMD_PALLET_SEARCH)

                                    Dim PALLET_PACKAGES_SEQ As Integer = 0
                                    Dim UPDATED_STATUS As Integer = 0

                                    If DS_PALLET_SEARCH.Tables(0).Rows.Count > 0 Then
                                        Dim STRSQL_PALLET_UPDATE_STATUS As String
                                        Dim SQLCMD_PALLET_UPDATE_STATUS As New SqlCommand

                                        UPDATED_STATUS = Convert.ToInt32(DS_PALLET_SEARCH.Tables(0).Rows(0).Item("UPDATED_STATUS").ToString.Trim)

                                        If UPDATED_STATUS = 0 Then
                                            STRSQL_PALLET_UPDATE_STATUS = "UPDATE MULTIPLE_PALLET_STATUS SET UPDATED_STATUS=1 WHERE UNIT_CODE='" & UnitCode & "' AND DOC_NO='" & pstrInvoice & "' AND  PALLET_NO='" & PALLET_01 & "'"
                                            With SQLCMD_PALLET_UPDATE_STATUS
                                                .CommandText = STRSQL_PALLET_UPDATE_STATUS
                                                .CommandType = CommandType.Text
                                            End With
                                            SqlConnectionclass.ExecuteNonQuery(SQLCMD_PALLET_UPDATE_STATUS)
                                            PALLET_PACKAGES_SEQ = 1
                                        Else
                                            PALLET_PACKAGES_SEQ = 0
                                        End If
                                    End If


                                    'strRecord_P = strRecord_P & "," & ds_Pallet.Tables(0).Rows(cnt_Pallet).Item("No_Of_Pkgs").ToString.Trim
                                    strRecord_P = strRecord_P & "," & PALLET_PACKAGES_SEQ

                                    strRecord_P = strRecord_P & "," & ds_Pallet.Tables(0).Rows(cnt_Pallet).Item("Postion_No_Shipment_Inv").ToString.Trim
                                    strRecord_P = strRecord_P & "," & ds_Pallet.Tables(0).Rows(cnt_Pallet).Item("Product_Per_Packing").ToString.Trim
                                    strRecord_P = strRecord_P & "," & ds_Pallet.Tables(0).Rows(cnt_Pallet).Item("Package_No_From").ToString.Trim
                                    strRecord_P = strRecord_P & "," & ds_Pallet.Tables(0).Rows(cnt_Pallet).Item("Label_Identification").ToString.Trim
                                    'strRecord_P = strRecord_P & "," & ds_Pallet.Tables(0).Rows(cnt_Pallet).Item("Pallet_No").ToString.Trim
                                    PrintLine(1, strRecord_P) : intLineNo = intLineNo + 1
                                    ''
                                    For cnt_box = 0 To ds_BOX.Tables(0).Rows.Count - 1
                                        min1 = Convert.ToInt32(ds_BOX.Tables(0).Rows(cnt_box).Item("Package_No_From").ToString.Trim)
                                        max1 = Convert.ToInt32(ds_BOX.Tables(0).Rows(cnt_box).Item("PACKAGE_NO_TO").ToString.Trim)
                                        no_of_package = (max1 - min1) + 1
                                        strRecord_D = ds_BOX.Tables(0).Rows(cnt_box).Item("DTL").ToString.Trim
                                        strRecord_D = strRecord_D & "," & ds_BOX.Tables(0).Rows(cnt_box).Item("Customer_Pkg_No").ToString.Trim
                                        strRecord_D = strRecord_D & "," & ds_BOX.Tables(0).Rows(cnt_box).Item("Supplier_Pkg_No").ToString.Trim
                                        strRecord_D = strRecord_D & "," & no_of_package
                                        strRecord_D = strRecord_D & "," & ds_BOX.Tables(0).Rows(cnt_box).Item("Postion_No_Shipment_Inv").ToString.Trim
                                        strRecord_D = strRecord_D & "," & ds_BOX.Tables(0).Rows(cnt_box).Item("Product_Per_Packing").ToString.Trim
                                        strRecord_D = strRecord_D & "," & ds_BOX.Tables(0).Rows(cnt_box).Item("Package_No_From").ToString.Trim
                                        strRecord_D = strRecord_D & "," & ds_BOX.Tables(0).Rows(cnt_box).Item("Label_Identification").ToString.Trim
                                        strRecord_D = strRecord_D & "," & ds_BOX.Tables(0).Rows(cnt_box).Item("PACKAGE_NO_TO").ToString.Trim
                                        PrintLine(1, strRecord_D) : intLineNo = intLineNo + 1
                                    Next
                                End If
                            Next
                        End If
                        'Next
                    End If
                Next
            End If
            FileClose(1)
            ASN_4913_MP2 = "TRUE|Invoice Text File Generated Successfully."
            Exit Function

        Catch Ex As Exception
            ASN_4913_MP2 = Ex.Message.ToString()
            Exit Function
        Finally
            'Kill(strFileName)
            'FileClose(1)
        End Try

    End Function

    Public Function isGalliaASN() As Boolean
        Dim strsql As String
        Try

            strsql = "select distinct Key2 from lists where unit_code = '" + gstrUNITID + "' and key1 like 'GaliaASNCustomer' " &
            " and key2 ='" + txtCustomerCode.Text + "' "

            If IsRecordExists(strsql) Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception

        End Try
    End Function

    Public Function isYazakiASN() As Boolean
        Dim strsql As String
        Try

            strsql = "select distinct Key2 from lists where unit_code = '" + gstrUNITID + "' and key1 like 'YazakiASNCustomer' " &
            " and key2 ='" + txtCustomerCode.Text + "' "

            If IsRecordExists(strsql) Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception

        End Try
    End Function
    Public Function isBoschASN() As Boolean
        Dim strsql As String
        Try

            strsql = "select distinct Key2 from lists where unit_code = '" + gstrUNITID + "' and key1='BOSCH VDA LABEL PRINT' " &
            " and key2 ='" + txtCustomerCode.Text + "' "

            If IsRecordExists(strsql) Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception

        End Try
    End Function

    Public Function ASN_Gallia_MP2(ByVal UnitCode As String, ByVal pstrAccountCode As String, ByVal pstrInvoice As String, ByVal pstrSRV As String) As String
        Dim strLocation As String
        Dim strFileName As String
        Dim intLineCount As Short
        Dim intLineWidth As Short
        Dim intLineNo As Short
        Dim strSql As String
        Dim strsql2 As String

        Dim strRecord_H As String
        Dim strRecord_D As String
        Dim cnt As Integer = 0
        Dim cnt_T As Integer = 0

        Dim oSqlCmdLG As New SqlCommand
        Dim adap As SqlDataAdapter
        Dim ds As DataSet

        Dim ds_T As DataSet
        Dim adap_T As SqlDataAdapter

        Dim Item_Code As String = ""
        Dim IP_Address As String = ""

        Dim strSql_BOX As String = "'"
        Dim adap_BOX As SqlDataAdapter
        Dim ds_BOX As DataSet
        Dim cnt_box As Integer

        Try

            Dim min1, max1, no_of_package As Integer
            min1 = 0
            max1 = 0
            no_of_package = 0

            Dim cnt_pallet_no As Integer
            Dim sqlcmd_T As New SqlCommand
            Dim sqlcmd_ds As New SqlCommand
            Dim sqlcmd_BOX As New SqlCommand

            IP_Address = gstrIpaddressWinSck
            strLocation = Trim(Find_Value("SELECT ISNULL(TextFileDefaultLocation,'') FROM SALES_PARAMETER_ASN_MP2 WHERE UNIT_CODE='" & UnitCode & "'"))
            'strLocation = "C:\MUL_Txtfile"

            If Len(strLocation) = 0 Then
                ASN_Gallia_MP2 = "FALSE|Default location not defined in sales_parameter."
                Exit Function
            Else
                If Mid(Trim(strLocation), Len(Trim(strLocation))) <> "\" Then
                    strLocation = strLocation & "\"
                End If
                If Directory.Exists(strLocation) = False Then
                    Directory.CreateDirectory(strLocation)
                End If
                strFileName = strLocation & UnitCode & "_" & pstrInvoice & ".txt"

                If System.IO.File.Exists(strFileName) Then
                    Kill(strFileName)
                    FileClose(1)
                End If

                FileOpen(1, strFileName, OpenMode.Append)

            End If
            intLineCount = 60
            intLineWidth = 180
            Dim rsTextfileNonDI As ADODB.Recordset
            Dim strSalesTextFile As String

            strSql = ""
            strSql = "SELECT sales_dtl.Item_Code FROM sales_dtl INNER JOIN SalesChallan_Dtl ON sales_dtl.UNIT_CODE = "
            strSql = strSql & "SalesChallan_Dtl.UNIT_CODE AND sales_dtl.Location_Code = SalesChallan_Dtl.Location_Code "
            strSql = strSql & "AND    sales_dtl.Doc_No = SalesChallan_Dtl.Doc_No WHERE (sales_dtl.Doc_No = '" & pstrInvoice & "')"
            strSql = strSql & "AND (sales_dtl.UNIT_CODE = '" & UnitCode & "') AND (saleschallan_dtl.Account_Code='" & pstrAccountCode & "')"

            ds_T = New DataSet
            With sqlcmd_T
                .CommandText = strSql
                .CommandType = CommandType.Text
            End With
            ds_T = SqlConnectionclass.GetDataSet(sqlcmd_T)

            If ds_T.Tables(0).Rows.Count >= 1 Then
                For cnt_T = 0 To ds_T.Tables(0).Rows.Count - 1
                    Item_Code = ds_T.Tables(0).Rows(cnt_T).Item("Item_Code")
                    With oSqlCmdLG
                        .Parameters.Clear()
                        .CommandType = CommandType.StoredProcedure
                        .CommandTimeout = 0
                        .CommandText = "USP_ASN_TEXT_M_01"
                        .Parameters.AddWithValue("@UNIT_CODE", UnitCode.Trim)
                        .Parameters.AddWithValue("@CUSTOMERCODE", pstrAccountCode.Trim)
                        .Parameters.AddWithValue("@DOC_NO", pstrInvoice.Trim)
                        .Parameters.AddWithValue("@ITEM_CODE", Item_Code.Trim)
                        .Parameters.AddWithValue("@IP_ADDRESS", IP_Address.Trim)
                        SqlConnectionclass.ExecuteNonQuery(oSqlCmdLG)
                    End With

                    'strSql = "select * from  TBL_ASN_TEXT_M_01 WHERE S_FREIGHT_REF_NO=Right('" & pstrInvoice & "',8) AND "
                    strSql = "select *, CONVERT(VARCHAR,EXPIRY_DT,104) AS EXPIRYDATE from  TBL_ASN_TEXT_M_01 WHERE S_FREIGHT_REF_NO=Right('" & pstrInvoice & "',8) AND "
                    strSql = strSql & "UNIT_CODE='" & UnitCode & "' AND Article_Code_Supplier='" & Item_Code & "'"
                    strSql = strSql & "AND ACCOUNT_CODE = '" & pstrAccountCode & "'"
                    ds = New DataSet

                    With sqlcmd_ds
                        .CommandText = strSql
                        .CommandType = CommandType.Text
                    End With
                    ds = SqlConnectionclass.GetDataSet(sqlcmd_ds)

                    If ds.Tables(0).Rows.Count = 0 Then
                        ASN_Gallia_MP2 = "FALSE|No Invoice Records were Found."
                        FileClose(1)
                        Exit Function
                    End If
                    FileClose(1)
                    FileOpen(1, strFileName, OpenMode.Append)
                    strRecord_H = ""

                    If Len(pstrInvoice) > 0 Then
                        Dim seq_p As String = ""
                        strSql = ""
                        If cnt_T = 0 Then
                            seq_p = "1"
                        Else
                            seq_p = "0"
                        End If

                        With oSqlCmdLG
                            .Parameters.Clear()
                            .CommandType = CommandType.StoredProcedure
                            .CommandTimeout = 0
                            .CommandText = "USP_ASN_BOX_MP2"
                            .Parameters.AddWithValue("@UNIT_CODE", UnitCode.Trim)
                            .Parameters.AddWithValue("@ACCOUNT_NO", pstrAccountCode.Trim)
                            .Parameters.AddWithValue("@DOC_NO", pstrInvoice.Trim)
                            .Parameters.AddWithValue("@ITEM_CODE", Item_Code.Trim)
                            .Parameters.AddWithValue("@IP_ADDRESS", IP_Address.Trim)
                            SqlConnectionclass.ExecuteNonQuery(oSqlCmdLG)
                        End With

                        Dim Invoice_String_O As String = ""
                        Dim Invoice_String_A As String = ""

                        If ds.Tables(0).Rows.Count > 0 Then
                            For cnt = 0 To ds.Tables(0).Rows.Count - 1
                                strRecord_H = ds.Tables(0).Rows(cnt).Item("HDR")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Cust_Plant_Code")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Supplier_Plant_Code")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Transm_No_O")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Transm_No_N")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Transm_Date")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Subcontratcor_No")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Freight_Carrier_No")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("STOCK_KEEPER_CODE")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Delivery_Identification")

                                Invoice_String_O = ds.Tables(0).Rows(cnt).Item("S_Freight_Ref_No").ToString.Trim
                                If Len(Invoice_String_O) >= 8 Then
                                    Invoice_String_A = Invoice_String_O.Substring(Len(Invoice_String_O) - 8)
                                    strRecord_H = strRecord_H & "," & Invoice_String_A
                                Else
                                    strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("S_Freight_Ref_No")
                                End If

                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Plant_Supplier")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Freight_Carrier")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("FREIGHT_CARR_TRANS_DT")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("FREIGHT_CARR_TRANS_TM")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Gr_Wt")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Nt_Wt")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Freight_Code")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Carrier_EDI_Code")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("No_Of_Pkg_Pcs")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Trnsprt_Partnr_No")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Means_Of_Trnsprt_Code")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Means_Of_Trnsprt_No")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Code_POS17")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Content_Conform_POS16")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Arrival_Date_Target")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Arrival_Time_Target")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Loading_Meter")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Code_Type_Truck")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Shipment_Inv_No")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Delivery_Date")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Unload_Point")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Type_Of_Dispatch")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Client_Re_LAB")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Closing_Ord_No")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Remark1")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Transaction_Code")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Client_Plant")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Consignment")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("NO_OF_RCVR")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Remark2")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Client_Storage_Location")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Supplier_No")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Point_of_consumption")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Call_Off_Number")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Client_Ref")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Client_Doc_No")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Article_Code_Client")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Article_Code_Supplier")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Country_Of_Origin")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Delivery_Quantity1")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Quantity_Unit1")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Delivery_Quantity2")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Quantity_Unit2")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Vat_rate")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Position_No_Ship_Inv")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Batch_No")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Use_Code")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Preference_Status")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Customs_Goods")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Stock_Status")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Changed_Dispatch_Code")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("BLANK")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("EXPIRYDATE")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("INDEX_VALUE")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Remark3")
                                strRecord_H = strRecord_H & "," & ds.Tables(0).Rows(cnt).Item("Remark4")
                                PrintLine(1, strRecord_H) : intLineNo = intLineNo + 1
                            Next

                            strSql_BOX = ""
                            strSql_BOX = "SELECT DTL,CUSTOMER_PKG_NO,ITEM_CODE,SUPPLIER_PKG_NO,NO_OF_PKGS,POSTION_NO_SHIPMENT_INV," &
                                   " PRODUCT_PER_PACKING,MIN(PACKAGE_NO_FROM) AS PACKAGE_NO_FROM," &
                                   " LABEL_IDENTIFICATION,MAX(PACKAGE_NO_FROM) AS PACKAGE_NO_TO " &
                                   " FROM  ASN_BOX_MP2_TEMP WHERE UNIT_CODE='" & UnitCode & "' AND ACCOUNT_CODE='" & pstrAccountCode & "'"
                            strSql_BOX = strSql_BOX & " AND DOC_NO='" & pstrInvoice & "' AND ITEM_CODE='" & Item_Code & "' "
                            strSql_BOX = strSql_BOX & " GROUP BY CUSTOMER_PKG_NO,SUPPLIER_PKG_NO,NO_OF_PKGS,POSTION_NO_SHIPMENT_INV,PRODUCT_PER_PACKING, LABEL_IDENTIFICATION,ITEM_CODE,DTL"

                            ds_BOX = New DataSet
                            ''Dim sqlcmd_BOX As New SqlCommand
                            With sqlcmd_BOX
                                .CommandText = strSql_BOX
                                .CommandType = CommandType.Text
                            End With
                            ds_BOX = SqlConnectionclass.GetDataSet(sqlcmd_BOX)

                            If ds_BOX.Tables(0).Rows.Count > 0 Then
                                For cnt_box = 0 To ds_BOX.Tables(0).Rows.Count - 1
                                    min1 = Convert.ToInt32(ds_BOX.Tables(0).Rows(cnt_box).Item("Package_No_From").ToString.Trim)
                                    max1 = Convert.ToInt32(ds_BOX.Tables(0).Rows(cnt_box).Item("PACKAGE_NO_TO").ToString.Trim)
                                    no_of_package = (max1 - min1) + 1
                                    strRecord_D = ds_BOX.Tables(0).Rows(cnt_box).Item("DTL").ToString.Trim
                                    strRecord_D = strRecord_D & "," & ds_BOX.Tables(0).Rows(cnt_box).Item("Customer_Pkg_No").ToString.Trim
                                    strRecord_D = strRecord_D & "," & ds_BOX.Tables(0).Rows(cnt_box).Item("Supplier_Pkg_No").ToString.Trim
                                    strRecord_D = strRecord_D & "," & no_of_package
                                    strRecord_D = strRecord_D & "," & ds_BOX.Tables(0).Rows(cnt_box).Item("Postion_No_Shipment_Inv").ToString.Trim
                                    strRecord_D = strRecord_D & "," & ds_BOX.Tables(0).Rows(cnt_box).Item("Product_Per_Packing").ToString.Trim
                                    strRecord_D = strRecord_D & "," & ds_BOX.Tables(0).Rows(cnt_box).Item("Package_No_From").ToString.Trim
                                    strRecord_D = strRecord_D & "," & ds_BOX.Tables(0).Rows(cnt_box).Item("Label_Identification").ToString.Trim
                                    strRecord_D = strRecord_D & "," & ds_BOX.Tables(0).Rows(cnt_box).Item("PACKAGE_NO_TO").ToString.Trim
                                    PrintLine(1, strRecord_D) : intLineNo = intLineNo + 1
                                Next
                            End If
                        End If
                    End If
                Next
            End If
            FileClose(1)
            ASN_Gallia_MP2 = "TRUE|Invoice Text File Generated Successfully."
            Exit Function

        Catch Ex As Exception
            ASN_Gallia_MP2 = Ex.Message.ToString()
            Kill(strFileName)
            FileClose(1)
        End Try

    End Function

    Private Sub dtToDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtToDate.ValueChanged
        On Error GoTo ErrHandler
        txtInvoice_no.Text = ""
        Me.spgrid.MaxRows = 0
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub dtFromDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtFromDate.ValueChanged
        On Error GoTo ErrHandler
        txtInvoice_no.Text = ""
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
    Private Sub spgrid_ClickEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles spgrid.ClickEvent
        txtInvoice_no.Text = ""
        Dim strinvoices As Object = Nothing
        With spgrid
            .Row = .ActiveRow
            .Col = .ActiveCol
            If .Col = 2 Then
                txtInvoice_no.Text = .Text
            Else
                txtInvoice_no.Text = ""
            End If

        End With
    End Sub

    Private Sub spgrid_DblClick(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_DblClickEvent) Handles spgrid.DblClick

        txtInvoice_no.Text = ""
        Dim strinvoices As Object = Nothing
        With spgrid
            .Row = .ActiveRow
            .Col = .ActiveCol
            If .Col = 2 Then
                txtInvoice_no.Text = .Text
            Else
                txtInvoice_no.Text = ""
            End If
        End With
    End Sub

    Private Sub frmMKTTRN0096_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        msqlcon = SqlConnectionclass.GetConnection(gstrConnectSQLClient)
        Call DisableControls()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub frmMKTTRN0096_Deactivate(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub frmMKTTRN0096_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub

    Private Sub frmMKTTRN0096_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Public Function GenerateYazakiASN(ByVal UnitCode As String, ByVal pstrAccountCode As String, ByVal pstrInvoice As String, ByVal pstrSRV As String) As String

        Dim connString As String = "Server=" & gstrCONNECTIONSERVER & ";Database=" & gstrDatabaseName & ";user=" & gstrCONNECTIONUSER & ";password=" & gstrCONNECTIONPASSWORD & " "
        Dim sqlCmd As New SqlCommand
        Dim strFileName As String
        Try
            Dim outputFilePath As String = Trim(Find_Value("SELECT ISNULL(asnfilepath,'') FROM customer_mst WHERE customer_code='" & pstrAccountCode & "' and UNIT_CODE='" & UnitCode & "'"))
            'Dim outputFilePath As String = "D:\procedure_output.txt"
            If outputFilePath.Length = 0 Then
                GenerateYazakiASN = "FALSE|Default location not defined in customer_mst."
                Exit Function
            Else
                If Mid(Trim(outputFilePath), Len(Trim(outputFilePath))) <> "\" Then
                    outputFilePath = outputFilePath & "\"
                End If
                If Directory.Exists(outputFilePath) = False Then
                    Directory.CreateDirectory(outputFilePath)
                End If
                strFileName = outputFilePath & UnitCode & "_" & pstrInvoice & ".txt"

                If System.IO.File.Exists(strFileName) Then
                    Kill(strFileName)
                    FileClose(1)
                End If
            End If
            Using conn As New SqlConnection(connString)
                Using cmd As New SqlCommand("YAZAKIASN", conn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.Parameters.AddWithValue("@UNITCODE", UnitCode.Trim)
                    cmd.Parameters.AddWithValue("@CUSTCODE", pstrAccountCode.Trim)
                    cmd.Parameters.AddWithValue("@INVOICENO", pstrInvoice.Trim)
                    '.Parameters.AddWithValue("@IP_ADDRESS", IP_Address.Trim)
                    conn.Open()
                    Using writer As New StreamWriter(strFileName, False)
                        Using reader As SqlDataReader = cmd.ExecuteReader()
                            If reader.HasRows Then
                                ' Write column headers
                                'For i As Integer = 0 To reader.FieldCount - 1
                                '    writer.Write(reader.GetName(i) & vbTab)
                                'Next
                                'writer.WriteLine()

                                ' Write each row of data
                                While reader.Read()
                                    For i As Integer = 0 To reader.FieldCount - 1
                                        writer.Write(reader(i).ToString() & vbTab)
                                    Next
                                    writer.WriteLine()
                                End While
                            Else
                                GenerateYazakiASN = "FALSE|Default location not defined in customer_mst."
                            End If
                        End Using
                    End Using

                End Using
            End Using
            GenerateYazakiASN = "TRUE|Invoice Text File Generated Successfully."
        Catch ex As Exception
            GenerateYazakiASN = ex.Message
            'Console.WriteLine("Error: " & ex.Message)
        End Try
    End Function
End Class