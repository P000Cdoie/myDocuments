Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0043
    Inherits System.Windows.Forms.Form
    '===================================================================================
    '(c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
    'File Name          :   frmMKTTRN0043.frm
    'Function           :   Used for Invoice Text File Generation
    'Created By         :   Arshad Ali
    'Created On         :   18 July, 2005
    'Revision  By       :
    'Revision On        :
    'History            :
    'Revised By         : Manoj Kr. Vaish
    'Issue ID           : eMpro-20090216-27468
    'Revision Date      : 16-Feb-2009
    'History            : Assign dsn name to prj_InvoicePrinting Dll and rectification of .Net conversion
    '---------------------------------------------------------------------------------------
    'Modified by    :   Virendra Gupta
    'Modified ON    :   18/05/2011
    'Modified to support MultiUnit functionality
    '---------------------------------------------------------------------------------------
    '***********************************************************************************
    '===================================================================================
    Dim mintFormIndex As Double
    Dim objInvoicePrint As New prj_InvoicePrinting.clsTextPrinting(gstrDateFormat)

    Private Sub chkSchefencker_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkSchefencker.CheckStateChanged
        On Error GoTo ErrHandler
        If chkSchefencker.CheckState Then
            optNotGenerated.Checked = False
            optRange.Checked = False
            Call EnableDisable()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdGRNFrom_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGRNFrom.Click
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim strGRN() As String
        Dim intCount As Short
        Dim strAccounCode As String
        For intCount = 0 To lvwCustomers.Items.Count - 1
            If lvwCustomers.Items.Item(intCount).Checked = True Then
                strAccounCode = Trim(lvwCustomers.Items.Item(intCount).Text)
            End If
        Next
        If Len(strAccounCode) = 0 Then
            MsgBox("Please select customer.", MsgBoxStyle.Information, "eMPro")
            lvwCustomers.Focus()
            Exit Sub
        End If
        strsql = "select distinct convert(char(20),Doc_No) as Doc_no, " & DateColumnNameInShowList("ddt") & "  as grin_date FROM printedSRV_dtl P"
        strsql = strsql & " INNER JOIN Cust_Ord_Hdr ch ON p.account_code = ch.account_code and p.cust_ref = ch.cust_ref and p.Unit_Code = ch.Unit_code"
        strsql = strsql & " WHERE ch.account_code='" & strAccounCode & "' and ch.Unit_Code= '" & gstrUNITID & "' order by doc_no"
        mP_Connection.CommandTimeout = 0
        strGRN = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strsql, "GRIN Details")
        mP_Connection.CommandTimeout = 30
        If UBound(strGRN) < 0 Then Exit Sub
        If strGRN(0) = "0" Then
            MsgBox("No Record Available to Display", MsgBoxStyle.Information, "eMPro")
            Exit Sub
        Else
            txtGRNFrom.Text = IIf(IsDBNull(strGRN(0)), "", strGRN(0))
            txtGRNTo.Focus()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdGRNTo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGRNTo.Click
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim strGRN() As String
        Dim intCount As Short
        Dim strAccounCode As String

        For intCount = 0 To lvwCustomers.Items.Count - 1
            If lvwCustomers.Items.Item(intCount).Checked = True Then
                strAccounCode = Trim(lvwCustomers.Items.Item(intCount).Text)
            End If
        Next
        If Len(strAccounCode) = 0 Then
            MsgBox("Please select customer.", MsgBoxStyle.Information, "eMPro")
            lvwCustomers.Focus()
            Exit Sub
        End If
        If Len(Trim(txtGRNFrom.Text)) = 0 Then
            MsgBox("Please select from GRIN first.", MsgBoxStyle.Information, "eMPro")
            txtGRNFrom.Focus()
            Exit Sub
        End If
        strsql = "Select distinct convert(char(20),Doc_No) as Doc_no,  " & DateColumnNameInShowList("ddt") & "   as Grin_Date FROM printedSRV_dtl P"
        strsql = strsql & " INNER JOIN Cust_Ord_Hdr ch ON p.account_code = ch.account_code and p.cust_ref = ch.cust_ref and p.Unit_Code = ch.Unit_code"
        strsql = strsql & " WHERE ch.account_code='" & strAccounCode & "' and ch.Unit_Code= '" & gstrUNITID & "'"
        strsql = strsql & " and doc_no >'" & Trim(txtGRNFrom.Text) & "' order by doc_no"
        mP_Connection.CommandTimeout = 0
        strGRN = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strsql, "GRIN Details")
        mP_Connection.CommandTimeout = 30
        If UBound(strGRN) < 0 Then Exit Sub
        If strGRN(0) = "0" Then
            MsgBox("No Record Available to Display", MsgBoxStyle.Information, "eMPro")
            Exit Sub
        Else
            txtGRNTo.Text = IIf(IsDBNull(strGRN(0)), "", strGRN(0))
            Call FillGRNDetails()
            System.Windows.Forms.SendKeys.Send(vbTab)
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdInvoice_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.cmdGrpAuthorise.ButtonClickEventArgs) Handles cmdInvoice.ButtonClick
        Dim strAccountCode As String
        Dim strInvoice As String
        Dim strSRV As String
        Dim strmessage As String
        Dim intCount As Short
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_AUTHORIZE
                For intCount = 0 To lvwCustomers.Items.Count - 1
                    If lvwCustomers.Items.Item(intCount).Checked = True Then
                        strAccountCode = Trim(lvwCustomers.Items.Item(intCount).Text)
                    End If
                Next
                If chkSchefencker.Visible = True And chkSchefencker.CheckState Then
                    strInvoice = ""
                    For intCount = 0 To lstInvoice.Items.Count - 1
                        If lstInvoice.Items.Item(intCount).Checked = True Then
                            strInvoice = strInvoice & "'" & Trim(lstInvoice.Items.Item(intCount).Text) & "',"
                        End If
                    Next
                    If Len(strInvoice) > 0 Then strInvoice = Mid(strInvoice, 1, Len(strInvoice) - 1)
                    If Len(strInvoice) = 0 Then
                        MsgBox("Please select at least one Invoice.", MsgBoxStyle.Information, ResolveResString(100))
                        Exit Sub
                    End If
                    strmessage = SchefenckerTextFile(strAccountCode, strInvoice)
                End If
                objInvoicePrint = New prj_InvoicePrinting.clsTextPrinting(gstrDateFormat)
                objInvoicePrint.mstrDSNforTextPrint = gstrDSNName

                ' objInvoicePrint 
                If optNotGenerated.Checked Then
                    strmessage = objInvoicePrint.GenerateInvoiceTextFileOfNotGenerated(gstrUNITID, strAccountCode)
                    objInvoicePrint = Nothing
                ElseIf optRange.Checked Then
                    strInvoice = ""
                    strSRV = ""
                    For intCount = 0 To lstInvoice.Items.Count - 1
                        If lstInvoice.Items.Item(intCount).Checked = True Then
                            strInvoice = strInvoice & "'" & Trim(lstInvoice.Items.Item(intCount).Text) & "',"
                        End If
                    Next
                    For intCount = 0 To lstGRN.Items.Count - 1
                        If lstGRN.Items.Item(intCount).Checked = True Then
                            strSRV = strSRV & "'" & Trim(lstGRN.Items.Item(intCount).Text) & "',"
                        End If
                    Next
                    If Len(strInvoice) > 0 Then strInvoice = Mid(strInvoice, 1, Len(strInvoice) - 1)
                    If Len(strSRV) > 0 Then strSRV = Mid(strSRV, 1, Len(strSRV) - 1)
                    If Len(strInvoice) = 0 And Len(strSRV) = 0 Then
                        MsgBox("Please select at least either one Invoice or one GRIN.", MsgBoxStyle.Information, "eMPro")
                        Exit Sub
                    End If
                    If gstrUNITID = "MST" Then
                        strmessage = objInvoicePrint.GenerateInvoiceTextFileOfRange(gstrUNITID, strAccountCode, strInvoice, strSRV, chkspare.Checked)
                    Else
                        strmessage = objInvoicePrint.GenerateInvoiceTextFileOfRange(gstrUNITID, strAccountCode, strInvoice, strSRV)
                    End If

                    objInvoicePrint = Nothing
                End If
                If UCase(Mid(strmessage, 1, 5)) = "FALSE" Then
                    MsgBox(Mid(strmessage, 7), MsgBoxStyle.Critical, ResolveResString(100))
                ElseIf UCase(Mid(strmessage, 1, 4)) = "TRUE" Then
                    MsgBox(Mid(strmessage, 6), MsgBoxStyle.Information, ResolveResString(100))
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_REFRESH
                Call RefreshForm()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
        End Select
    End Sub
    Private Sub cmdInvoiceFrom_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdInvoiceFrom.Click
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim strInvoice() As String
        Dim intCount As Short
        Dim strAccounCode As String
        Dim blnIRNewayRequired As Boolean = False
        For intCount = 0 To lvwCustomers.Items.Count - 1
            If lvwCustomers.Items.Item(intCount).Checked = True Then
                strAccounCode = Trim(lvwCustomers.Items.Item(intCount).Text)
            End If
        Next
        If Len(strAccounCode) = 0 Then
            MsgBox("Please select customer.", MsgBoxStyle.Information, "eMPro")
            lvwCustomers.Focus()
            Exit Sub
        End If
        strsql = "select convert(char(20),Doc_No) as Doc_no , " & DateColumnNameInShowList("invoice_date") & "   as invoice_date from salesChallan_dtl "
        strsql = strsql & " Where Account_code='" & strAccounCode & "' and Unit_Code= '" & gstrUNITID & "' "
        strsql = strsql & " and doc_no like '" & Trim(txtInvoiceFrom.Text) & "%' And bill_flag = 1 And cancel_flag = 0"
        strsql = strsql & "  AND EWAY_IRN_REQUIRED='N' "
        strsql = strsql & " UNION select convert(char(20),S.Doc_No) as Doc_no , " & DateColumnNameInShowList("invoice_date") & "   as invoice_date from salesChallan_dtl S"
        strsql = strsql & " LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO where  S.UNIT_CODE = '" & gstrUNITID & "' and S.Invoice_Type <> 'EXP' "
        strsql = strsql & " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) "
        strsql = strsql & " AND Account_code='" & strAccounCode & "' and S.Unit_Code= '" & gstrUNITID & "' "
        strsql = strsql & " AND S.bill_flag =1 and S.CANCEL_FLAG = 0 "
        strsql = strsql & " order by doc_no desc"
        mP_Connection.CommandTimeout = 0
        strInvoice = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strsql, "Invoice Details", 1)
        mP_Connection.CommandTimeout = 30
        If UBound(strInvoice) < 0 Then Exit Sub
        If strInvoice(0) = "0" Then
            MsgBox("No Record Available to Display", MsgBoxStyle.Information, "eMPro")
            Exit Sub
        Else
            txtInvoiceFrom.Text = IIf(IsDBNull(strInvoice(0)), "", strInvoice(0))
            txtInvoiceTo.Focus()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdInvoiceTo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdInvoiceTo.Click
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim strInvoice() As String
        Dim intCount As Short
        Dim strAccounCode As String
        For intCount = 0 To lvwCustomers.Items.Count - 1
            If lvwCustomers.Items.Item(intCount).Checked = True Then
                strAccounCode = Trim(lvwCustomers.Items.Item(intCount).Text)
            End If
        Next
        If Len(strAccounCode) = 0 Then
            MsgBox("Please select customer.", MsgBoxStyle.Information, "eMPro")
            lvwCustomers.Focus()
            Exit Sub
        End If
        If Len(Trim(txtInvoiceFrom.Text)) = 0 Then
            MsgBox("Please select from Invoice first.", MsgBoxStyle.Information, "eMPro")
            txtInvoiceFrom.Focus()
            Exit Sub
        End If
        strsql = "select convert(char(20),Doc_No) as Doc_no, " & DateColumnNameInShowList("invoice_date") & "  as invoice_date from salesChallan_dtl "
        strsql = strsql & " Where Account_code='" & strAccounCode & "' and Doc_No >= '" & Trim(txtInvoiceFrom.Text) & "' and Unit_Code= '" & gstrUNITID & "'"
        strsql = strsql & " and doc_no like '" & Trim(txtInvoiceTo.Text) & "%' And bill_flag = 1 And cancel_flag = 0"
        strsql = strsql & "  AND EWAY_IRN_REQUIRED='N' "
        strsql = strsql & " UNION select convert(char(20),s.Doc_No) as Doc_no , " & DateColumnNameInShowList("invoice_date") & "   as invoice_date from salesChallan_dtl S"
        strsql = strsql & " LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO where  S.UNIT_CODE = '" & gstrUNITID & "' and S.Invoice_Type <> 'EXP' "
        strsql = strsql & " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) "
        strsql = strsql & " AND Account_code='" & strAccounCode & "' and S.Unit_Code= '" & gstrUNITID & "' "
        strsql = strsql & " and s.Doc_No >= '" & Trim(txtInvoiceFrom.Text) & "' "
        strsql = strsql & " AND S.bill_flag =1 and S.CANCEL_FLAG = 0 "

        strsql = strsql & " order by doc_no desc"
        strInvoice = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strsql, "Invoice Details")
        If UBound(strInvoice) < 0 Then Exit Sub
        If strInvoice(0) = "0" Then
            MsgBox("No Record Available to Display", MsgBoxStyle.Information, "eMPro")
            Exit Sub
        Else
            txtInvoiceTo.Text = IIf(IsDBNull(strInvoice(0)), "", strInvoice(0))
            Call FillInvoiceDetails()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0043_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '----------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To intialise required
        '----------------------------------------------------
        On Error GoTo Err_Handler
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0043_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '----------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To relesed Values
        '----------------------------------------------------
        On Error GoTo Err_Handler
        frmModules.NodeFontBold(Tag) = False
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0043_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '----------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To initialise required data
        '----------------------------------------------------
        On Error GoTo Err_Handler
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FitToClient(Me, fraInvoice, ctlFormHeader1, cmdInvoice) 'To fit the form in the MDI
        Call FillCustomerDetails()
        Dim blnSchefFile As Boolean
        blnSchefFile = CBool(Trim(Find_Value("SELECT ISNULL(SchefenckerFile,0) FROM SALES_PARAMETER where Unit_Code= '" & gstrUNITID & "'")))
        If blnSchefFile Then
            chkSchefencker.Visible = True
        Else
            chkSchefencker.Visible = False
        End If
        
        With Me.lstInvoice
            .Items.Clear()
            .View = System.Windows.Forms.View.Details
            .GridLines = True
            .CheckBoxes = True
            .Enabled = True
            .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            .Columns.Insert(0, "")
            .Columns.Insert(1, "")
            .Columns.Item(0).Text = "Invoice No"
            .Columns.Item(1).Text = "Invoice Date"
            .Columns.Item(0).Width = VB6.TwipsToPixelsX(2000)
            .Columns.Item(1).Width = VB6.TwipsToPixelsX(2000)
        End With
        With Me.lstGRN
            .Items.Clear()
            .View = System.Windows.Forms.View.Details
            .GridLines = True
            .CheckBoxes = True
            .Enabled = True
            .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            .Columns.Insert(0, "")
            .Columns.Insert(1, "")
            .Columns.Item(0).Text = "GRIN No"
            .Columns.Item(1).Text = "GRIN Date"
            .Columns.Item(0).Width = VB6.TwipsToPixelsX(2000)
            .Columns.Item(1).Width = VB6.TwipsToPixelsX(2000)
        End With
        Call EnableDisable()
        cmdInvoice.Enabled(0) = True
        cmdInvoice.Enabled(1) = False
        cmdInvoice.Caption(1) = "Refresh"
        cmdInvoice.Enabled(2) = False
        cmdInvoice.Caption(0) = "Generate"
        optNotGenerated.Checked = True
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0043_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '----------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Release memory
        '----------------------------------------------------
        On Error GoTo Err_Handler
        'REFRESH
        'Removing the form name from list
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        'Setting the corresponding node's tag
        frmModules.NodeFontBold(Tag) = False
        'Closing the recordset
        'Releasing the form reference
        Me.Dispose()
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function Find_Value(ByRef strField As String) As String
        '----------------------------------------------------------------------------
        'Author         :   Arshad Ali
        'Argument       :   Sql query string as strField
        'Return Value   :   selected table field value as String
        'Function       :   Return a field value from a table
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim RS As New ADODB.Recordset
        RS = New ADODB.Recordset
        RS.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        RS.Open(strField, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
        If RS.RecordCount > 0 Then
            If IsDBNull(RS.Fields(0).Value) = False Then
                Find_Value = RS.Fields(0).Value
            Else
                Find_Value = ""
            End If
        Else
            Find_Value = ""
        End If
        RS.Close()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Sub FillCustomerDetails()
        With Me.lvwCustomers
            .Items.Clear()
            .View = System.Windows.Forms.View.Details
            .GridLines = True
            .CheckBoxes = True
            .Enabled = True
            .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            'sub to populate the list view for customers
            Call PopulateListView((Me.lvwCustomers), "customer_code,cust_name", "customer_mst", " WHERE customer_code IN (SELECT DISTINCT(account_code) FROM cust_ord_hdr where Unit_Code= '" & gstrUNITID & "' ) and Unit_Code= '" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
            .Columns.Item(0).Text = "Customer Code"
            .Columns.Item(1).Text = "Customer Name"
            .Columns.Item(0).Width = VB6.TwipsToPixelsX(1500)
            .Columns.Item(1).Width = VB6.TwipsToPixelsX(3400)
            If Me.lvwCustomers.Items.Count = 0 Then
                MsgBox("No Records Exist for the Selected Period.", MsgBoxStyle.Information, "eMPower")
                Exit Sub
            End If
        End With
    End Sub
    Private Sub PopulateListView(ByRef ctlListViewName As System.Windows.Forms.ListView, ByVal pstrSQLfields As String, ByVal pstrTableName As String, Optional ByVal pstrCond As String = "", Optional ByVal pstrSQL As String = "")
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim rstCodes As ClsResultSetDB
        Dim LstItem As System.Windows.Forms.ListViewItem
        Dim lngLoop, lngloop1 As Integer
        Dim lngRows As Integer
        Dim strFields() As String 'To Get Info About Fields Wanted
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.AppStarting)
        If Len(pstrSQL) > 0 Then
            strsql = pstrSQL
        Else
            strsql = " Select DISTINCT " & pstrSQLfields & " From " & pstrTableName
            If Len(Trim(pstrCond)) > 0 Then 'If Condtion is SENT
                strsql = strsql & pstrCond
            End If
        End If
        strFields = Split(pstrSQLfields, ",") 'Split Fields Info
        rstCodes = New ClsResultSetDB
        mP_Connection.Execute("set dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        ctlListViewName.Items.Clear()
        Call rstCodes.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        Dim strpos As Object
        strpos = InStr(1, Trim(strFields(0)), "=")
        If strpos > 0 Then
            strFields(0) = Mid(strFields(0), 1, strpos - 1)
        End If
        For lngLoop = 0 To UBound(strFields) 'ADD Column Headers
            ctlListViewName.Columns.Insert(lngLoop, "", Replace(Trim(strFields(lngLoop)), "_", " ", , , CompareMethod.Text), -2) 'To Replace "_" in Field Name with a Space (" ")
        Next
        lngRows = rstCodes.GetNoRows
        rstCodes.MoveFirst()
        If lngRows > 0 Then
            For lngLoop = 0 To lngRows - 1
                LstItem = ctlListViewName.Items.Add(Trim(rstCodes.GetValue(Trim(strFields(0)))))
                For lngloop1 = 1 To rstCodes.GetFieldCount - 1
                    If LstItem.SubItems.Count > lngloop1 Then
                        LstItem.SubItems(lngloop1).Text = rstCodes.GetValueByNo(lngloop1)
                    Else
                        LstItem.SubItems.Insert(lngloop1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rstCodes.GetValueByNo(lngloop1)))
                    End If
                Next lngloop1
                rstCodes.MoveNext()
            Next lngLoop
        End If
        rstCodes.ResultSetClose()
        rstCodes = Nothing
        ctlListViewName.Columns.Item(0).Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(ctlListViewName.Width)) / 2 - 300)
        ctlListViewName.Columns.Item(1).Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(ctlListViewName.Width) - (VB6.PixelsToTwipsX(ctlListViewName.Columns.Item(1).Width) + 400))) ' - ctlListViewName.ColumnHeaders(1).Width) - 80
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Me.Cursor = System.Windows.Forms.Cursors.Default
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub search(ByRef lvwListView As System.Windows.Forms.ListView, ByRef txtSearchBox As System.Windows.Forms.TextBox, ByRef optFistOption As System.Windows.Forms.RadioButton, ByRef optSecOption As System.Windows.Forms.RadioButton)
        On Error GoTo ErrHandler
        Dim Intcounter As Short
        With lvwListView
            If optFistOption.Checked = True Then
                For Intcounter = 0 To .Items.Count - 1
                    If .Items.Item(Intcounter).Font.Bold = True Then
                        .Items.Item(Intcounter).Font = VB6.FontChangeBold(.Items.Item(Intcounter).Font, False)
                        .Refresh()
                    End If
                    If .Items.Item(Intcounter).SubItems.Item(1).Font.Bold = True Then
                        .Items.Item(Intcounter).SubItems.Item(1).Font = VB6.FontChangeBold(.Items.Item(Intcounter).SubItems.Item(1).Font, False)
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
                    If .Items.Item(Intcounter).SubItems.Item(1).Font.Bold = True Then
                        .Items.Item(Intcounter).SubItems.Item(1).Font = VB6.FontChangeBold(.Items.Item(Intcounter).SubItems.Item(1).Font, False)
                        .Refresh()
                    End If
                Next
                If Len(txtSearchBox.Text) = 0 Then Exit Sub
                For Intcounter = 0 To .Items.Count - 1
                    If Trim(UCase(Mid(.Items.Item(Intcounter).SubItems.Item(1).Text, 1, Len(txtSearchBox.Text)))) = Trim(UCase(txtSearchBox.Text)) Then
                        .Items.Item(Intcounter).Font = VB6.FontChangeBold(.Items.Item(Intcounter).Font, True)
                        Call .Items.Item(Intcounter).EnsureVisible()
                        .Refresh()
                        Exit For
                    End If
                Next
            End If
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub optAllGRN_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAllGRN.CheckedChanged
        If eventSender.Checked Then
            Dim intCount As Short
            If optAllGRN.Checked Then
                For intCount = 0 To lstGRN.Items.Count - 1
                    lstGRN.Items.Item(intCount).Checked = True
                Next
            Else
                For intCount = 0 To lstGRN.Items.Count - 1
                    lstGRN.Items.Item(intCount).Checked = False
                Next
            End If
        End If
    End Sub
    Private Sub optallInvoice_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptAllInvoice.CheckedChanged
        If eventSender.Checked Then
            Dim intCount As Short
            If OptAllInvoice.Checked Then
                For intCount = 0 To lstInvoice.Items.Count - 1
                    lstInvoice.Items.Item(intCount).Checked = True
                Next
            Else
                For intCount = 0 To lstInvoice.Items.Count - 1
                    lstInvoice.Items.Item(intCount).Checked = False
                Next
            End If
        End If
    End Sub
    Private Sub optNotGenerated_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optNotGenerated.CheckedChanged
        If eventSender.Checked Then
            If Me.chkSchefencker.Visible = True And Me.chkSchefencker.CheckState Then
                If Me.optNotGenerated.Checked = True Then
                    Me.optNotGenerated.Checked = False
                End If
            Else
                Call EnableDisable()
            End If
        End If
    End Sub
    Private Sub optRange_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optRange.CheckedChanged
        If eventSender.Checked Then
            If chkSchefencker.Visible = True And chkSchefencker.CheckState Then
                If optRange.Checked = True Then
                    optRange.Checked = False
                End If
            Else
                Call EnableDisable()
            End If
        End If
    End Sub
    Private Sub optSelectedGRN_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSelectedGRN.CheckedChanged
        If eventSender.Checked Then
            Dim intCount As Short
            If optAllGRN.Checked Then
                For intCount = 0 To lstGRN.Items.Count - 1
                    lstGRN.Items.Item(intCount).Checked = True
                Next
            Else
                For intCount = 0 To lstGRN.Items.Count - 1
                    lstGRN.Items.Item(intCount).Checked = False
                Next
            End If
        End If
    End Sub
    Private Sub OptSelectedInvoice_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptSelectedInvoice.CheckedChanged
        If eventSender.Checked Then
            Dim intCount As Short
            If OptAllInvoice.Checked Then
                For intCount = 0 To lstInvoice.Items.Count - 1
                    lstInvoice.Items.Item(intCount).Checked = True
                Next
            Else
                For intCount = 0 To lstInvoice.Items.Count - 1
                    lstInvoice.Items.Item(intCount).Checked = False
                Next
            End If
        End If
    End Sub
    Private Sub txtGRNFrom_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtGRNFrom.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If Len(Trim(txtGRNFrom.Text)) > 0 And KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send(vbTab)
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtGRNFrom_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtGRNFrom.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If cmdGRNFrom.Enabled Then Call cmdGRNFrom_Click(cmdGRNFrom, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtGRNTo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtGRNTo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If Len(Trim(txtGRNTo.Text)) > 0 And KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send(vbTab)
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtGRNTo_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtGRNTo.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If cmdGRNTo.Enabled Then Call cmdGRNTo_Click(cmdGRNTo, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtInvoiceFrom_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtInvoiceFrom.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If Len(Trim(txtInvoiceFrom.Text)) > 0 And KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send(vbTab)
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtInvoiceFrom_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtInvoiceFrom.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If cmdInvoiceFrom.Enabled Then Call cmdInvoiceFrom_Click(cmdInvoiceFrom, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtInvoiceTo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtInvoiceTo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If Len(Trim(txtInvoiceTo.Text)) > 0 And KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send(vbTab)
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtInvoiceTo_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtInvoiceTo.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If cmdInvoiceTo.Enabled Then Call cmdInvoiceTo_Click(cmdInvoiceTo, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtInvoiceTo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtInvoiceTo.Leave
        If Me.txtInvoiceFrom.Text <> "" And Me.txtInvoiceTo.Text <> "" Then
            Call FillInvoiceDetails()
        End If

    End Sub
    Private Sub txtSearchCustomer_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSearchCustomer.TextChanged
        On Error GoTo ErrHandler
        'sub to search in the list box
        Call search((Me.lvwCustomers), (Me.txtSearchCustomer), (Me.optSearchCustCode), (Me.optSearchCustName))
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Sub EnableDisable()
        If optNotGenerated.Checked Then
            frmInvoice.Enabled = False
            frmSRV.Enabled = False
            'Invoice
            txtInvoiceFrom.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            cmdInvoiceFrom.Enabled = False
            txtInvoiceTo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            cmdInvoiceTo.Enabled = False
            OptAllInvoice.Enabled = False
            OptSelectedInvoice.Enabled = False
            lstInvoice.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            'SRV 57F4
            txtGRNFrom.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            cmdGRNFrom.Enabled = False
            txtGRNTo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            cmdGRNTo.Enabled = False
            optAllGRN.Enabled = False
            optSelectedGRN.Enabled = False
            lstGRN.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        Else
            frmInvoice.Enabled = True
            frmSRV.Enabled = True
            'Invoice
            txtInvoiceFrom.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            cmdInvoiceFrom.Enabled = True
            txtInvoiceTo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            cmdInvoiceTo.Enabled = True
            OptAllInvoice.Enabled = True
            OptSelectedInvoice.Enabled = True
            lstInvoice.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            'SRV 57F4
            txtGRNFrom.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            cmdGRNFrom.Enabled = True
            txtGRNTo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            cmdGRNTo.Enabled = True
            optAllGRN.Enabled = True
            optSelectedGRN.Enabled = True
            lstGRN.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        End If
    End Sub
    Sub RefreshForm()
        Dim Intcounter As Short
        For Intcounter = 0 To lvwCustomers.Items.Count - 1
            lvwCustomers.Items.Item(Intcounter).Checked = False
        Next
        txtSearchCustomer.Text = ""
        optNotGenerated.Checked = True
        txtInvoiceFrom.Text = ""
        txtInvoiceTo.Text = ""
        OptAllInvoice.Checked = False : OptSelectedInvoice.Checked = False
        lstInvoice.Items.Clear()
        txtGRNFrom.Text = ""
        txtGRNTo.Text = ""
        optAllGRN.Checked = False : optSelectedGRN.Checked = False
        lstGRN.Items.Clear()
        lvwCustomers.Focus()
        cmdInvoice.Enabled(0) = True
        cmdInvoice.Enabled(1) = False
        cmdInvoice.Caption(1) = "Refresh"
        cmdInvoice.Enabled(2) = False
        cmdInvoice.Caption(0) = "Generate"
    End Sub
    Sub FillInvoiceDetails()
        Dim intCount As Short
        Dim strAccounCode As String
        For intCount = 0 To lvwCustomers.Items.Count - 1
            If lvwCustomers.Items.Item(intCount).Checked = True Then
                strAccounCode = Trim(lvwCustomers.Items.Item(intCount).Text)
            End If
        Next
        lstInvoice.Columns.RemoveAt(1)
        lstInvoice.Columns.RemoveAt(0)
        If txtInvoiceFrom.Text.Length > 0 And txtInvoiceTo.Text.Length > 0 Then
            Call PopulateListView((Me.lstInvoice), "Doc_no, Invoice_date ", "vw_maruti_invoicetextfile", " WHERE Account_code='" & strAccounCode & "' and Doc_No >= '" & Trim(txtInvoiceFrom.Text) & "' and Unit_Code = '" & gstrUNITID & "' and Doc_No <= '" & Trim(txtInvoiceTo.Text) & "' And bill_flag = 1 And cancel_flag = 0")
            For intCount = 0 To lstInvoice.Items.Count - 1
                lstInvoice.Items.Item(intCount).Checked = True
            Next
        End If

        
        OptAllInvoice.Checked = True
        OptAllInvoice.Focus()
    End Sub
    Sub FillGRNDetails()
        Dim intCount As Short
        Dim strAccounCode As String
        Dim strsql As String
        For intCount = 0 To lvwCustomers.Items.Count - 1
            If lvwCustomers.Items.Item(intCount).Checked = True Then
                strAccounCode = Trim(lvwCustomers.Items.Item(intCount).Text)
            End If
        Next
        lstGRN.Columns.RemoveAt(2)
        lstGRN.Columns.RemoveAt(1)
        strsql = "select distinct Doc_No, ddt as Grin_Date FROM printedSRV_dtl P"
        strsql = strsql & " INNER JOIN Cust_Ord_Hdr ch ON p.account_code = ch.account_code and p.cust_ref = ch.cust_ref and p.unit_code = ch.Unit_code"
        strsql = strsql & " WHERE ch.account_code='" & strAccounCode & "' and ch.Unit_Code= '" & gstrUNITID & "'"
        strsql = strsql & " and doc_no >='" & Trim(txtGRNFrom.Text) & "'"
        strsql = strsql & " and doc_no <='" & Trim(txtGRNTo.Text) & "' order by doc_no"
        Call PopulateListView((Me.lstGRN), "Doc_no, GRN_Date", "", "", strsql)
        For intCount = 0 To lstGRN.Items.Count - 1
            lstGRN.Items.Item(intCount).Checked = True
        Next
        optAllGRN.Checked = True
        optAllGRN.Focus()
    End Sub
    Private Function SchefenckerTextFile(ByRef pstrAccountCode As String, ByRef pstrInvoice As String) As String
        '----------------------------------------------------------------------------
        'Author         :   Ashutosh Verma
        'Argument       :   Customer code & Invoice Numbers.
        'Return Value   :   Message string with value False or True.
        'Function       :   Generate text file for Schefencker.
        'Comments       :   Issue Id:19702
        '
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim intCount As Short
        Dim strLocation As String
        Dim strFileName As String
        Dim intLineNo As Short
        Dim strsql As String
        Dim strRecord As String
        Dim rsInvoice As New ClsResultSetDB
        strLocation = Trim(Find_Value("SELECT ISNULL(SchefenckerTextFileLocation,'') FROM SALES_PARAMETER where Unit_Code= '" & gstrUNITID & "'"))
        If Len(strLocation) = 0 Then
            SchefenckerTextFile = "FALSE|Default location not defined in sales_parameter."
            Exit Function
        Else
            If Mid(Trim(strLocation), Len(Trim(strLocation))) <> "\" Then
                strLocation = strLocation & "\"
            End If
            strFileName = strLocation & "INPUTDATA" & ".csv"
            On Error Resume Next
            Kill(strLocation & "*.csv")
            FileClose(1)
            On Error GoTo ErrHandler
            FileOpen(1, strFileName, OpenMode.Append)
        End If
        If Len(pstrInvoice) > 0 Then
            If UCase(GetPlantName) = "MATE" Then
                strsql = "SELECT right(challan_dtl.doc_no,6) as Invoice_no , convert(varchar,challan_dtl.invoice_date,106) as Invoice_date,"
            Else
                strsql = "SELECT challan_dtl.doc_no as Invoice_no , convert(varchar,challan_dtl.invoice_date,106) as Invoice_date,"
            End If
            strsql = strsql & " sales_dtl.cust_item_code as Item_code, sales_dtl.sales_quantity,Isnull(challan_dtl.remarks,'')  as remarks,"
            strsql = strsql & " isnull(ord.cust_ref,'') as PO_No,isnull(sales_dtl.cust_item_Desc,'') as cust_item_Desc ,isnull(sales_dtl.rate,0) as Rate ,isnull(challan_dtl.Total_Amount,0) as Total_Amount "
            strsql = strsql & " from saleschallan_dtl as challan_dtl"
            strsql = strsql & " inner Join Sales_dtl as sales_dtl on challan_dtl.location_code = sales_dtl.location_code and challan_dtl.Doc_No = sales_dtl.doc_no and challan_dtl.unit_code = sales_dtl.Unit_code"
            strsql = strsql & " Inner Join Cust_Ord_Hdr as Ord on challan_dtl.Account_code = Ord.Account_code and sales_dtl.cust_ref = Ord.Cust_ref and Isnull(sales_dtl.Amendment_no,'') = isnull(ord.Amendment_no,'') and sales_dtl.Unit_Code = Ord.Unit_Code"
            strsql = strsql & " Where challan_dtl.bill_flag = 1 And cancel_flag = 0 "
            strsql = strsql & " and challan_dtl.account_code='" & pstrAccountCode & "'"
            strsql = strsql & " and challan_dtl.Unit_Code='" & gstrUNITID & "'"
            strsql = strsql & " and challan_dtl.doc_no in (" & pstrInvoice & ") order by challan_dtl.doc_no"
            rsInvoice.GetResult(strsql)
            If rsInvoice.GetNoRows > 0 Then
                rsInvoice.MoveFirst()
                While Not rsInvoice.EOFRecord
                    strRecord = ""
                    strRecord = rsInvoice.GetValue("Invoice_no")
                    strRecord = strRecord & "," & VB6.Format(rsInvoice.GetValue("Invoice_date"), "dd-MM-yy")
                    strRecord = strRecord & "," & IIf(IsDBNull(rsInvoice.GetValue("PO_No")), "", Trim(rsInvoice.GetValue("PO_No")))
                    strRecord = strRecord & "," & IIf(IsDBNull(rsInvoice.GetValue("Item_Code")), "", Trim(rsInvoice.GetValue("Item_Code")))
                    strRecord = strRecord & "," & IIf(IsDBNull(rsInvoice.GetValue("cust_item_Desc")), "", Trim(rsInvoice.GetValue("cust_item_Desc")))
                    strRecord = strRecord & "," & IIf(IsDBNull(rsInvoice.GetValue("sales_quantity")), 0, rsInvoice.GetValue("sales_quantity"))
                    strRecord = strRecord & "," & IIf(IsDBNull(rsInvoice.GetValue("rate")), 0, rsInvoice.GetValue("rate"))
                    strRecord = strRecord & "," & IIf(IsDBNull(rsInvoice.GetValue("Total_amount")), 0, rsInvoice.GetValue("Total_amount"))
                    strRecord = strRecord & "," & IIf(IsDBNull(rsInvoice.GetValue("remarks")), "", Trim(rsInvoice.GetValue("remarks")))
                    PrintLine(1, strRecord) : intLineNo = intLineNo + 1
                    rsInvoice.MoveNext()
                End While
            Else
                SchefenckerTextFile = "FALSE|No Invoice Records found to generate the File."
                FileClose(1)
                Exit Function
            End If
        Else
            SchefenckerTextFile = "FALSE| File Not Generated."
        End If
        rsInvoice.ResultSetClose()
        FileClose(1)
        SchefenckerTextFile = "TRUE| Schefencker File Generated Successfully."
        Exit Function
ErrHandler:
        FileClose(1)
        Kill(strFileName)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        '----------------------------------------------------
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call eMPro help
        '----------------------------------------------------
        MsgBox("No Help Attached to This Form", MsgBoxStyle.Information, "eMPro")
    End Sub

    Private Sub cmdInvoice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInvoice.Load

    End Sub


    Private Sub optSearchCustCode_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSearchCustCode.CheckedChanged
        If eventSender.Checked Then
            Call SortGrid(0)
            txtSearchCustomer.Text = ""
        End If
    End Sub
    Private Sub SortGrid(ByRef Index As Short)
        On Error GoTo ErrHandler
        With lvwCustomers
            .Sort()
            ListViewColumnSorter.SortListView(lvwCustomers, Index, SortOrder.Ascending)
            .Sorting = System.Windows.Forms.SortOrder.Ascending
        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub optSearchCustName_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSearchCustName.CheckedChanged
        If eventSender.Checked Then
            Call SortGrid(1)
            txtSearchCustomer.Text = ""
        End If
    End Sub

    Private Sub chkspare_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class