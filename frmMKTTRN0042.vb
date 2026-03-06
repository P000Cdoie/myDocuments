Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMKTTRN0042
	Inherits System.Windows.Forms.Form
	'---------------------------------------------------------------------------
	'Copyright          :   MIND Ltd.
	'Form Name          :   frmMKTTRN0040
	'Created By         :   Sourabh Khatri
	'Created on         :   29 July 2005
	'Modified Date      :
	'Description        :   Form shall work to post invoice into account
	'---------------------------------------------------------------------------
	'Revision  By       : Ashutosh , Issue Id :17431
	'Revision On        : 30-03-2006
	'History            : Debit credit mismatch problem while posting invoice (MSSLED).
	'-----------------------------------------------------------------------------------
	'Revised By      : Davinder Singh
	'Issue ID        : 19575
	'Revision Date   : 27 Feb 2007
	'History         : New Tax (SEcess) added
	'-----------------------------------------------------------------------------------
	'Revised By      : Manoj Kr. Vaish
	'Issue ID        : 20052
	'Revision Date   : 05 June 2007
	'History         : New Tax (SEcess) added on Service Tax for Job Work Invoice
	'-----------------------------------------------------------------------------------
	'Revised By      : Manoj Kr. Vaish
	'Issue ID        : 19992
	'Revision Date   : 29 June 2007
	'History         : To add the functionality of Multiple SO for Export Invoice.
	'***********************************************************************************
	'Revised By      : Manoj Kr. Vaish
	'Issue ID        : 21551
	'Revision Date   : 20-Nov-2007
	'History         : Add New Tax VAT with Sale Tax help
    '***********************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : eMpro-20080430-18033
    'Revision Date   : 30 Apr 2008
    'History         : Posting of New tax head for the calculation of CVD Excise,Ecess & SEcess
    '***********************************************************************************
    'Revised By      : Manoj Kr.Vaish
    'Issue ID        : eMpro-20080508-18500
    'Revision Date   : 12 May 2008
    'History         : To Allow SAD Tax for Transfer Invoice in Mate Noida

    'Revised By      : ASHISH SHARMA
    'Issue ID        : 101188073
    'Revision Date   : 14 JUL 2017
    'History         : GST CHANGES TO ACCOUNT POSTING
    '***********************************************************************************
    'Modified by AJAY SHUKLA ON 12/MAY/2011 FOR MULTI UNIT CHANGE
    Dim lngCounter As Integer
    Dim blnDate As Boolean = True
    Private Enum InvoiceGrid
        invSel = 1
        InvNo = 2
        invDate = 3
        invTypeDesc = 4
        invSubTypeDesc = 5
        invCustName = 6
        invType = 7
        invSubType = 8
    End Enum
    Dim mStrCustMst As String
    Dim mresult As New ClsResultSetDB
    Dim MintFormIndex As Short
    Dim salesconf As String
    Dim msubTotal, mInvNo, mExDuty, mBasicAmt, mOtherAmt As Double
    Dim mFrAmt, mGrTotal, mStAmt, mCustmtrl As Double
    Dim mDoc_No As Short
    Dim mAccount_Code, mInvType, mSubCat, mlocation As String
    Dim mstrAnnex As String
    Dim arrQty() As Double 'used in BomCheck() insertupdateAnnex()
    Dim arrItem() As String 'used in BomCheck() insertupdateAnnex()
    Dim arrReqQty() As Double
    Dim arrCustAnnex() As Object
    Dim ref57f4 As String 'used in BomCheck() insertupdateAnnex()
    Dim dblFinishedQty As Double 'To get Qty of Finished Item from Spread
    Dim strCustCode As String 'used in BomCheck() insertupdateAnnex()
    Dim strItemcode As String 'used in BomCheck() insertupdateAnnex()
    Dim inti As Short 'To Change Array Size used in BomCheck() insertupdateAnnex()
    Dim strsaledetails As String
    Dim strupdateGrinhdr As String
    Dim strupdateitbalmst As String
    Dim strupdatecustodtdtl As String
    Dim strUpdateAmorDtl As String
    Dim strupdateamordtlbom As String
    Dim mCust_Ref, mAmendment_No As String
    Dim saleschallan As String
    Dim ValidRecord As Boolean
    Dim updatestockflag, updatePOflag As Boolean
    Dim strStockLocation As String
    Dim mAmortization As Double
    Dim mblnEOUUnit As Boolean
    Dim mAssessableValue As Double
    Dim mOpeeningBalance As Double
    Dim mblnCustSupp As Boolean
    Dim strBomItem As String 'For Latest Item To Explore
    Dim blnFIFOFlag As Boolean
    Dim rsBomMst As ClsResultSetDB
    Dim mstrMasterString As String 'To store master string for passing to Dr Cr COM
    Dim mstrDetailString As String 'To store detail string for passing to Dr Cr COM
    Dim mstrPurposeCode As String 'To store the Purpose Code which will be used for the fetching of GL and SL
    Dim mblnAddCustomerMaterial As Boolean 'To decide whether to add customer material in basic or not
    Dim mblnSameSeries As Boolean 'To store the flag whether the selected invoice will have same series as others
    Dim mstrReportFilename As String 'To store the report filename
    Dim mblnInsuranceFlag As Boolean 'To store insurance flag
    Dim mblnpostinfin As Boolean
    Dim mblnExciseRoundOFFFlag As Boolean
    Dim mSaleConfNo As Double
    Dim mstrExcisePriorityUpdationString As String
    Dim intNoCopies As Short
    Dim mblnServiceInvoiceWithoutSO As Boolean
    Dim mstrGrinQtyUpdate As String
    Dim mstrInvRejSQL As String
    Dim mblnJobWkFormulation As String
    Dim cmbInvType As String
    Dim CmbCategory As String
    Dim lbldescription As String
    Dim lblcategory As String
    Dim txtUnitCode As String
    Dim Ctlinvoice As String
    Dim mstrInvoiceType As String
    Dim mstrInvoiceSubType As String
    Dim mintIndex As Short
    Dim mblnMultipleSOAllowed As Boolean
    Dim mblnSORequired As Boolean
    Dim mstrCreditTermId As String
    Private Sub cmbSort_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbSort.SelectedIndexChanged
        spGrid.Row = 0
        spGrid.Row2 = spGrid.MaxRows
        spGrid.Col = 1
        spGrid.Col2 = spGrid.MaxCols
        spGrid.SortBy = 0
        spGrid.set_SortKey(1, Me.cmbSort.SelectedIndex + 2)
        spGrid.set_SortKeyOrder(1, 1)
        spGrid.Action = 25
    End Sub
    Private Sub cmdLockInvoice_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtnEditGrp.ButtonClickEventArgs) Handles cmdLockInvoice.ButtonClick
        Dim strInvoiceNo As String
        On Error GoTo Errorhandler
        Dim blnflag As Boolean
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                Me.cmdLockInvoice.Caption(0) = "POST"
                If Me.spGrid.MaxRows <= 0 Then
                    Call MsgBox("Grid does not contain any invoice", MsgBoxStyle.OkOnly, ResolveResString(100))
                    Call disableControls()
                    Exit Sub
                Else
                    With Me.spGrid
                        blnflag = False
                        For lngCounter = 1 To .MaxRows
                            .Row = lngCounter : .Col = InvoiceGrid.invSel
                            If CDbl(.Value) = 1 Then
                                blnflag = True
                                Exit For
                            End If
                        Next
                        If blnflag = False Then
                            Call MsgBox(" Select at least one invoice for posting", MsgBoxStyle.OkOnly, ResolveResString(100))
                            Call disableControls()
                            Exit Sub
                        End If
                        strInvoiceNo = "Following invoice(s) have been posted successfully"
                        For lngCounter = 1 To .MaxRows
                            .Row = lngCounter : .Col = InvoiceGrid.invSel
                            If CDbl(.Value) = 1 Then
                                .Col = InvoiceGrid.InvNo : Ctlinvoice = Trim(.Text)
                                .Col = InvoiceGrid.invTypeDesc : cmbInvType = Trim(.Text)
                                .Col = InvoiceGrid.invSubTypeDesc : CmbCategory = Trim(.Text)
                                .Col = InvoiceGrid.invType : lbldescription = Trim(.Text) : mstrInvoiceType = Trim(.Text)
                                .Col = InvoiceGrid.invSubType : lblcategory = Trim(.Text) : mstrInvoiceSubType = Trim(.Text)
                                txtUnitCode = Find_Value("Select Location_Code from saleschallan_dtl where doc_no= '" & Ctlinvoice & "' and unit_code='" & gstrUNITID & "'")
                                Call CheckMultipleSOAllowed(cmbInvType, CmbCategory)
                                If PostInvoice() = False Then
                                    If MsgBox("Unable to post invoice no " & Ctlinvoice & ".Do you want to countinue ?", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Else
                                    .Col = InvoiceGrid.InvNo
                                    strInvoiceNo = strInvoiceNo & Trim(.Text) & "," & "  "
                                End If
                            End If
                        Next
                        If Len(strInvoiceNo) > 50 Then
                            strInvoiceNo = Mid(strInvoiceNo, 1, Len(strInvoiceNo) - 1)
                            Call MsgBox(strInvoiceNo, MsgBoxStyle.OkOnly, ResolveResString(100))
                        End If
                        Call disableControls()
                    End With
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                Call disableControls()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
        End Select
        Exit Sub
Errorhandler:
        Me.cmdLockInvoice.Revert() : Me.cmdLockInvoice.Caption(0) = "POST"
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdShowInvoices_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdShowInvoices.Click
        On Error GoTo Errorhandler
        Call ShowPendingInvoices()
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0042_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0042_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0042_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'Add Form Name To Window List
        blnDate = True
        Call disableControls()
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        'Fill Lebels From Resource File
        Call FillLabelFromResFile(Me) 'Fill Labels >From Resource File
        Me.ctlFormHeader1.HeaderString = (Mid(Me.ctlFormHeader1.HeaderString(), InStr(1, Me.ctlFormHeader1.HeaderString(), "-") + 1, Len(Me.ctlFormHeader1.HeaderString())))
        Call FitToClient(Me, frmMain, ctlFormHeader1, cmdLockInvoice)
        SetGridCells()
        Me.dtFromDate.Format = DateTimePickerFormat.Custom
        Me.dtFromDate.CustomFormat = gstrDateFormat
        Me.dtFromDate.Value = GetServerDate()
        Me.dtToDate.Format = DateTimePickerFormat.Custom
        Me.dtToDate.CustomFormat = gstrDateFormat
        Me.dtToDate.Value = GetServerDate()
        Call disableControls()
        blnDate = False
    End Sub
    Private Sub SetGridCells()
        On Error GoTo Errorhandler
        With Me.spGrid
            .MaxRows = 0 : .MaxCols = 8
            .Row = 0 : .set_RowHeight(0, 300)
            .Col = InvoiceGrid.invSel : .Text = "Selected Invoice" : .set_ColWidth(InvoiceGrid.invSel, 1500)
            .Col = InvoiceGrid.InvNo : .Text = "Invoice No" : .set_ColWidth(InvoiceGrid.InvNo, 1000)
            .Col = InvoiceGrid.invDate : .Text = "Invoice Date" : .set_ColWidth(InvoiceGrid.invDate, 1000)
            .Col = InvoiceGrid.invTypeDesc : .Text = "Invoice Type" : .set_ColWidth(InvoiceGrid.invTypeDesc, 1500)
            .Col = InvoiceGrid.invSubTypeDesc : .Text = "Invoice Category" : .set_ColWidth(InvoiceGrid.invSubTypeDesc, 1500)
            .Col = InvoiceGrid.invCustName : .Text = "Customer Name" : .set_ColWidth(InvoiceGrid.invCustName, 4000)
            .Col = InvoiceGrid.invType : .Text = "Invoice Type" : .set_ColWidth(InvoiceGrid.invType, 100)
            .Col = InvoiceGrid.invSubType : .Text = "Invoice Sub Type" : .set_ColWidth(InvoiceGrid.invSubType, 100)
            .Col = InvoiceGrid.invType : .Col2 = InvoiceGrid.invSubType : .BlockMode = True
            .ColHidden = True : .BlockMode = False
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
        strsql = "unpostedinvoice '" & gstrUNITID & "', '" & getDateForDB(Me.dtFromDate.Value) & "','" & getDateForDB(Me.dtToDate.Value) & "'"
        rsobject.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
        With cmdObject
            .let_ActiveConnection(mP_Connection)
            .CommandTimeout = 0
            .CommandType = ADODB.CommandTypeEnum.adCmdText
            .CommandText = strsql
            rsobject = .Execute
        End With
        cmdObject = Nothing
        Me.spGrid.MaxRows = 0
        Me.cmbSearch.Items.Clear()
        Me.cmbSort.Items.Clear()
        Me.txtSearch.Text = ""
        If Not rsobject.EOF Then
            rsobject.MoveFirst()
            lngCounter = 1
            ' Code to fill Combo box
            For lngCounter = 0 To rsobject.Fields.Count - 3
                Me.cmbSearch.Items.Add((rsobject.Fields.Item(lngCounter).Name))
                Me.cmbSort.Items.Add((rsobject.Fields.Item(lngCounter).Name))
            Next
            Me.cmbSearch.SelectedIndex = 0
            Me.cmbSort.SelectedIndex = 0
            lngCounter = 1
            ' Code end here
            With spGrid
                While Not rsobject.EOF
                    AddNewRow()
                    .Row = lngCounter
                    .Col = InvoiceGrid.InvNo : .Text = rsobject.Fields("doc_No").Value
                    .Col = InvoiceGrid.invDate : .Text = setDateFormat(rsobject.Fields("Invoice_Date").Value, gstrDateFormat)
                    .Col = InvoiceGrid.invTypeDesc : .Text = rsobject.Fields("description").Value
                    .Col = InvoiceGrid.invSubTypeDesc : .Text = rsobject.Fields("Sub_Type_Description").Value
                    .Col = InvoiceGrid.invCustName : .Text = rsobject.Fields("Cust_Name").Value
                    .Col = InvoiceGrid.invType : .Text = rsobject.Fields("invoice_type").Value
                    .Col = InvoiceGrid.invSubType : .Text = rsobject.Fields("sub_category").Value
                    rsobject.MoveNext() : lngCounter = lngCounter + 1
                End While

                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                Me.optCheckAll.Checked = True
            End With
        Else
            Call MsgBox("No data found between selected dates", MsgBoxStyle.OkOnly, ResolveResString(100))
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        End If
        rsobject.Close()
        rsobject = Nothing
        Exit Sub
Errorhandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub AddNewRow()
        On Error GoTo Errorhandler
        With spGrid
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows : .set_RowHeight(.Row, 300)
            .Col = InvoiceGrid.invSel : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .Value = CStr(1) : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeCheckCenter = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
            .Col = InvoiceGrid.InvNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
            .Col = InvoiceGrid.invDate : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .Col = InvoiceGrid.invTypeDesc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .Col = InvoiceGrid.invSubTypeDesc : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .Col = InvoiceGrid.invCustName : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .Col = InvoiceGrid.invType : .ColHidden = True
            .Col = InvoiceGrid.invSubType : .ColHidden = True
            .Row = .MaxRows : .Row2 = .MaxRows : .Col = InvoiceGrid.InvNo
            .Col2 = InvoiceGrid.invSubType : .BlockMode = True
            .Lock = True : .BlockMode = False
            .Row = .MaxRows : .Row2 = .MaxRows : .Col = InvoiceGrid.invSel
            .Col2 = InvoiceGrid.InvNo : .BlockMode = True
            .ColsFrozen = True : .BlockMode = False
        End With
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0042_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Me.Dispose()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub optCheckAll_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optCheckAll.CheckedChanged
        If eventSender.Checked Then
            With Me.spGrid
                If .MaxRows > 0 Then
                    For lngCounter = 1 To .MaxRows
                        .Row = lngCounter : .Col = InvoiceGrid.invSel : .Value = CStr(1)
                    Next
                End If
            End With
        End If
    End Sub
    Private Sub optUncheckAll_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optunCheckAll.CheckedChanged
        If eventSender.Checked Then
            With Me.spGrid
                If .MaxRows > 0 Then
                    For lngCounter = 1 To .MaxRows
                        .Row = lngCounter : .Col = InvoiceGrid.invSel : .Value = CStr(0)
                    Next
                End If
            End With
        End If
    End Sub
    Private Sub TxtSearch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSearch.TextChanged
        Dim blnflag As Boolean
        With Me.spGrid
            If CDbl(Trim(CStr(Len(Me.txtSearch.Text)))) > 0 Then
                blnflag = False
                For lngCounter = 1 To .MaxRows
                    .Row = lngCounter : .Col = Me.cmbSearch.SelectedIndex + 2
                    If InStr(1, UCase(VB.Left(.Text, Len(Me.txtSearch.Text))), UCase(Me.txtSearch.Text), CompareMethod.Binary) Then
                        blnflag = True
                        Exit For
                    End If
                Next
                If blnflag Then
                    .Row = lngCounter : .Col = Me.cmbSearch.SelectedIndex + 2
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                End If
            Else
                .Row = 0 : .Col = 0
                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
            End If
        End With
    End Sub
    Private Function PostInvoice() As Boolean
        Dim rsSalesConf As ClsResultSetDB
        Dim rsSalesChallan As ClsResultSetDB
        Dim rssaledtl As New ClsResultSetDB
        Dim strSalesconf As String
        Dim strAccountCode As String
        Dim strCustRef As String
        Dim StrAmendmentNo As String
        Dim intRow As Short
        Dim intLoopcount As Short
        Dim strRetval As String
        Dim objDrCr As New prj_DrCrNote.cls_DrCrNote(GetServerDate)
        Dim strInvoiceDate As String
        Dim rsbom As New ClsResultSetDB
        Dim irowcount As Short
        Dim intRwCount1 As Short
        Dim varItemQty1 As Double
        Dim SALEDTL As String
        PostInvoice = True
        On Error GoTo Err_Handler
        SALEDTL = "select * from Saleschallan_Dtl where Doc_No =" & Ctlinvoice & "  and Location_Code='" & Trim(txtUnitCode) & "' and unit_code='" & gstrUNITID & "'"
        rsSalesChallan = New ClsResultSetDB
        rssaledtl.GetResult(SALEDTL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        strAccountCode = rssaledtl.GetValue("Account_code")
        strCustRef = rssaledtl.GetValue("Cust_ref")
        StrAmendmentNo = rssaledtl.GetValue("Amendment_No")
        strInvoiceDate = VB6.Format(rssaledtl.GetValue("Invoice_Date"), gstrDateFormat)
        rssaledtl.ResultSetClose()
        rssaledtl = Nothing
        strSalesconf = "Select UpdatePO_Flag,UpdateStock_Flag,Stock_Location,OpenningBal,Preprinted_Flag,NoCopies from saleconf where "
        strSalesconf = strSalesconf & "Invoice_type = '" & lbldescription & "' and sub_type = '"
        strSalesconf = strSalesconf & lblcategory & "' and Location_Code='" & Trim(txtUnitCode) & "' and datediff(dd,convert(datetime,'" & getDateForDB(strInvoiceDate) & "',103),convert(datetime,fin_start_date,103))<=0  and datediff(dd,convert(datetime,fin_end_date,103),convert(datetime,'" & getDateForDB(strInvoiceDate) & "',103))<=0 and unit_code='" & gstrUNITID & "'"
        rsSalesConf = New ClsResultSetDB
        rsSalesConf.GetResult(strSalesconf)
        If rsSalesConf.RowCount > 0 Then
            updatePOflag = rsSalesConf.GetValue("UpdatePO_Flag")
            updatestockflag = rsSalesConf.GetValue("UpdateStock_Flag")
            strStockLocation = rsSalesConf.GetValue("Stock_Location")
            mOpeeningBalance = Val(rsSalesConf.GetValue("OpenningBal"))
            intNoCopies = rsSalesConf.GetValue("NoCopies")
        End If
        rsSalesConf.ResultSetClose()
        rsSalesConf = Nothing
        rssaledtl = New ClsResultSetDB
        SALEDTL = "Select Sales_Quantity,Item_code,Cust_Item_Code,toolcost_amount from sales_Dtl where Doc_No = " & Ctlinvoice & " and Location_Code='" & Trim(txtUnitCode) & "' and unit_code='" & gstrUNITID & "'"
        rssaledtl.GetResult(SALEDTL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        intRow = rssaledtl.GetNoRows
        rssaledtl.MoveFirst()
        If InvoiceGeneration() = False Then
            PostInvoice = False
            rssaledtl.ResultSetClose()
            rssaledtl = Nothing
            Exit Function
        End If
        rssaledtl.ResultSetClose()
        rssaledtl = Nothing
        mP_Connection.BeginTrans()
        mP_Connection.Execute("set Dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.Execute("UpDate Forms_Dtl Set PO_No='" & mInvNo & "' where PO_No='" & Trim(Ctlinvoice) & "' and Doc_Type='9999' and unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        If Len(Trim(mstrExcisePriorityUpdationString)) > 0 Then
            mP_Connection.Execute("update Saleschallan_dtl set Excise_type = '" & mstrExcisePriorityUpdationString & "' where Doc_no = " & Ctlinvoice & "  and unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        End If
        If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
            strRetval = objDrCr.SetARInvoiceDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING)
        Else
            prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = "DR"
            strRetval = objDrCr.SetAPDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING)
            prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = ""
        End If
        strRetval = CheckString(strRetval)
        If Not strRetval = "Y" Then
            PostInvoice = False
            MsgBox(strRetval, MsgBoxStyle.Information, "eMPro")
            mP_Connection.RollbackTrans()
            Exit Function
        Else
            mP_Connection.CommitTrans()
            If UCase(Trim(cmbInvType)) = "REJECTION" Then
                If UCase(Trim(CmbCategory)) = "REJECTION" Then
                    mP_Connection.Execute("update salesChallan_Dtl set RejectionPosting = 1 where doc_no = " & mInvNo & " and unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
            End If
            Ctlinvoice = ""
        End If
        Exit Function
Err_Handler:
        If Err.Number = 20545 Then
            Resume Next
        Else
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Function
    Public Sub ValuetoVariables()
        Dim strsql As String
        Dim rsSalesChallan As ClsResultSetDB
        On Error GoTo Err_Handler
        strsql = "select INVOICE_DATE from Saleschallan_Dtl where Doc_No =" & Ctlinvoice & "  and Location_Code='" & Trim(txtUnitCode) & "' and unit_code='" & gstrUNITID & "'"
        rsSalesChallan = New ClsResultSetDB
        rsSalesChallan.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        mInvType = lbldescription
        mSubCat = lblcategory
        mInvNo = CDbl(Ctlinvoice)
        mresult = New ClsResultSetDB
        strsql = " Select Asseccable= SUM(Accessible_amount) from sales_dtl "
        strsql = strsql & " where Doc_No =" & Ctlinvoice & " and Location_Code='" & Trim(txtUnitCode) & "' and unit_code='" & gstrUNITID & "'"
        mresult.GetResult(strsql)
        mAssessableValue = mresult.GetValue("Asseccable")
        mresult.ResultSetClose()
        mresult = Nothing
        rsSalesChallan.ResultSetClose()
        rsSalesChallan = Nothing
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function InvoiceGeneration() As Boolean
        Dim rsCompMst As ClsResultSetDB
        Dim rsGrnHdr As ClsResultSetDB
        Dim rsSalesConf As New ADODB.Recordset
        Dim rsSalesInvoiceDate As ClsResultSetDB
        Dim Phone, Range, RegNo, EccNo, Address, Invoice_Rule As String
        Dim CST, PLA, Fax, EMail, UPST, Division As String
        Dim Commissionerate As String
        Dim strsql As String
        Dim strCompMst, DeliveredAdd As String
        Dim strGRNDate As String
        Dim strVendorInvNo As String
        Dim strVendorInvDate As String
        Dim strCustRefForGrn As String
        Dim strSuffix As String
        'Dim DeliveredAdd As String
        On Error GoTo Err_Handler
        rsCompMst = New ClsResultSetDB
        strCompMst = "Select * from Company_Mst where unit_code='" & gstrUNITID & "'"
        rsCompMst.GetResult(strCompMst)
        If rsCompMst.GetNoRows = 1 Then
            RegNo = rsCompMst.GetValue("Reg_NO")
            EccNo = rsCompMst.GetValue("Ecc_No")
            Range = rsCompMst.GetValue("Range_1")
            Phone = rsCompMst.GetValue("Phone")
            Fax = rsCompMst.GetValue("Fax")
            EMail = rsCompMst.GetValue("Email")
            PLA = rsCompMst.GetValue("PLA_No")
            UPST = rsCompMst.GetValue("LST_No")
            CST = rsCompMst.GetValue("CST_No")
            Division = rsCompMst.GetValue("Division")
            Commissionerate = rsCompMst.GetValue("Commissionerate")
            Invoice_Rule = rsCompMst.GetValue("Invoice_Rule")
        End If
        rsCompMst.ResultSetClose()
        rsCompMst = Nothing
        If rsSalesConf.State = ADODB.ObjectStateEnum.adStateOpen Then rsSalesConf.Close()
        rsSalesConf.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        'FOR FINANCIAL ROLLOVER
        rsSalesConf.Open("SELECT * FROM SaleConf WHERE Invoice_Type='" & lbldescription & "' AND Sub_Type_description ='" & CmbCategory & "' AND Location_Code='" & Trim(txtUnitCode) & "' and datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0  and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If Not rsSalesConf.EOF Then
            mstrPurposeCode = IIf(IsDBNull(rsSalesConf.Fields("inv_GLD_prpsCode").Value), "", Trim(rsSalesConf.Fields("inv_GLD_prpsCode").Value))
            If mstrPurposeCode = "" Then
                MsgBox("Please select Purpose Code in Sales Configuration", MsgBoxStyle.Information, "eMPro")
                CmbCategory = ""
                lblcategory = ""
                cmbInvType = "'"
                lbldescription = ""
                mstrPurposeCode = ""
                Exit Function
            End If
        Else
            MsgBox("No record found in Sales Configuration for the purpose code", MsgBoxStyle.Information, "eMPro")
            CmbCategory = ""
            lblcategory = ""
            cmbInvType = "'"
            lbldescription = ""
            mstrPurposeCode = ""
            rsSalesConf.Close()
            rsSalesConf = Nothing
            Exit Function
        End If
        Call InitializeValues()
        Call ValuetoVariables()
        If Not CreateStringForAccounts() Then
            InvoiceGeneration = False
            Exit Function
        End If
        InvoiceGeneration = True
        rsSalesConf.Close()
        rsSalesConf = Nothing
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub InitializeValues()
        On Error GoTo ErrHandler
        mExDuty = 0 : mInvNo = 0 : mBasicAmt = 0 : msubTotal = 0 : mOtherAmt = 0 : mGrTotal = 0 : mStAmt = 0 : mFrAmt = 0
        mDoc_No = 0 : mCustmtrl = 0 : mAmortization = 0 : mstrAnnex = "" : strupdateGrinhdr = "" : mblnCustSupp = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function GetTaxGlSl(ByVal TaxType As String) As String
        Dim objRecordSet As New ADODB.Recordset
        On Error GoTo ErrHandler
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        objRecordSet.Open("SELECT tx_glCode, tx_slCode FROM fin_TaxGlRel where tx_rowType = 'ARTAX' AND tx_taxId ='" & TaxType & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objRecordSet.EOF Then
            GetTaxGlSl = "N"
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        GetTaxGlSl = Trim(IIf(IsDBNull(objRecordSet.Fields("tx_glCode").Value), "", objRecordSet.Fields("tx_glCode").Value)) & "»" & Trim(IIf(IsDBNull(objRecordSet.Fields("tx_slCode").Value), "", objRecordSet.Fields("tx_slCode").Value))
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GetTaxGlSl = "N"
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
    End Function
    Private Function GetItemGLSL(ByVal InventoryGlGroup As String, ByVal PurposeCode As String) As String
        Dim objRecordSet As New ADODB.Recordset
        Dim strGL As String
        Dim strSL As String
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        objRecordSet.Open("SELECT invGld_glcode, invGld_slcode FROM fin_InvGLGrpDtl WHERE invGld_prpsCode = '" & PurposeCode & "' AND invGld_invGLGrpId = '" & InventoryGlGroup & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objRecordSet.EOF Then
            objRecordSet.Close()
            objRecordSet.Open("SELECT gbl_glCode, gbl_slCode FROM fin_globalGL WHERE gbl_prpsCode = '" & PurposeCode & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If objRecordSet.EOF Then
                GetItemGLSL = "N"
                MsgBox("GL and SL not defined for Purpose Code: " & PurposeCode, MsgBoxStyle.Information, "eMPro")
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            Else
                strGL = Trim(IIf(IsDBNull(objRecordSet.Fields("gbl_glCode").Value), "", objRecordSet.Fields("gbl_glCode").Value))
                strSL = Trim(IIf(IsDBNull(objRecordSet.Fields("gbl_slCode").Value), "", objRecordSet.Fields("gbl_slCode").Value))
            End If
        Else
            strGL = Trim(IIf(IsDBNull(objRecordSet.Fields("invGld_glcode").Value), "", objRecordSet.Fields("invGld_glcode").Value))
            strSL = Trim(IIf(IsDBNull(objRecordSet.Fields("invGld_slcode").Value), "", objRecordSet.Fields("invGld_slcode").Value))
        End If
        If strGL = "" Then
            GetItemGLSL = "N"
            MsgBox("GL and SL not defined for Purpose Code:" & PurposeCode, MsgBoxStyle.Information, "eMPro")
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        GetItemGLSL = strGL & "»" & strSL
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GetItemGLSL = "N"
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
    End Function
    Public Function CheckExcPriority() As Boolean
        Dim strsql As String
        Dim strTaxGL As String
        Dim strTaxSL As String
        Dim rsTaxPriority As ClsResultSetDB
        rsTaxPriority = New ClsResultSetDB
        strsql = "Select * from Tax_PriorityMst where  unit_code='" & gstrUNITID & "'"
        rsTaxPriority.GetResult(strsql)
        If rsTaxPriority.GetNoRows > 0 Then
            rsTaxPriority.MoveFirst()
            CheckExcPriority = True
            If Len(Trim(rsTaxPriority.GetValue("VarExPriority1"))) = 0 Then
                If Len(Trim(rsTaxPriority.GetValue("VarExPriority2"))) = 0 Then
                    If Len(Trim(rsTaxPriority.GetValue("VarExPriority3"))) = 0 Then
                        CheckExcPriority = False
                        Exit Function
                    End If
                End If
            End If
        Else
            CheckExcPriority = False
        End If
    End Function
    Public Function ReturnGLSLAccExcPriority(ByRef pintPriority As Object, ByRef pdblamount As Double) As String()
        Dim strsql As String
        Dim strBalance As String
        Dim strExcGL As String
        Dim strExcSL As String
        Dim StrData(2) As String
        Dim strExcType As String
        Dim rsExGLSLCode As ClsResultSetDB
        Dim rsCheckBalance As ClsResultSetDB
        rsExGLSLCode = New ClsResultSetDB
        rsCheckBalance = New ClsResultSetDB
        strsql = "Select VarExPriority1,VarExGL1,VarExSL1,VarExPriority2,VarExGL2,VarExSL2,VarExPriority3,VarExGL3,VarExSL3 from Tax_PriorityMst where unit_code='" & gstrUNITID & "'"
        rsExGLSLCode.GetResult(strsql)
        rsExGLSLCode.MoveFirst()
        Select Case pintPriority
            Case 1
                strExcGL = Trim(rsExGLSLCode.GetValue("VarExGL1"))
                strExcSL = Trim(rsExGLSLCode.GetValue("VarExSL1"))
                strExcType = Trim(rsExGLSLCode.GetValue("VarExPriority1"))
                If Len(Trim(strExcGL)) > 0 Then
                    If Len(Trim(strExcSL)) > 0 Then
                        '********To check about in case data is found on first Priority
                        strBalance = "Select sum(isnull(br_amount,0)) as br_amount From fin_balRel where br_UntCodeID = '"
                        strBalance = strBalance & Trim(txtUnitCode) & "' and br_slCode = '" & strExcSL & "'"
                        strBalance = strBalance & " and br_glCode = '" & strExcGL & "' and br_UntCodeID='" & gstrUNITID & "'"
                        rsCheckBalance.GetResult(strBalance)
                        If rsCheckBalance.GetNoRows > 0 Then
                            rsCheckBalance.MoveFirst()
                            If Val(rsCheckBalance.GetValue("br_amount")) >= pdblamount Then
                                StrData(0) = strExcGL
                                StrData(1) = strExcSL
                                StrData(2) = strExcType
                                ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                            Else
                                ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                            End If
                        Else
                            ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                        End If
                    Else
                        ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                    End If
                Else
                    ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                End If
            Case 2
                strExcGL = Trim(rsExGLSLCode.GetValue("VarExGL2"))
                strExcSL = Trim(rsExGLSLCode.GetValue("VarExSL2"))
                strExcType = Trim(rsExGLSLCode.GetValue("VarExPriority2"))
                If Len(Trim(strExcGL)) > 0 Then
                    If Len(Trim(strExcSL)) > 0 Then
                        '********To check about in case data is found on first Priority
                        strBalance = "Select sum(isnull(br_amount,0)) as br_amount From fin_balRel where br_UntCodeID = '"
                        strBalance = strBalance & Trim(txtUnitCode) & "' and br_slCode = '" & strExcSL & "'"
                        strBalance = strBalance & " and br_glCode = '" & strExcGL & "' and br_UntCodeID='" & gstrUNITID & "'"
                        rsCheckBalance.GetResult(strBalance)
                        If rsCheckBalance.GetNoRows > 0 Then
                            rsCheckBalance.MoveFirst()
                            If rsCheckBalance.GetValue("br_amount") >= pdblamount Then
                                StrData(0) = strExcGL
                                StrData(1) = strExcSL
                                StrData(2) = strExcType
                                ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                            Else
                                ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                            End If
                        Else
                            ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                        End If
                    Else
                        ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                    End If
                Else
                    ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                End If
            Case 3
                strExcGL = Trim(rsExGLSLCode.GetValue("VarExGL3"))
                strExcSL = Trim(rsExGLSLCode.GetValue("VarExSL3"))
                strExcType = Trim(rsExGLSLCode.GetValue("VarExPriority3"))
                If Len(Trim(strExcGL)) > 0 Then
                    If Len(Trim(strExcSL)) > 0 Then
                        '********To check about in case data is found on first Priority
                        strBalance = "Select sum(isnull(br_amount,0)) as br_amount From fin_balRel where br_UntCodeID = '"
                        strBalance = strBalance & Trim(txtUnitCode) & "' and br_slCode = '" & strExcSL & "'"
                        strBalance = strBalance & " and br_glCode = '" & strExcGL & "' and br_UntCodeID='" & gstrUNITID & "'"
                        rsCheckBalance.GetResult(strBalance)
                        If rsCheckBalance.GetNoRows > 0 Then
                            rsCheckBalance.MoveFirst()
                            If Val(rsCheckBalance.GetValue("br_amount")) >= pdblamount Then
                                StrData(0) = strExcGL
                                StrData(1) = strExcSL
                                StrData(2) = strExcType
                                ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                            Else
                                ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                            End If
                        Else
                            ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                        End If
                    Else
                        ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                    End If
                Else
                    ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                End If
        End Select
    End Function
    Public Function Find_Value(ByRef strField As String) As String
        '----------------------------------------------------------------------------
        'Author         :   Arshad Ali
        'Argument       :   Sql query string as strField
        'Return Value   :   selected table field value as String
        'Function       :   Return a field value from a table
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rs As New ADODB.Recordset
        rs = New ADODB.Recordset
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.Open(strField, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
        If rs.RecordCount > 0 Then
            If IsDBNull(rs.Fields(0).Value) = False Then
                Find_Value = rs.Fields(0).Value
            Else
                Find_Value = ""
            End If
        Else
            Find_Value = ""
        End If
        rs.Close()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CreateStringForAccounts() As Boolean
        '-----------------------------------------------------------------------------------
        'Revised By      : Davinder Singh
        'Issue ID        : 19575
        'Revision Date   : 27 Feb 2007
        'History         : New Tax (SEcess) added
        '-----------------------------------------------------------------------------------
        'Revised By      : Manoj Kr. Vaish
        'Issue ID        : 19992
        'Revision Date   : 27 June 2007
        'History         : Group same Item_Code for diiferent SO.in Multiple SO Export Invoice for posting
        '                : Fetch credit term from sales_dtl for saving in ar_docmaster
        '***************************************************************************************
        'Revised By      : Manoj Kr. Vaish
        'Issue ID        : eMpro-20080430-18033
        'Revision Date   : 27 Apr 2008
        'History         : Posting under New tax head for the calculation of CVD Excise,
        '                  Ecess,SEcess,Add. Ecess on Total Duty & Add. Secess on Total Duty.
        '***********************************************************************************
        '***********************************************************************************
        'Revised By      : Manoj Kr.Vaish
        'Issue ID        : eMpro-20080508-18500
        'Revision Date   : 12 May 2008
        'History         : Posting of SAD Tax for Transfer Invoice in Mate Noida

        'Revised By      : Ashish sharma
        'Issue ID        : 101188073
        'Revision Date   : 14 JUL 2017
        'History         : GST CHANGES FOR ACCOUNT POSTING
        '***********************************************************************************

        Dim objRecordSet As New ADODB.Recordset
        Dim objTmpRecordset As New ADODB.Recordset
        Dim strRetval As String
        Dim strInvoiceNo As String
        Dim strInvoiceDate As String
        Dim strCurrencyCode As String
        Dim dblInvoiceAmt As Double
        Dim dblInvoiceAmtRoundOff_diff As Double
        Dim dblTCStaxAmt As Double
        Dim dblExchangeRate As Double
        Dim dblBasicAmount As Double
        Dim dblBaseCurrencyAmount As Double
        Dim dblTaxAmt As Double
        Dim strTaxType As String
        Dim strCreditTermsID As String
        Dim strBasicDueDate As String
        Dim strPaymentDueDate As String
        Dim strExpectedDueDate As String
        Dim strCustomerGL As String
        Dim strCustomerSL As String
        Dim strTaxGL As String
        Dim strTaxSL As String
        Dim strItemGL As String
        Dim strItemSL As String
        Dim strGlGroupId As String
        Dim dblTaxRate As Double
        Dim varTmp As Object
        Dim dblCCShare As Double
        Dim iCtr As Short
        Dim strCustRef As String
        Dim blnExciseExumpted As Boolean
        Dim dblDiscountAmt As Double
        Dim arrstrExcPriority() As String
        Dim rsFULLExciseAmount As ClsResultSetDB
        Dim dblFullExciseAmount As Double
        Dim blnMsgBox As Boolean
        rsFULLExciseAmount = New ClsResultSetDB
        Dim rsobject As New ADODB.Recordset
        Dim blnFOC As Boolean
        mstrExcisePriorityUpdationString = ""
        blnMsgBox = False
        On Error GoTo ErrHandler
        objRecordSet.Open("SELECT * FROM  saleschallan_dtl WHERE Doc_No='" & Trim(Ctlinvoice) & "' and Location_Code='" & Trim(txtUnitCode) & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objRecordSet.EOF Then
            MsgBox("Invoice details not found", MsgBoxStyle.Information, "empower")
            CreateStringForAccounts = False
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        strInvoiceNo = CStr(mInvNo)
        strInvoiceDate = VB6.Format(objRecordSet.Fields("Invoice_Date").Value, "dd-MMM-yyyy")
        strCurrencyCode = Trim(IIf(IsDBNull(objRecordSet.Fields("Currency_Code").Value), "", objRecordSet.Fields("Currency_Code").Value))
        dblInvoiceAmt = IIf(IsDBNull(objRecordSet.Fields("total_amount").Value), 0, objRecordSet.Fields("total_amount").Value)
        dblInvoiceAmtRoundOff_diff = IIf(IsDBNull(objRecordSet.Fields("TotalInvoiceAmtRoundOff_diff").Value), 0, objRecordSet.Fields("TotalInvoiceAmtRoundOff_diff").Value)
        dblExchangeRate = IIf(IsDBNull(objRecordSet.Fields("Exchange_Rate").Value), 1, objRecordSet.Fields("Exchange_Rate").Value)
        dblTCStaxAmt = IIf(IsDBNull(objRecordSet.Fields("TCSTaxAmount").Value), 1, objRecordSet.Fields("TCSTaxAmount").Value)
        strCustCode = Trim(objRecordSet.Fields("Account_Code").Value)
        strCustRef = Trim(IIf(IsDBNull(objRecordSet.Fields("cust_ref").Value), "", objRecordSet.Fields("cust_ref").Value))
        If Not IsDBNull(objRecordSet.Fields("cust_ref").Value) Then
            blnExciseExumpted = objRecordSet.Fields("ExciseExumpted").Value
        Else
            blnExciseExumpted = False
        End If

        strCreditTermsID = Trim(IIf(IsDBNull(objRecordSet.Fields("payment_terms").Value), "", objRecordSet.Fields("payment_terms").Value))
        mstrCreditTermId = strCreditTermsID
        Dim objCreditTerms As New prj_CreditTerm.clsCR_Term_Resolver
        If UCase(mstrInvoiceType) <> "SMP" Then 'if invoice type is not sample sales then
            'Retreiving the customer gl, sl and credit term id
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If UCase(Trim(mstrInvoiceType)) = "REJ" Then
                objTmpRecordset.Open("SELECT ISNULL(SUM(Basic_Amount),0) AS Basic_Amt FROM sales_dtl WHERE doc_no =" & Ctlinvoice & " and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If Not objTmpRecordset.EOF Then
                    dblBasicAmount = objTmpRecordset.Fields("Basic_Amt").Value
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If (UCase(Trim(mstrInvoiceType)) = "REJ" And strCustRef <> "") Then 'In case of non line rejections Basic posting is not done
                    dblInvoiceAmt = dblInvoiceAmt - dblBasicAmount
                End If
                dblBasicAmount = 0
                objTmpRecordset.Open("SELECT GL_AccountID, Ven_slCode, CrTrm_Termid FROM Pur_VendorMaster where Prty_PartyID='" & strCustCode & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            Else
                objTmpRecordset.Open("SELECT Cst_ArCode, Cst_slCode, Cst_CreditTerm FROM Sal_CustomerMaster where Prty_PartyID='" & strCustCode & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            End If
            If objTmpRecordset.EOF Then
                If UCase(Trim(mstrInvoiceType)) = "REJ" Then
                    MsgBox("Vendor details not found", MsgBoxStyle.Information, "empower")
                Else
                    MsgBox("Customer details not found", MsgBoxStyle.Information, "empower")
                End If
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            If UCase(Trim(mstrInvoiceType)) = "REJ" Then
                strCustomerGL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("GL_AccountID").Value), "", objTmpRecordset.Fields("GL_AccountID").Value))
                strCustomerSL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Ven_slCode").Value), "", objTmpRecordset.Fields("Ven_slCode").Value))
                strCreditTermsID = Trim(IIf(IsDBNull(objTmpRecordset.Fields("CrTrm_Termid").Value), "", objTmpRecordset.Fields("CrTrm_Termid").Value))
            Else
                strCustomerGL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_ArCode").Value), "", objTmpRecordset.Fields("Cst_ArCode").Value))
                strCustomerSL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_slCode").Value), "", objTmpRecordset.Fields("Cst_slCode").Value))
                If strCreditTermsID = "" Then
                    strCreditTermsID = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_CreditTerm").Value), "", objTmpRecordset.Fields("Cst_CreditTerm").Value))
                    mstrCreditTermId = strCreditTermsID
                End If
            End If
            If strCreditTermsID = "" Then
                MsgBox("Credit Terms not found", MsgBoxStyle.Information, "empower")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            strRetval = objCreditTerms.RetCR_Term_Dates("", "INV", strCreditTermsID, strInvoiceDate, gstrUNITID, "", "", gstrCONNECTIONSTRING)
            If CheckString(strRetval) = "Y" Then
                strRetval = Mid(strRetval, 3)
                varTmp = Split(strRetval, "»")
                strBasicDueDate = VB6.Format(varTmp(0), "dd-MMM-yyyy")
                strPaymentDueDate = VB6.Format(varTmp(1), "dd-MMM-yyyy")
                strExpectedDueDate = VB6.Format(varTmp(1), "dd-MMM-yyyy")
            Else
                MsgBox(CheckString(strRetval), MsgBoxStyle.Information, "empower")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
        Else 'if  the invoice type is sample sales then
            'Code add by Sourabh on 31 dec 2004 For Post Tax to Customer Account For FOC
            blnFOC = CBool(Find_Value("select foc_invoice=isnull(foc_invoice,0) from salesChallan_dtl where Location_Code='" & Trim(txtUnitCode) & "' and doc_no='" & Trim(Ctlinvoice) & "' and unit_code='" & gstrUNITID & "'"))
            If blnFOC = True Then
                If rsobject.State = ADODB.ObjectStateEnum.adStateOpen Then rsobject.Close()
                rsobject.Open("SELECT Cst_ArCode, Cst_slCode, Cst_CreditTerm FROM Sal_CustomerMaster where Prty_PartyID='" & strCustCode & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If rsobject.EOF Then
                    MsgBox("Customer details not found", MsgBoxStyle.Information, "empower")
                    CreateStringForAccounts = False
                    If rsobject.State = ADODB.ObjectStateEnum.adStateOpen Then
                        rsobject.Close()
                        rsobject = Nothing
                    End If
                    Exit Function
                End If
                strCustomerGL = Trim(IIf(IsDBNull(rsobject.Fields("Cst_ArCode").Value), "", rsobject.Fields("Cst_ArCode").Value))
                strCustomerSL = Trim(IIf(IsDBNull(rsobject.Fields("Cst_slCode").Value), "", rsobject.Fields("Cst_slCode").Value))
            Else
                strRetval = GetItemGLSL("", "Sample_Expences")
                If strRetval = "N" Then
                    CreateStringForAccounts = False
                    If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                        objRecordSet.Close()
                        objRecordSet = Nothing
                    End If
                    Exit Function
                End If
                varTmp = Split(strRetval, "»")
                strCustomerGL = varTmp(0)
                strCustomerSL = varTmp(1)
            End If
        End If
        mstrMasterString = ""
        mstrDetailString = ""
        'to round off Total invoice amount according to parameter
        Dim rsSalesParameter As New ADODB.Recordset
        Dim blnTotalInvoiceAmountRoundOff As Boolean
        Dim intTotalInvoiceAmountRoundOff As Short
        If rsSalesParameter.State = ADODB.ObjectStateEnum.adStateOpen Then rsSalesParameter.Close()
        ' Check for EOU flag)
        rsSalesParameter.Open("SELECT EOU_Flag,TotalInvoiceAmount_RoundOff, TotalInvoiceAmountRoundOff_Decimal FROM SALES_PARAMETER where unit_code='" & gstrUNITID & "'", mP_Connection)
        If Not rsSalesParameter.EOF Then
            blnTotalInvoiceAmountRoundOff = rsSalesParameter.Fields("TotalInvoiceAmount_RoundOff").Value
            intTotalInvoiceAmountRoundOff = rsSalesParameter.Fields("TotalInvoiceAmountRoundOff_Decimal").Value
            mblnEOUUnit = rsSalesParameter.Fields("EOU_Flag").Value
        End If
        If rsSalesParameter.State = ADODB.ObjectStateEnum.adStateOpen Then
            rsSalesParameter.Close()
            rsSalesParameter = Nothing
        End If
        'Ends here
        If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
            mstrMasterString = "I»" & strInvoiceNo & "»Dr»»" & strInvoiceDate & "»»»»»SAL»I»" & strInvoiceNo & "»" & strInvoiceDate & "»"
            If UCase(mstrInvoiceType) <> "SMP" Then
                mstrMasterString = mstrMasterString & Trim(strCustCode) & "»" & gstrUNITID & "»" & strCurrencyCode & "»"
            Else
                mstrMasterString = mstrMasterString & "»" & gstrUNITID & "»" & strCurrencyCode & "»"
            End If
            'IF Condition to round off Total invoice amount according to parameter
            If blnTotalInvoiceAmountRoundOff Then
                mstrMasterString = mstrMasterString & System.Math.Round(dblInvoiceAmt, 0) & "»" & System.Math.Round(dblInvoiceAmt * dblExchangeRate, 0) & "»" & dblExchangeRate & "»" & strCreditTermsID & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCustomerGL & "»" & strCustomerSL & "»" & mP_User & "»getdate()»»"
            Else
                mstrMasterString = mstrMasterString & System.Math.Round(dblInvoiceAmt, intTotalInvoiceAmountRoundOff) & "»" & System.Math.Round(System.Math.Round(dblInvoiceAmt, intTotalInvoiceAmountRoundOff) * dblExchangeRate, intTotalInvoiceAmountRoundOff) & "»" & dblExchangeRate & "»" & strCreditTermsID & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCustomerGL & "»" & strCustomerSL & "»" & mP_User & "»getdate()»»"
            End If
            'IF Condition Ends here
        Else
            'IF Condition to round off Total invoice amount according to parameter
            If blnTotalInvoiceAmountRoundOff Then
                'mstrMasterString = "M»»" & VB6.Format(GetServerDate, "dd/MMM/yyyy") & "»0»»" & gstrUNITID & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & strInvoiceDate & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt) & "»0»»»Rej. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»DR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AP»»»»0»»¦"
                mstrMasterString = "M»»" & VB6.Format(GetServerDate, "dd/MMM/yyyy") & "»0»»" & gstrUNITID & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & strInvoiceDate & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt, intTotalInvoiceAmountRoundOff) & "»0»»»Rej. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»DR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AP»»»»0»»¦"
            Else
                mstrMasterString = "M»»" & VB6.Format(GetServerDate, "dd/MMM/yyyy") & "»0»»" & gstrUNITID & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & strInvoiceDate & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt, intTotalInvoiceAmountRoundOff) & "»0»»»Rej. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»DR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AP»»»»0»»¦"
            End If
            'IF Condition Ends here
        End If
        iCtr = 1
        'CST/LST/SRT/VAT Posting
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("SalesTax_Type").Value), "", objRecordSet.Fields("SalesTax_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SalesTax_Type").Value), "", objRecordSet.Fields("SalesTax_Type").Value)) & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("Tax type not found", MsgBoxStyle.Information, "empower")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If strTaxType = "LST" Or strTaxType = "CST" Or strTaxType = "SRT" Or strTaxType = "VAT" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Sales_Tax_Amount").Value), 0, objRecordSet.Fields("Sales_Tax_Amount").Value)
                dblBaseCurrencyAmount = dblTaxAmt
                dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("SalesTax_Per").Value), 0, objRecordSet.Fields("SalesTax_Per").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetval = GetTaxGlSl(strTaxType)
                    If strRetval = "N" Then
                        MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, "empower")
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetval, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»CST/LST/VAT for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    End If
                    iCtr = iCtr + 1
                End If
            End If
        End If
        ' Add to Post the State Developement Tax (04-May-2005)
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("SDTax_Type").Value), "", objRecordSet.Fields("SDTax_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SalesTax_Type").Value), "", objRecordSet.Fields("SDTax_Type").Value)) & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("Tax type not found", MsgBoxStyle.Information, "eMPro")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If strTaxType = "SDT" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("SDTax_Amount").Value), 0, objRecordSet.Fields("SDTax_Amount").Value)
                dblBaseCurrencyAmount = dblTaxAmt
                dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("SDTax_Per").Value), 0, objRecordSet.Fields("SDTax_Per").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetval = GetTaxGlSl(strTaxType)
                    If strRetval = "N" Then
                        MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, "eMPro")
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetval, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    If UCase(Trim(cmbInvType)) <> "REJECTION" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»SDT for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    End If
                    iCtr = iCtr + 1
                End If
            End If
        End If
        'ECS Posting
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("ECESS_Type").Value), "", objRecordSet.Fields("ECESS_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("ECESS_Type").Value), "", objRecordSet.Fields("ECESS_Type").Value)) & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("Tax type not found", MsgBoxStyle.Information, "empower")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If strTaxType = "ECS" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("ECESS_Amount").Value), 0, objRecordSet.Fields("ECESS_Amount").Value)
                dblBaseCurrencyAmount = dblTaxAmt
                dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("ECESS_Per").Value), 0, objRecordSet.Fields("ECESS_Per").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetval = GetTaxGlSl(strTaxType)
                    If strRetval = "N" Then
                        MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, "empower")
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetval, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»ECS for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    End If
                    iCtr = iCtr + 1
                End If
            End If
        End If
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("SECESS_Type").Value), "", objRecordSet.Fields("SECESS_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SECESS_Type").Value), "", objRecordSet.Fields("SECESS_Type").Value)) & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("Tax type not found", MsgBoxStyle.Information, "empower")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If strTaxType = "ECSSH" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("SECESS_Amount").Value), 0, objRecordSet.Fields("SECESS_Amount").Value)
                dblBaseCurrencyAmount = dblTaxAmt
                dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("SECESS_Per").Value), 0, objRecordSet.Fields("SECESS_Per").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetval = GetTaxGlSl(strTaxType)
                    If strRetval = "N" Then
                        MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, "empower")
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetval, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»ECSSH for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    End If
                    iCtr = iCtr + 1
                End If
            End If
        End If
        If mblnEOUUnit = True Then
            'Posting of ECS on CVD
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If Trim(IIf(IsDBNull(objRecordSet.Fields("CVDCESS_Type").Value), "", objRecordSet.Fields("CVDCESS_Type").Value)) <> "" Then
                objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("CVDCESS_Type").Value), "", objRecordSet.Fields("CVDCESS_Type").Value)) & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If Not objTmpRecordset.EOF Then
                    strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                Else
                    MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                    CreateStringForAccounts = False
                    If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                    Exit Function
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If strTaxType = "ECSCV" Then
                    dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("CVDCESS_Amount").Value), 0, objRecordSet.Fields("CVDCESS_Amount").Value)
                    dblBaseCurrencyAmount = dblTaxAmt
                    dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("CVDCESS_Per").Value), 0, objRecordSet.Fields("CVDCESS_Per").Value)
                    If dblBaseCurrencyAmount > 0 Then
                        'initializing the tax gl and sl here
                        strRetval = GetTaxGlSl(strTaxType)
                        If strRetval = "N" Then
                            MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, "eMPro")
                            CreateStringForAccounts = False
                            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                            Exit Function
                        End If
                        varTmp = Split(strRetval, "»")
                        strTaxGL = varTmp(0)
                        strTaxSL = varTmp(1)
                        If UCase$(Trim(mstrInvoiceType)) <> "REJ" Then
                            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & _
                                    iCtr & "»TAX»" & strTaxType & "»0»" & _
                                    "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                        Else
                            mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & _
                                               dblTaxAmt & "»»ECSCV for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        End If
                        iCtr = iCtr + 1
                    End If
                End If
            End If
            ''---- Posting of S.ECS on CVD
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If Trim(IIf(IsDBNull(objRecordSet.Fields("CVDSECESS_Type").Value), "", objRecordSet.Fields("CVDSECESS_Type").Value)) <> "" Then
                objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("CVDSECESS_Type").Value), "", objRecordSet.Fields("CVDSECESS_Type").Value)) & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If Not objTmpRecordset.EOF Then
                    strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                Else
                    MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                    CreateStringForAccounts = False
                    If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                    Exit Function
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If strTaxType = "HCSCV" Then
                    dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("CVDSECESS_Amount").Value), 0, objRecordSet.Fields("CVDSECESS_Amount").Value)
                    dblBaseCurrencyAmount = dblTaxAmt
                    dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("CVDSECESS_Per").Value), 0, objRecordSet.Fields("CVDSECESS_Per").Value)
                    If dblBaseCurrencyAmount > 0 Then
                        'initializing the tax gl and sl here
                        strRetval = GetTaxGlSl(strTaxType)
                        If strRetval = "N" Then
                            MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                            CreateStringForAccounts = False
                            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                            Exit Function
                        End If
                        varTmp = Split(strRetval, "»")
                        strTaxGL = varTmp(0)
                        strTaxSL = varTmp(1)
                        If UCase$(Trim(mstrInvoiceType)) <> "REJ" Then
                            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & _
                                    iCtr & "»TAX»" & strTaxType & "»0»" & _
                                    "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                        Else
                            mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & _
                                               dblTaxAmt & "»»HCSCV for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        End If
                        iCtr = iCtr + 1
                    End If
                End If
            End If
            'Posting of Additional ECS on Total Duty
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If Trim(IIf(IsDBNull(objRecordSet.Fields("Ecess_TotalDuty_Type").Value), "", objRecordSet.Fields("Ecess_TotalDuty_Type").Value)) <> "" Then
                objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("Ecess_TotalDuty_Type").Value), "", objRecordSet.Fields("Ecess_TotalDuty_Type").Value)) & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If Not objTmpRecordset.EOF Then
                    strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                Else
                    MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                    CreateStringForAccounts = False
                    If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                    Exit Function
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If strTaxType = "ADECS" Then
                    dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Ecess_TotalDuty_Amount").Value), 0, objRecordSet.Fields("Ecess_TotalDuty_Amount").Value)
                    dblBaseCurrencyAmount = dblTaxAmt
                    dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("Ecess_TotalDuty_Per").Value), 0, objRecordSet.Fields("Ecess_TotalDuty_Per").Value)
                    If dblBaseCurrencyAmount > 0 Then
                        'initializing the tax gl and sl here
                        strRetval = GetTaxGlSl(strTaxType)
                        If strRetval = "N" Then
                            MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                            CreateStringForAccounts = False
                            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                            Exit Function
                        End If
                        varTmp = Split(strRetval, "»")
                        strTaxGL = varTmp(0)
                        strTaxSL = varTmp(1)
                        If UCase$(Trim(mstrInvoiceType)) <> "REJ" Then
                            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & _
                                    iCtr & "»TAX»" & strTaxType & "»0»" & _
                                    "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                        Else
                            mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & _
                                               dblTaxAmt & "»»ADECS for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        End If
                        iCtr = iCtr + 1
                    End If
                End If
            End If
            ''---- Posting of Additionla S.ECS on Total Duty
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If Trim(IIf(IsDBNull(objRecordSet.Fields("SEcess_TotalDuty_Type").Value), "", objRecordSet.Fields("SEcess_TotalDuty_Type").Value)) <> "" Then
                objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SEcess_TotalDuty_Type").Value), "", objRecordSet.Fields("SEcess_TotalDuty_Type").Value)) & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If Not objTmpRecordset.EOF Then
                    strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                Else
                    MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                    CreateStringForAccounts = False
                    If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                    Exit Function
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If strTaxType = "ADHCS" Then
                    dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("SEcess_TotalDuty_Amount").Value), 0, objRecordSet.Fields("SEcess_TotalDuty_Amount").Value)
                    dblBaseCurrencyAmount = dblTaxAmt
                    dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("SEcess_TotalDuty_Per").Value), 0, objRecordSet.Fields("SEcess_TotalDuty_Per").Value)
                    If dblBaseCurrencyAmount > 0 Then
                        'initializing the tax gl and sl here
                        strRetval = GetTaxGlSl(strTaxType)
                        If strRetval = "N" Then
                            MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                            CreateStringForAccounts = False
                            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                            Exit Function
                        End If
                        varTmp = Split(strRetval, "»")
                        strTaxGL = varTmp(0)
                        strTaxSL = varTmp(1)
                        If UCase$(Trim(mstrInvoiceType)) <> "REJ" Then
                            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & _
                                    iCtr & "»TAX»" & strTaxType & "»0»" & _
                                    "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                        Else
                            mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & _
                                               dblTaxAmt & "»»ADHCS for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        End If
                        iCtr = iCtr + 1
                    End If
                End If
            End If
        End If
        'Turn Over Tax Posting
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("TurnOverTaxType").Value), "", objRecordSet.Fields("TurnOverTaxType").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("TurnOverTaxType").Value), "", objRecordSet.Fields("TurnOverTaxType").Value)) & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("Tax type not found", MsgBoxStyle.Information, "empower")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If strTaxType = "TOVT" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("turnOver_amt").Value), 0, objRecordSet.Fields("turnOver_amt").Value)
                dblBaseCurrencyAmount = dblTaxAmt
                dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("TurnOverTax_per").Value), 0, objRecordSet.Fields("TurnOverTax_per").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetval = GetTaxGlSl(strTaxType)
                    If strRetval = "N" Then
                        MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, "empower")
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetval, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»TOVT for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    End If
                    iCtr = iCtr + 1
                End If
            End If
        End If
        'Service Tax Posting
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("ServiceTax_Type").Value), "", objRecordSet.Fields("ServiceTax_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("ServiceTax_Type").Value), "", objRecordSet.Fields("ServiceTax_Type").Value)) & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("SRT Tax type not found", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If strTaxType = "SRT" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("ServiceTax_Amount").Value), 0, objRecordSet.Fields("ServiceTax_Amount").Value)
                dblBaseCurrencyAmount = dblTaxAmt
                dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("ServiceTax_Per").Value), 0, objRecordSet.Fields("ServiceTax_Per").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetval = GetTaxGlSl(strTaxType)
                    If strRetval = "N" Then
                        MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetval, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»SRTAX for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    End If
                    iCtr = iCtr + 1
                End If
            End If
        End If
        'ECS on Sale Tax Posting
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("SRCESS_Type").Value), "", objRecordSet.Fields("SRCESS_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SRCESS_Type").Value), "", objRecordSet.Fields("SRCESS_Type").Value)) & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("ECSR Tax type not found", MsgBoxStyle.Information, "empower")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If strTaxType = "ECSR" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("SRCESS_Amount").Value), 0, objRecordSet.Fields("SRCESS_Amount").Value)
                dblBaseCurrencyAmount = dblTaxAmt
                dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("SRCESS_Per").Value), 0, objRecordSet.Fields("SRCESS_Per").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetval = GetTaxGlSl(strTaxType)
                    If strRetval = "N" Then
                        MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, "empower")
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetval, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»ECS for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    End If
                    iCtr = iCtr + 1
                End If
            End If
        End If
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("SRSECESS_Type").Value), "", objRecordSet.Fields("SRSECESS_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SRSECESS_Type").Value), "", objRecordSet.Fields("SRSECESS_Type").Value)) & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("HECSR Tax type not found", MsgBoxStyle.Information, ResolveResString(100))
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If strTaxType = "HECSR" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("SRSECESS_Amount").Value), 0, objRecordSet.Fields("SRSECESS_Amount").Value)
                dblBaseCurrencyAmount = dblTaxAmt
                dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("SRSECESS_Per").Value), 0, objRecordSet.Fields("SRSECESS_Per").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetval = GetTaxGlSl(strTaxType)
                    If strRetval = "N" Then
                        MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetval, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»SHECS for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    End If
                    iCtr = iCtr + 1
                End If
            End If
        End If
        'SST Posting
        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Surcharge_Sales_Tax_Amount").Value), 0, objRecordSet.Fields("Surcharge_Sales_Tax_Amount").Value)
        dblBaseCurrencyAmount = dblTaxAmt
        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("Surcharge_SalesTax_Per").Value), 0, objRecordSet.Fields("Surcharge_SalesTax_Per").Value)
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            strRetval = GetTaxGlSl("SST")
            If strRetval = "N" Then
                MsgBox("GL for ARTAX is not defined for SST", MsgBoxStyle.Information, "empower")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetval, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»SST»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Surcharge for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
            End If
            iCtr = iCtr + 1
        End If
        'Insurance Posting
        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Insurance").Value), 0, objRecordSet.Fields("Insurance").Value)
        dblBaseCurrencyAmount = dblTaxAmt
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            strRetval = GetTaxGlSl("INS")
            If strRetval = "N" Then
                MsgBox("GL for ARTAX is not defined for INS", MsgBoxStyle.Information, "empower")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetval, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»INS»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Insurance for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
            End If
            iCtr = iCtr + 1
        End If
        'Freight Posting
        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Frieght_Amount").Value), 0, objRecordSet.Fields("Frieght_Amount").Value)
        dblBaseCurrencyAmount = dblTaxAmt
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            strRetval = GetTaxGlSl("FRT")
            If strRetval = "N" Then
                MsgBox("GL for ARTAX is not defined for FRT", MsgBoxStyle.Information, "empower")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetval, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»FRT»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Freight for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
            End If
            iCtr = iCtr + 1
        End If
        '******************Discount Posting code added by nisha on 18/09/2003
        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Discount_Amount").Value), 0, objRecordSet.Fields("Discount_Amount").Value)
        dblBaseCurrencyAmount = dblTaxAmt
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            strRetval = GetItemGLSL("", "Discount_Interest")
            If strRetval = "N" Then
                MsgBox("GL For Purpose Code Discount_Interest is not defined. ", MsgBoxStyle.Information, "empower")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetval, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If System.Math.Abs(dblBaseCurrencyAmount) > 0 Then
                If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»»TAX»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & System.Math.Abs(dblBaseCurrencyAmount) & "»" & "Dr»»»»»»0»0»0»0»0" & "¦"
                Else
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»DR»" & System.Math.Abs(dblBaseCurrencyAmount) & "»Discount amount for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                End If
            End If
            iCtr = iCtr + 1
        End If
        '********************************** changes ends here by nisha on 18/09/2003
        '******************TCS Tax Posting code added by nisha on 26/02/2004
        If (UCase(Trim(mstrInvoiceType)) = "INV") And (UCase(Trim(mstrInvoiceSubType)) = "L") Then
            dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("TCSTaxAmount").Value), 0, objRecordSet.Fields("TCSTaxAmount").Value)
            dblBaseCurrencyAmount = dblTaxAmt
            If dblBaseCurrencyAmount > 0 Then
                'initializing the tax gl and sl here
                strRetval = GetTaxGlSl("TCS")
                If strRetval = "N" Then
                    MsgBox("GL For Purpose Code TCS Tax is not defined. ", MsgBoxStyle.Information, "empower")
                    CreateStringForAccounts = False
                    If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                        objRecordSet.Close()
                        objRecordSet = Nothing
                    End If
                    Exit Function
                End If
                varTmp = Split(strRetval, "»")
                strTaxGL = varTmp(0)
                strTaxSL = varTmp(1)
                If System.Math.Abs(dblBaseCurrencyAmount) > 0 Then
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»»TCS»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & System.Math.Abs(dblBaseCurrencyAmount) & "»" & "Cr»»»»»»0»0»0»0»0" & "¦"
                End If
                iCtr = iCtr + 1
            End If
        End If
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        If mblnMultipleSOAllowed = False Then
            objRecordSet.Open("SELECT sales_dtl.*, item_mst.GlGrp_code FROM sales_dtl, item_mst WHERE sales_dtl.Doc_No='" & Trim(Ctlinvoice) & "' and sales_dtl.Item_Code=item_mst.Item_Code and sales_dtl.unit_Code=item_mst.unit_code and sales_dtl.Location_Code='" & Trim(txtUnitCode) & "' and sales_dtl.unit_code='" & gstrUNITID & "'")
        Else
            objRecordSet.Open("SELECT isnull(sum(a.basic_amount),0) as Basic_Amount,isnull(sum(a.CustMtrl_Amount),0) as CustMtrl_Amount, isnull(sum(a.Excise_tax),0) as Excise_tax,isnull(sum(ItemPacking_Amount),0)  as ItemPacking_Amount,isnull(sum(others),0) as Others,a.item_code, b.GlGrp_code," & "isnull(sum(packing),0) as packing,pkg_amount FROM sales_dtl a, item_mst b WHERE a.Doc_No='" & Trim(Ctlinvoice) & "' and a.Item_Code=b.Item_Code and a.unit_code=b.unit_code and a.unit_code ='" & gstrUNITID & "' and a.Location_Code='" & Trim(txtUnitCode) & "'" & " group by a.item_code,b.GlGrp_code")
        End If
        If objRecordSet.EOF Then
            MsgBox("Item details not found.", MsgBoxStyle.Information, "empower")
            CreateStringForAccounts = False
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        Dim dblPacking_per As Double
        While Not objRecordSet.EOF
            strGlGroupId = Trim(IIf(IsDBNull(objRecordSet.Fields("GlGrp_code").Value), "", objRecordSet.Fields("GlGrp_code").Value))
            'Basic Amount Posting
            blnFOC = CBool(Find_Value("select foc_invoice from salesChallan_dtl where Location_Code='" & Trim(txtUnitCode) & "' and doc_no='" & Trim(Ctlinvoice) & "' and unit_code='" & gstrUNITID & "'"))
            If UCase(Trim(mstrInvoiceType)) = "SMP" And blnFOC Then
                'skip posting of basic if invoice is FOC Sample invoice
            ElseIf (UCase(Trim(mstrInvoiceType)) = "REJ" And strCustRef = "") Or UCase(Trim(mstrInvoiceType)) <> "REJ" Then  'In case of non line rejections Basic posting is not done
                dblBasicAmount = IIf(IsDBNull(objRecordSet.Fields("Basic_Amount").Value), 0, objRecordSet.Fields("Basic_Amount").Value)
                If mblnAddCustomerMaterial Then
                    dblBaseCurrencyAmount = dblBasicAmount + IIf(IsDBNull(objRecordSet.Fields("CustMtrl_Amount").Value), 0, objRecordSet.Fields("CustMtrl_Amount").Value)
                Else
                    dblBaseCurrencyAmount = dblBasicAmount
                End If
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the item gl and sl************************
                    strRetval = GetItemGLSL(strGlGroupId, mstrPurposeCode)
                    If strRetval = "N" Then
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetval, "»")
                    strItemGL = varTmp(0)
                    strItemSL = varTmp(1)
                    'initializing of item gl and sl ends here****************
                    'Posting the basic amount into cost centers, percentage wise
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                    objTmpRecordset.Open("SELECT * FROM invcc_dtl WHERE Invoice_Type='" & mstrInvoiceType & "' AND Sub_Type = '" & mstrInvoiceSubType & "' AND Location_Code ='" & Trim(txtUnitCode) & "' AND ccM_cc_Percentage > 0 and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                    If Not objTmpRecordset.EOF Then
                        While Not objTmpRecordset.EOF
                            dblCCShare = (dblBaseCurrencyAmount / 100) * objTmpRecordset.Fields("ccM_cc_Percentage").Value
                            If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»ITM»SAL»" & iCtr & "»" & Trim(objRecordSet.Fields("item_code").Value) & "»" & strGlGroupId & "»0»" & strItemGL & "»" & strItemSL & "»" & dblCCShare & "»Cr»»" & Trim(objTmpRecordset.Fields("ccM_ccCode").Value) & "»»»»0»0»0»0»0¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»" & Trim(objTmpRecordset.Fields("ccM_ccCode").Value) & "»»»CR»" & dblCCShare & "»»Basic for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            End If
                            objTmpRecordset.MoveNext()
                            iCtr = iCtr + 1
                        End While
                    Else
                        If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»ITM»SAL»" & iCtr & "»" & Trim(objRecordSet.Fields("item_code").Value) & "»" & strGlGroupId & "»0»" & strItemGL & "»" & strItemSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                        Else
                            mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»»»»CR»" & dblBaseCurrencyAmount & "»»Basic for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        End If
                        iCtr = iCtr + 1
                    End If
                    '*********************************************************
                End If
            End If
            'EXC Duty Posting
            'IF Condition added by nisha for Excise Exumption on 10/07/2003
            If blnExciseExumpted = False Then
                If mblnEOUUnit = False Then
                    dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Excise_Tax").Value), 0, objRecordSet.Fields("Excise_Tax").Value)
                Else
                    dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("TotalExciseAmount").Value), 0, objRecordSet.Fields("TotalExciseAmount").Value)
                End If
                If mblnExciseRoundOFFFlag Then dblTaxAmt = System.Math.Round(dblTaxAmt, 0)
                dblBaseCurrencyAmount = dblTaxAmt
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    rsFULLExciseAmount.GetResult("Select Sum(isnull(TotalExciseAmount,0)) as TotalExciseAmount from Sales_dtl where Doc_no =" & Ctlinvoice & " and unit_code='" & gstrUNITID & "'")
                    dblFullExciseAmount = rsFULLExciseAmount.GetValue("TotalExciseAmount")
                    If CheckExcPriority() = 0 Then
                        If blnMsgBox = False Then
                            If MsgBox("No Excise Priority is Defined Would like to Post in ARTax ?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "empower") = MsgBoxResult.Yes Then
                                blnMsgBox = True
                            Else
                                CreateStringForAccounts = False
                                Exit Function
                            End If
                        End If
                        strRetval = GetTaxGlSl("EXC")
                        If strRetval = "N" Then
                            MsgBox("GL for ARTAX is not defined for EXC", MsgBoxStyle.Information, "empower")
                            CreateStringForAccounts = False
                            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                                objRecordSet.Close()
                                objRecordSet = Nothing
                            End If
                            Exit Function
                        End If
                        varTmp = Split(strRetval, "»")
                        strTaxGL = varTmp(0)
                        strTaxSL = varTmp(1)
                        mstrExcisePriorityUpdationString = ""
                    Else
                        arrstrExcPriority = ReturnGLSLAccExcPriority(1, dblFullExciseAmount)
                        If Len(Trim(arrstrExcPriority(0))) = 0 Then
                            arrstrExcPriority = ReturnGLSLAccExcPriority(2, dblFullExciseAmount)
                            If Len(Trim(arrstrExcPriority(1))) = 0 Then
                                arrstrExcPriority = ReturnGLSLAccExcPriority(3, dblFullExciseAmount)
                                If Len(Trim(arrstrExcPriority(1))) = 0 Then
                                    If blnMsgBox = False Then
                                        If MsgBox("Excise amount To be Posted is Greater then available in All the Three Priorities Defined. would You like to Post in ARTax ?", MsgBoxStyle.YesNo, "empower") = MsgBoxResult.Yes Then
                                            blnMsgBox = True
                                        Else
                                            CreateStringForAccounts = False
                                            Exit Function
                                        End If
                                    End If
                                    strRetval = GetTaxGlSl("EXC")
                                    If strRetval = "N" Then
                                        MsgBox("GL for ARTAX is not defined for EXC", MsgBoxStyle.Information, "empower")
                                        CreateStringForAccounts = False
                                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                                            objRecordSet.Close()
                                            objRecordSet = Nothing
                                        End If
                                        Exit Function
                                    End If
                                    varTmp = Split(strRetval, "»")
                                    'To be Posted agianest ARtax 3
                                    strTaxGL = varTmp(0)
                                    strTaxSL = varTmp(1)
                                    mstrExcisePriorityUpdationString = ""
                                Else
                                    'To be Posted agianest Priority 3
                                    strTaxGL = arrstrExcPriority(0)
                                    strTaxSL = arrstrExcPriority(1)
                                    mstrExcisePriorityUpdationString = arrstrExcPriority(2)
                                End If
                            Else
                                'To be Posted agianest Priority 2
                                strTaxGL = arrstrExcPriority(0)
                                strTaxSL = arrstrExcPriority(1)
                                mstrExcisePriorityUpdationString = arrstrExcPriority(2)
                            End If
                        Else
                            'To be Posted agianest Priority 1
                            strTaxGL = arrstrExcPriority(0)
                            strTaxSL = arrstrExcPriority(1)
                            mstrExcisePriorityUpdationString = arrstrExcPriority(2)
                        End If
                    End If
                    If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»EXC»0»" & Trim(objRecordSet.Fields("item_code").Value) & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Excise for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    End If
                    iCtr = iCtr + 1
                End If
            End If
            'Changes Ends Here 10/07/2003
            'Added by Manoj on 10 Jul 2007 for Issue Id 19992(Not added By Ashutosh on 28-07-2006,Issue Id:18350)
            '*******************************************************************************
            'Packing Value Posting
            '*******************************************************************************
            'changes on 25 july 2017
            dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("pkg_Amount").Value), 0, objRecordSet.Fields("Pkg_Amount").Value)
            dblBaseCurrencyAmount = dblTaxAmt
            If dblBaseCurrencyAmount > 0 Then
                strRetval = GetTaxGlSl("PKT")
                If strRetval = "N" Then
                    MsgBox("GL for ARTAX is not defined for PACKING", MsgBoxStyle.Information, "empower")
                    CreateStringForAccounts = False
                    If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                        objRecordSet.Close()
                        objRecordSet = Nothing
                    End If
                    Exit Function
                End If
                varTmp = Split(strRetval, "»")
                strTaxGL = varTmp(0)
                strTaxSL = varTmp(1)
                If UCase(Trim(cmbInvType)) = "REJECTION" Then
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Packing Charges for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                Else
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»PKT»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                End If
                iCtr = iCtr + 1
            End If

            'changes on 25 july 2017

            dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("ItemPacking_Amount").Value), 0, objRecordSet.Fields("ItemPacking_Amount").Value)
            dblPacking_per = IIf(IsDBNull(objRecordSet.Fields("Packing").Value), 0, objRecordSet.Fields("Packing").Value)
            dblBaseCurrencyAmount = dblTaxAmt
            If dblBaseCurrencyAmount > 0 Then
                strRetval = GetTaxGlSl("PKT")
                If strRetval = "N" Then
                    MsgBox("GL for ARTAX is not defined for PACKING", MsgBoxStyle.Information, "eMPro")
                    CreateStringForAccounts = False
                    If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                        objRecordSet.Close()
                        objRecordSet = Nothing
                    End If
                    Exit Function
                End If
                varTmp = Split(strRetval, "»")
                strTaxGL = varTmp(0)
                strTaxSL = varTmp(1)
                If dblBaseCurrencyAmount > 0 Then
                    If UCase(Trim(mstrInvoiceType)) <> "REJECTION" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»PKT»0»" & Trim(objRecordSet.Fields("item_code").Value) & "»»" & dblPacking_per & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Packing Charges for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    End If
                    iCtr = iCtr + 1
                End If
            End If
            If mblnEOUUnit = True Then
                'Posting of CVD Excise
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If Trim(IIf(IsDBNull(objRecordSet.Fields("CVD_type").Value), "", objRecordSet.Fields("CVD_type").Value)) <> "" Then
                    objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("CVD_type").Value), "", objRecordSet.Fields("CVD_type").Value)) & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                    If Not objTmpRecordset.EOF Then
                        strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                    Else
                        MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                        Exit Function
                    End If
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                    If strTaxType = "CVD" Then
                        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("CVD_Amount").Value), 0, objRecordSet.Fields("CVD_Amount").Value)
                        dblBaseCurrencyAmount = dblTaxAmt
                        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("CVD_per").Value), 0, objRecordSet.Fields("CVD_per").Value)
                        If dblBaseCurrencyAmount > 0 Then
                            'initializing the tax gl and sl here
                            strRetval = GetTaxGlSl(strTaxType)
                            If strRetval = "N" Then
                                MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                                CreateStringForAccounts = False
                                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                                Exit Function
                            End If
                            varTmp = Split(strRetval, "»")
                            strTaxGL = varTmp(0)
                            strTaxSL = varTmp(1)
                            If UCase$(Trim(mstrInvoiceType)) <> "REJ" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & _
                                iCtr & "»TAX»CVD»0»" & Trim(objRecordSet.Fields("item_code").Value) & _
                               "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & _
                                         dblTaxAmt & "»»CVD for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            End If
                            iCtr = iCtr + 1
                        End If
                    End If
                End If
            End If
            If mblnEOUUnit = False And UCase$(Trim(mstrInvoiceType)) = "TRF" Then
                'Posting of SAD Tax
                If objTmpRecordset.State = 1 Then objTmpRecordset.Close()
                If Trim(IIf(IsDBNull(objRecordSet.Fields("SAD_type").Value), "", objRecordSet.Fields("SAD_type").Value)) <> "" Then
                    objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SAD_type").Value), "", objRecordSet.Fields("SAD_type").Value)) & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                    If Not objTmpRecordset.EOF Then
                        strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                    Else
                        MsgBox("Tax type not found", MsgBoxStyle.Information, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = 1 Then objRecordSet.Close() : objRecordSet = Nothing
                        If objTmpRecordset.State = 1 Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                        Exit Function
                    End If
                    If objTmpRecordset.State = 1 Then objTmpRecordset.Close()
                    If strTaxType = "SAD" Then
                        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("SVD_Amount").Value), 0, objRecordSet.Fields("SVD_Amount").Value)
                        dblBaseCurrencyAmount = dblTaxAmt
                        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("SVD_per").Value), 0, objRecordSet.Fields("SVD_per").Value)
                        If dblBaseCurrencyAmount > 0 Then
                            'initializing the tax gl and sl here
                            strRetval = GetTaxGlSl(strTaxType)
                            If strRetval = "N" Then
                                MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                                CreateStringForAccounts = False
                                If objRecordSet.State = 1 Then objRecordSet.Close() : objRecordSet = Nothing
                                Exit Function
                            End If
                            varTmp = Split(strRetval, "»")
                            strTaxGL = varTmp(0)
                            strTaxSL = varTmp(1)
                            If UCase$(Trim(mstrInvoiceType)) <> "REJ" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & _
                                iCtr & "»TAX»SAD»0»" & Trim(objRecordSet.Fields("item_code").Value) & _
                               "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & _
                                         dblTaxAmt & "»»SVD for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            End If
                            iCtr = iCtr + 1
                        End If
                    End If
                End If
            End If
            'Others Posting
            dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Others").Value), 0, objRecordSet.Fields("Others").Value)
            dblBaseCurrencyAmount = dblTaxAmt
            'initialize the tax gl and sl here
            If dblBaseCurrencyAmount > 0 Then
                If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»OTH»0»" & Trim(objRecordSet.Fields("item_code").Value) & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                Else
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Other Charges for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                End If
                iCtr = iCtr + 1
            End If
            'GST CHANGES 101188073
            If gblnGSTUnit Then
                'Posting of CGST
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If Trim(IIf(IsDBNull(objRecordSet.Fields("CGSTTXRT_TYPE").Value), "", objRecordSet.Fields("CGSTTXRT_TYPE").Value)) <> "" Then
                    objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("CGSTTXRT_TYPE").Value), "", objRecordSet.Fields("CGSTTXRT_TYPE").Value)) & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                    If Not objTmpRecordset.EOF Then
                        strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                    Else
                        MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                        Exit Function
                    End If
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                    If strTaxType = "CGST" Then
                        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("CGST_AMT").Value), 0, objRecordSet.Fields("CGST_AMT").Value)
                        dblBaseCurrencyAmount = dblTaxAmt
                        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("CGST_PERCENT").Value), 0, objRecordSet.Fields("CGST_PERCENT").Value)
                        If dblBaseCurrencyAmount > 0 Then
                            'initializing the tax gl and sl here
                            strRetval = GetTaxGlSl(strTaxType)
                            If strRetval = "N" Then
                                MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                                CreateStringForAccounts = False
                                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                                Exit Function
                            End If
                            varTmp = Split(strRetval, "»")
                            strTaxGL = varTmp(0)
                            strTaxSL = varTmp(1)
                            If UCase$(Trim(mstrInvoiceType)) <> "REJ" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & _
                                iCtr & "»TAX»CGST»0»" & Trim(objRecordSet.Fields("item_code").Value) & _
                               "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & _
                                         dblTaxAmt & "»»CGST for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            End If
                            iCtr = iCtr + 1
                        End If
                    End If
                End If
                'Posting of SGST
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If Trim(IIf(IsDBNull(objRecordSet.Fields("SGSTTXRT_TYPE").Value), "", objRecordSet.Fields("SGSTTXRT_TYPE").Value)) <> "" Then
                    objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SGSTTXRT_TYPE").Value), "", objRecordSet.Fields("SGSTTXRT_TYPE").Value)) & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                    If Not objTmpRecordset.EOF Then
                        strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                    Else
                        MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                        Exit Function
                    End If
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                    If strTaxType = "SGST" Then
                        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("SGST_AMT").Value), 0, objRecordSet.Fields("SGST_AMT").Value)
                        dblBaseCurrencyAmount = dblTaxAmt
                        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("SGST_PERCENT").Value), 0, objRecordSet.Fields("SGST_PERCENT").Value)
                        If dblBaseCurrencyAmount > 0 Then
                            'initializing the tax gl and sl here
                            strRetval = GetTaxGlSl(strTaxType)
                            If strRetval = "N" Then
                                MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                                CreateStringForAccounts = False
                                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                                Exit Function
                            End If
                            varTmp = Split(strRetval, "»")
                            strTaxGL = varTmp(0)
                            strTaxSL = varTmp(1)
                            If UCase$(Trim(mstrInvoiceType)) <> "REJ" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & _
                                iCtr & "»TAX»SGST»0»" & Trim(objRecordSet.Fields("item_code").Value) & _
                               "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & _
                                         dblTaxAmt & "»»SGST for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            End If
                            iCtr = iCtr + 1
                        End If
                    End If
                End If
                'Posting of UTGST
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If Trim(IIf(IsDBNull(objRecordSet.Fields("UTGSTTXRT_TYPE").Value), "", objRecordSet.Fields("UTGSTTXRT_TYPE").Value)) <> "" Then
                    objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("UTGSTTXRT_TYPE").Value), "", objRecordSet.Fields("UTGSTTXRT_TYPE").Value)) & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                    If Not objTmpRecordset.EOF Then
                        strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                    Else
                        MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                        Exit Function
                    End If
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                    If strTaxType = "UTGST" Then
                        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("UTGST_AMT").Value), 0, objRecordSet.Fields("UTGST_AMT").Value)
                        dblBaseCurrencyAmount = dblTaxAmt
                        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("UTGST_PERCENT").Value), 0, objRecordSet.Fields("UTGST_PERCENT").Value)
                        If dblBaseCurrencyAmount > 0 Then
                            'initializing the tax gl and sl here
                            strRetval = GetTaxGlSl(strTaxType)
                            If strRetval = "N" Then
                                MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                                CreateStringForAccounts = False
                                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                                Exit Function
                            End If
                            varTmp = Split(strRetval, "»")
                            strTaxGL = varTmp(0)
                            strTaxSL = varTmp(1)
                            If UCase$(Trim(mstrInvoiceType)) <> "REJ" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & _
                                iCtr & "»TAX»UTGST»0»" & Trim(objRecordSet.Fields("item_code").Value) & _
                               "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & _
                                         dblTaxAmt & "»»UTGST for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            End If
                            iCtr = iCtr + 1
                        End If
                    End If
                End If
                'Posting of IGST
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If Trim(IIf(IsDBNull(objRecordSet.Fields("IGSTTXRT_TYPE").Value), "", objRecordSet.Fields("IGSTTXRT_TYPE").Value)) <> "" Then
                    objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("IGSTTXRT_TYPE").Value), "", objRecordSet.Fields("IGSTTXRT_TYPE").Value)) & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                    If Not objTmpRecordset.EOF Then
                        strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                    Else
                        MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                        Exit Function
                    End If
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                    If strTaxType = "IGST" Then
                        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("IGST_AMT").Value), 0, objRecordSet.Fields("IGST_AMT").Value)
                        dblBaseCurrencyAmount = dblTaxAmt
                        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("IGST_PERCENT").Value), 0, objRecordSet.Fields("IGST_PERCENT").Value)
                        If dblBaseCurrencyAmount > 0 Then
                            'initializing the tax gl and sl here
                            strRetval = GetTaxGlSl(strTaxType)
                            If strRetval = "N" Then
                                MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                                CreateStringForAccounts = False
                                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                                Exit Function
                            End If
                            varTmp = Split(strRetval, "»")
                            strTaxGL = varTmp(0)
                            strTaxSL = varTmp(1)
                            If UCase$(Trim(mstrInvoiceType)) <> "REJ" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & _
                                iCtr & "»TAX»IGST»0»" & Trim(objRecordSet.Fields("item_code").Value) & _
                               "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & _
                                         dblTaxAmt & "»»IGST for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            End If
                            iCtr = iCtr + 1
                        End If
                    End If
                End If
                'Posting of COMPENSATION CESS
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If Trim(IIf(IsDBNull(objRecordSet.Fields("COMPENSATION_CESS_TYPE").Value), "", objRecordSet.Fields("COMPENSATION_CESS_TYPE").Value)) <> "" Then
                    objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("COMPENSATION_CESS_TYPE").Value), "", objRecordSet.Fields("COMPENSATION_CESS_TYPE").Value)) & "' and unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                    If Not objTmpRecordset.EOF Then
                        strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                    Else
                        MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                        Exit Function
                    End If
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                    If strTaxType = "GSTEC" Then
                        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("COMPENSATION_CESS_AMT").Value), 0, objRecordSet.Fields("COMPENSATION_CESS_AMT").Value)
                        dblBaseCurrencyAmount = dblTaxAmt
                        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("COMPENSATION_CESS_PERCENT").Value), 0, objRecordSet.Fields("COMPENSATION_CESS_PERCENT").Value)
                        If dblBaseCurrencyAmount > 0 Then
                            'initializing the tax gl and sl here
                            strRetval = GetTaxGlSl(strTaxType)
                            If strRetval = "N" Then
                                MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                                CreateStringForAccounts = False
                                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                                Exit Function
                            End If
                            varTmp = Split(strRetval, "»")
                            strTaxGL = varTmp(0)
                            strTaxSL = varTmp(1)
                            If UCase$(Trim(mstrInvoiceType)) <> "REJ" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & _
                                iCtr & "»TAX»CCESS»0»" & Trim(objRecordSet.Fields("item_code").Value) & _
                               "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & _
                                         dblTaxAmt & "»»CCESS for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            End If
                            iCtr = iCtr + 1
                        End If
                    End If
                End If
            End If
            'GST CHANGES 101188073
            objRecordSet.MoveNext()
        End While
        'Posting of rounded off amount
        strRetval = GetItemGLSL("", "Rounded_Amt")
        If strRetval = "N" Then
            CreateStringForAccounts = False
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        varTmp = Split(strRetval, "»")
        strItemGL = varTmp(0)
        strItemSL = varTmp(1)
        dblBaseCurrencyAmount = dblInvoiceAmtRoundOff_diff
        dblBaseCurrencyAmount = System.Math.Round(dblBaseCurrencyAmount, 4)
        If System.Math.Abs(dblBaseCurrencyAmount) > 0 Then
            If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»»RND»0»" & "»»0»" & strItemGL & "»" & strItemSL & "»" & System.Math.Abs(dblBaseCurrencyAmount) & "»"
                If dblBaseCurrencyAmount < 0 Then
                    mstrDetailString = mstrDetailString & "Cr»»»»»»0»0»0»0»0" & "¦"
                Else
                    mstrDetailString = mstrDetailString & "Dr»»»»»»0»0»0»0»0" & "¦"
                End If
            Else
                If dblBaseCurrencyAmount < 0 Then
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»»»»CR»" & System.Math.Abs(dblBaseCurrencyAmount) & "»Rounding off amount for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                Else
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»»»»DR»" & System.Math.Abs(dblBaseCurrencyAmount) & "»Rounding off amount for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                End If
            End If
        End If
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
        CreateStringForAccounts = True
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        CreateStringForAccounts = False
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
    End Function
    Public Sub disableControls()
        On Error GoTo Errorhandler
        Me.cmdLockInvoice.Revert()
        Me.cmdLockInvoice.Caption(0) = "POST"
        Me.cmbSearch.Items.Clear() : Me.txtSearch.Text = "" : Me.cmbSort.Items.Clear()
        Me.optunCheckAll.Checked = True
        Exit Sub
Errorhandler:
        If Err.Number = 5 Then Resume Next
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub CheckMultipleSOAllowed(ByVal pInvType As String, ByVal pInvSubType As String)
        '-----------------------------------------------------------------------------------
        'Created By      : Manoj Kr.Vaish
        'Issue ID        : 19992
        'Creation Date   : 27 JUNE 2007
        'Procedure       : To Check MultipleSOAllowed for Any Invoice Type
        '-----------------------------------------------------------------------------------
        Dim rsCheckSo As ClsResultSetDB
        Dim strsql As String
        On Error GoTo ErrHandler
        rsCheckSo = New ClsResultSetDB
        strsql = "select isnull(sorequired,0) as SORequired,isnull(MultipleSOAllowed,0) as MultipleSOAllowed from saleconf where description='" & Trim(pInvType) & "' and sub_type_description='" & Trim(pInvSubType) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate()) and unit_code='" & gstrUNITID & "'"
        rsCheckSo.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsCheckSo.GetNoRows > 0 Then
            mblnMultipleSOAllowed = rsCheckSo.GetValue("MultipleSOAllowed")
            mblnSORequired = rsCheckSo.GetValue("SORequired")
        End If
        rsCheckSo.ResultSetClose()
        rsCheckSo = Nothing
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
        '------------------------------------------------------------------------------------------
    End Sub
    Private Sub dtFromDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtFromDate.ValueChanged
        If Not blnDate Then
            Me.spGrid.MaxRows = 0
        End If
    End Sub
    Private Sub dtToDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtToDate.ValueChanged
        If Not blnDate Then
            Me.spGrid.MaxRows = 0
        End If
    End Sub
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLPMKTTRN0040.HTM")
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
End Class