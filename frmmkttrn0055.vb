Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0055
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------------------------------------------------------------
	'Copyright(c)               -   MIND
	'Form Name (Physical Name)  -   frmMKTTRN0055.frm
	'Created by                 -   NEERAJ YADAV
	'Modified By                -   NEERAJ YADAV
	'Created Date               -   12-11-2007
	'----------------------------------------------------------------------------------------------------------
	'------------------------------------------------------------------------
	' Revised By                 -   neeraj yadav
	' Revision Date              -   20 dec 2007
	' Issue Id                   -   21870
	' Revision History           -   some more fields added on this forms(driver name,security personnel)
	'------------------------------------------------------------------------
    'Modified By JY on 12-May-2011
    '   Modified to support MultiUnit functionality
    '****************************************************
    Dim mintIndex As Short
    Private Enum grddetail
        check = 1
        Invoice_No = 2
        Item_Code = 3
        invoice_qty = 4
        picked_qty = 5
        barcode_track = 6
    End Enum
    Private Enum grdheader
        CheckBox = 1
        Invoice_No = 2
        Invoice_Date = 3
        ShowDetail = 4
    End Enum
    Private Sub CmdDocNo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmddocno.Click

        On Error GoTo errHandler
        Dim strHelp As String
        Dim strgatepassno() As String
        Dim strSQL As String
        ctlcmd.Enabled(0) = True
        ctlcmd.Enabled(1) = False
        ctlcmd.Enabled(4) = False
        ctlcmd.Enabled(3) = False
        If Len(Trim(txtdoc_no.Text)) = 0 Then
            strHelp = "set dateformat 'dmy' select GatePassNum," & DateColumnNameInShowList("Gatepasscreateddt") & " as Gatepasscreateddt from bar_gatepass_hdr where UNIT_CODE='" & gstrUNITID & "' and gatepasscreateddt between '" & getDateForDB(dtpfromdate.Value) & "' and '" & getDateForDB(dtptodate.Value) & "'"
            strgatepassno = ctlEMPHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, "Gate Pass No")
        Else
            strHelp = "set dateformat 'dmy' select GatePassNum," & DateColumnNameInShowList("Gatepasscreateddt") & " as Gatepasscreateddt  from bar_gatepass_hdr where UNIT_CODE='" & gstrUNITID & "' and gatepassnum=" & Trim(txtdoc_no.Text) & " and gatepasscreateddt between '" & getDateForDB(dtpfromdate.Value) & "' and '" & getDateForDB(dtptodate.Value) & "'"
            strgatepassno = ctlEMPHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, "Gate Pass No")
        End If
        If Not (UBound(strgatepassno) = -1) Then
            If Len(strgatepassno(0)) >= 1 And strgatepassno(0) = "0" Then
                MsgBox("Record Not Found", MsgBoxStyle.Information, ResolveResString(100))
                Call ClearControls()
                Exit Sub
            Else
                txtdoc_no.Text = strgatepassno(0)
                FillDataInViewMode()
                Call setcontrolenableproperty((ctlcmd.Mode))
            End If
            'ctlcmd.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
        End If
        Exit Sub
errHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdShow_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdshow.Click
        On Error GoTo errHandler
        fspdetail.MaxRows = 1
        Call ClearControls()
        txtdoc_no.Enabled = False : txtdoc_no.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmddocno.Enabled = False
        Call fillData()
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub dtpFromDate_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpfromdate.Leave
        On Error GoTo errHandler
        Dim objdate As New ClsResultSetDB
        Dim dtServerDate As Date
        Dim dtFromDate As Short
        dtServerDate = GetServerDate()
        objdate.GetResult("select isnull(gatepassfromdate,1) as gatepassfromdate from BARCODE_Parameters where UNIT_CODE='" & gstrUNITID & "'")
        If Not objdate.EOFRecord Then
            dtFromDate = objdate.GetValue("gatepassfromdate")
        End If
        If (DateDiff(Microsoft.VisualBasic.DateInterval.Day, dtpfromdate.Value, dtServerDate) > dtFromDate Or DateDiff(Microsoft.VisualBasic.DateInterval.Day, dtpfromdate.Value, dtServerDate) < 0) And ctlcmd.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            MsgBox("From Date Cannot Be Greater Than Current Date" & vbCrLf & " Or Greater Than " & dtFromDate & " Day Back", MsgBoxStyle.Information, ResolveResString(100))
            dtpfromdate.Value = GetServerDate()
            Exit Sub
        End If
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub dtptodate_Leave1(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtptodate.Leave
        On Error GoTo errHandler
        Dim dtToDate As Short
        Dim objdate As New ClsResultSetDB
        Dim dtServerDate As Date
        dtServerDate = GetServerDate()
        objdate.GetResult("select isnull(gatepasstodate,7) as gatepasstodate from  BARCODE_Parameters where UNIT_CODE='" & gstrUNITID & "'")
        If Not objdate.EOFRecord Then
            dtToDate = objdate.GetValue("gatepasstodate")
        End If
        If (DateDiff(Microsoft.VisualBasic.DateInterval.Day, dtServerDate, dtptodate.Value) > dtToDate Or DateDiff(Microsoft.VisualBasic.DateInterval.Day, dtServerDate, dtptodate.Value) < 0) And ctlcmd.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            MsgBox("To Date Cannot Be Less Than Current Date" & vbCrLf & "Or Greater Than " & dtToDate & " Days", MsgBoxStyle.Information, ResolveResString(100))
            dtptodate.Value = GetServerDate()
            Exit Sub
        End If
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0055_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo errHandler
        mdifrmMain.CheckFormName = mintIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0055_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo errHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0055_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo errHandler
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader.Tag)
        FitToClient(Me, framain, ctlFormHeader, ctlcmd, 0)
        dtpfromdate.Format = DateTimePickerFormat.Custom
        dtpfromdate.CustomFormat = gstrDateFormat
        dtpfromdate.Value = GetServerDate()
        dtptodate.Format = DateTimePickerFormat.Custom
        dtptodate.CustomFormat = gstrDateFormat
        dtptodate.Value = GetServerDate()
        cmddocno.Image = My.Resources.ico111.ToBitmap
        ctlcmd.Enabled(1) = False
        ctlcmd.Enabled(2) = False
        ctlcmd.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        cmdshow.Enabled = False
        lblfromdate.Text = "Gate Pass From"
        Call NewfillGrid()
        Call set_gridheader()
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0055_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo errHandler
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub set_gridheader()
        On Error GoTo errHandler
        With fspdetail
            .MaxCols = 6
            .MaxRows = 1
            .UserResize = FPSpreadADO.UserResizeConstants.UserResizeNone
            .Row = 0
            .Col = grddetail.check
            .set_ColWidth(grddetail.check, 0)
            .Text = "Select"
            .Row = 0
            .Col = grddetail.Invoice_No
            .set_ColWidth(grddetail.Invoice_No, 10)
            .Text = "Invoice No"
            .Row = 0
            .Col = grddetail.invoice_qty
            .set_ColWidth(grddetail.invoice_qty, 12)
            .Text = "Invoice Quantity"
            .Row = 0
            .Col = grddetail.Item_Code
            .set_ColWidth(grddetail.Item_Code, 16)
            .Text = "Item Code"
            .Row = 0
            .Col = grddetail.picked_qty
            .set_ColWidth(grddetail.picked_qty, 12)
            .Text = "Picked Quantity"
            .Row = 0
            .Col = grddetail.barcode_track
            .set_ColWidth(grddetail.barcode_track, 0)
            .Text = "Barcode Tracking"
        End With
        With fspheader
            .MaxCols = 4
            .MaxRows = 1
            .UserResize = FPSpreadADO.UserResizeConstants.UserResizeNone
            .Row = 0
            .Col = grdheader.CheckBox
            .set_ColWidth(grdheader.CheckBox, 5)
            .Text = "Select"
            .Row = 0
            .Col = grdheader.Invoice_No
            .set_ColWidth(grdheader.Invoice_No, 10)
            .Text = "Invoice No"
            .Row = 0
            .Col = grdheader.Invoice_Date
            .set_ColWidth(grdheader.Invoice_Date, 10)
            .Text = "Invoice Date"
            .Row = 0
            .Col = grdheader.ShowDetail
            .set_ColWidth(grdheader.ShowDetail, 10)
            .Text = "Show Detail"
        End With
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub fillData()
        On Error GoTo errHandler
        Dim strhdrdata As String
        Dim objfillhdrdata As ClsResultSetDB
        Dim intJ As Integer
        With fspheader
            strhdrdata = "select distinct doc_no,invoice_date from saleschallan_dtl scd" & _
                         " where not exists" & _
                            " (select distinct invoiceno from bar_gatepass_dtl gpdtl" & _
                            " where gpdtl.invoiceno=scd.doc_no and gpdtl.UNIT_CODE=scd.UNIT_CODE)" & _
                        " and invoice_date between '" & Format(dtpfromdate.Value, "dd MMM yyyy") & "' and '" & Format(dtptodate.Value, "dd MMM yyyy") & "'" & _
                        " and scd.cancel_flag=0 and scd.bill_flag=1 and scd.UNIT_CODE='" & gstrUNITID & "'"
            objfillhdrdata = New ClsResultSetDB
            objfillhdrdata.GetResult(strhdrdata)
            If Not objfillhdrdata.EOFRecord Then
                For intJ = 1 To objfillhdrdata.RowCount
                    Call .SetText(grdheader.Invoice_No, intJ, objfillhdrdata.GetValue("doc_no"))
                    Call .SetText(grdheader.Invoice_Date, intJ, VB6.Format(objfillhdrdata.GetValue("invoice_date"), gstrDateFormat))
                    If isfullypicked(objfillhdrdata.GetValue("doc_no")) = False Then
                        .Col = grdheader.CheckBox
                        .Row = intJ
                        .Col2 = 3
                        .Row2 = intJ
                        .BlockMode = True
                        .Lock = True
                        .BlockMode = False
                    Else
                        .Col = grdheader.CheckBox
                        .Row = intJ
                        .Col2 = 1
                        .Row2 = intJ
                        .BlockMode = True
                        .Lock = False
                        .BlockMode = False
                    End If
                    .Row = intJ
                    .Col = grdheader.Invoice_No
                    .Row2 = intJ
                    .Col2 = grdheader.Invoice_Date
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                    objfillhdrdata.MoveNext()
                    .MaxRows = intJ + 1
                Next
            Else
                MsgBox("Record Not Found", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
            objfillhdrdata.ResultSetClose()
            objfillhdrdata = Nothing
            .MaxRows = .MaxRows - 1
        End With
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        objfillhdrdata.ResultSetClose()
    End Sub
    Private Sub SaveData()
        On Error GoTo errHandler
        Dim strhdrdata As String
        Dim strupdatedocumenttypemst As String
        Dim intCounter As Short
        Dim intcheckbox As Short
        Dim strInvoiceNo As String
        If ctlcmd.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            mP_Connection.BeginTrans()
            strhdrdata = "insert into bar_gatepass_hdr(gatepassnum,gatepassuserid,gatepasscreateddt,drivername," & _
                " dispatchperson,securityperson,vehicleno,associatename,PODNo,UNIT_CODE) values(" & _
                " " & Trim(txtdoc_no.Text) & ",'" & mP_User & "','" & getDateForDB(GetServerDate()) & "'," & _
                " '" & Trim(txtdrivername.Text) & "','" & Trim(txtdispatchpersonnel.Text) & "'," & _
                " '" & Trim(txtsecuritypersonnel.Text) & "','" & Trim(txtvehicleno.Text) & "'," & _
                " '" & Trim(txtassociatename.Text) & "','" & txtPODNo.Text & "','" & gstrUNITID & "')"
            mP_Connection.Execute(strhdrdata, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            With fspheader
                For intCounter = 1 To .MaxRows
                    .Row = intCounter : .Col = grdheader.CheckBox : intcheckbox = CShort(.Value)
                    If intcheckbox = 1 Then
                        .Row = intCounter : .Col = grdheader.Invoice_No : strInvoiceNo = .Value
                        Call filldetail(strInvoiceNo)
                        Call insertdetail()
                    End If
                Next
            End With
            strupdatedocumenttypemst = "update documenttype_mst set current_no=" & Trim(txtdoc_no.Text) & " where doc_type=306 and description='Gate Pass' and UNIT_CODE='" & gstrUNITID & "'"
            mP_Connection.Execute(strupdatedocumenttypemst, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            mP_Connection.CommitTrans()
            MsgBox("Gate Pass Saved Successfully With Document No  " & Trim(txtdoc_no.Text), MsgBoxStyle.Information, ResolveResString(100))
            Call ClearControls()
            txtdoc_no.Enabled = True : txtdoc_no.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmddocno.Enabled = True
            cmdshow.Enabled = False
            ctlcmd.Revert()
            ctlcmd.Enabled(1) = False
            ctlcmd.Enabled(2) = False
            ctlcmd.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        End If
        If ctlcmd.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            mP_Connection.BeginTrans()
            mP_Connection.Execute("update bar_gatepass_hdr set PODNo='" & Trim(txtPODNo.Text) & "' where gatepassnum = '" & txtdoc_no.Text & "' and UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            mP_Connection.CommitTrans()
            MsgBox("Document No " & Trim(txtdoc_no.Text) & " Updated Successfully", MsgBoxStyle.Information, ResolveResString(100))
            Call ClearControls()
            txtdoc_no.Enabled = True : txtdoc_no.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmddocno.Enabled = True
            cmdshow.Enabled = False
            ctlcmd.Revert()
            ctlcmd.Enabled(1) = False
            ctlcmd.Enabled(2) = False
            ctlcmd.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        End If
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function Generate_GatePassNo() As Short
        On Error GoTo errHandler
        Dim objgatepassno As New ClsResultSetDB
        Dim strgatepassno As String
        Dim intgatepassno As Short
        strgatepassno = "select isnull(current_no,0)  as current_no from documenttype_mst where doc_type=306 and getdate() between Fin_Start_date and Fin_end_date and UNIT_CODE='" & gstrUNITID & "'"
        objgatepassno.GetResult(strgatepassno)
        If Not objgatepassno.EOFRecord Then
            intgatepassno = Val(objgatepassno.GetValue("current_no")) + 1
        End If
        Generate_GatePassNo = intgatepassno
        Exit Function
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub ClearControls()
        On Error GoTo errHandler
        With fspdetail
            .MaxRows = 1
            .Row = 1
            .Col = grddetail.Invoice_No
            .Text = ""
            .Col = grddetail.invoice_qty
            .Text = ""
            .Col = grddetail.Item_Code
            .Text = ""
            .Col = grddetail.picked_qty
            .Text = ""
            .Col = grddetail.check
            .Value = CStr(0)
        End With
        With fspheader
            .MaxRows = 1
            .Row = 1
            .Col = grdheader.CheckBox
            .Value = CStr(0)
            .Col = grdheader.Invoice_Date
            .Text = ""
            .Col = grdheader.Invoice_No
            .Text = ""
        End With
        txtdoc_no.Text = ""
        txtdrivername.Text = ""
        txtassociatename.Text = ""
        txtdispatchpersonnel.Text = ""
        txtvehicleno.Text = ""
        txtsecuritypersonnel.Text = ""
        txtPODNo.Text = ""
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function ValidateData() As Boolean
        On Error GoTo errHandler
        Dim intCounter As Short
        Dim blnitemselected As Boolean
        ValidateData = True
        blnitemselected = False
        With fspheader
            If .MaxRows = 1 Then
                .Row = 1 : .Col = grdheader.Invoice_No
                If Val(.Value) = 0 Then
                    MsgBox("No Detail To Save", MsgBoxStyle.Information, ResolveResString(100))
                    ValidateData = False
                    Exit Function
                End If
            End If
            For intCounter = 1 To .MaxRows
                .Row = intCounter
                .Col = grdheader.CheckBox
                If CDbl(.Value) = 1 Then
                    blnitemselected = True
                    ValidateData = True
                    Exit For
                End If
                blnitemselected = False
                ValidateData = False
            Next
            If blnitemselected = False Then
                MsgBox("Please Select Atleast One Invoice", MsgBoxStyle.Information, ResolveResString(100))
                ValidateData = False
                Exit Function
            End If
            If Trim(txtdrivername.Text) = "" Then
                ValidateData = False
                MsgBox("Driver Name Cannot Be Left Blank", MsgBoxStyle.Information, ResolveResString(100))
                If txtdrivername.Enabled = True Then txtdrivername.Focus()
                Exit Function
            End If
            If Trim(txtvehicleno.Text) = "" Then
                ValidateData = False
                MsgBox("Vehicle No Cannot Be Left Blank", MsgBoxStyle.Information, ResolveResString(100))
                If txtvehicleno.Enabled = True Then txtvehicleno.Focus()
                Exit Function
            End If
            If Trim(txtassociatename.Text) = "" Then
                ValidateData = False
                MsgBox("Associate Name Cannot Be Left Blank", MsgBoxStyle.Information, ResolveResString(100))
                If txtassociatename.Enabled = True Then txtassociatename.Focus()
                Exit Function
            End If
        End With
        Exit Function
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub FillDataInViewMode()
        On Error GoTo errHandler
        Dim strSQL As String
        Dim lnginvoiceno As Integer
        Dim strItemCode As String
        Dim dblInvoiceQty As Double
        Dim dblpickedqty As Double
        Dim intCounter As Short
        Dim objhdrdata As ClsResultSetDB
        Dim intJ As Short
        Dim objheaderfields As New ClsResultSetDB
        objhdrdata = New ClsResultSetDB
        objhdrdata.GetResult("select distinct invoiceno,invoice_date from bar_gatepass_dtl bgdtl" & _
            " inner join saleschallan_dtl scd on bgdtl.invoiceno=scd.doc_no" & _
            " and bgdtl.UNIT_CODE=scd.UNIT_CODE" & _
            " where gatepassnum=" & Trim(txtdoc_no.Text) & " and bgdtl.UNIT_CODE='" & gstrUNITID & "'")
        If Not objhdrdata.EOFRecord Then
            With fspheader
                For intJ = 1 To objhdrdata.RowCount
                    Call .SetText(grdheader.Invoice_No, intJ, objhdrdata.GetValue("invoiceno"))
                    Call .SetText(grdheader.Invoice_Date, intJ, VB6.Format(objhdrdata.GetValue("invoice_date"), gstrDateFormat))
                    Call .SetText(grdheader.CheckBox, intJ, 1)
                    .MaxRows = intJ + 1
                    objhdrdata.MoveNext()
                Next
                .MaxRows = .MaxRows - 1
                .Row = 1
                .Row2 = .MaxRows
                .Col = grdheader.CheckBox
                .Col2 = grdheader.Invoice_Date
                .BlockMode = True
                .Lock = True
                .BlockMode = False
            End With
        End If
        objheaderfields.GetResult("select drivername,vehicleno,associatename,securityperson,dispatchperson," & _
            " PODNo from bar_gatepass_hdr where gatepassnum=" & Trim(txtdoc_no.Text) & " and UNIT_CODE='" & gstrUNITID & "'")
        If Not objheaderfields.EOFRecord Then
            txtdrivername.Text = objheaderfields.GetValue("drivername")
            txtvehicleno.Text = objheaderfields.GetValue("vehicleno")
            txtassociatename.Text = objheaderfields.GetValue("associatename")
            txtsecuritypersonnel.Text = objheaderfields.GetValue("securityperson")
            txtdispatchpersonnel.Text = objheaderfields.GetValue("dispatchperson")
            txtPODNo.Text = objheaderfields.GetValue("PODNo")
        End If
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub NewfillGrid()
        On Error GoTo errHandler
        With fspdetail
            .Row = -1
            .Col = grddetail.check
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
            .TypeCheckCenter = True
            .Value = CStr(0)
            .Col = grddetail.Invoice_No
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Value = ""
            .Col = grddetail.invoice_qty
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .Value = ""
            .Col = grddetail.Item_Code
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Value = ""
            .Col = grddetail.picked_qty
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .Col = grddetail.barcode_track
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
        End With
        With fspheader
            .Row = -1
            .Col = grdheader.CheckBox
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
            .TypeCheckCenter = True
            .Value = CStr(0)
            .Col = grdheader.Invoice_No
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Value = ""
            .Col = grdheader.Invoice_Date
            '.CellType = FPSpreadADO.CellTypeConstants.CellTypeDate
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .Value = ""
            .Col = grdheader.ShowDetail
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
            .TypeButtonText = "Show Detail"
        End With


        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub fspheader_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles fspheader.ButtonClicked
        On Error GoTo errHandler
        Dim strInvoiceNo As String
        Dim intCounter As Short
        With fspheader
            If e.row = -1 Then
                Exit Sub
            End If
            .Row = e.row : .Col = e.col
            For intCounter = 1 To .MaxRows
                .Row = intCounter
                .TypeButtonColor = System.Drawing.Color.Beige
            Next
            If e.col = grdheader.ShowDetail Then
                e.row = e.row
                .TypeButtonColor = System.Drawing.ColorTranslator.FromOle(&H808080)
            End If
            If e.col = grdheader.ShowDetail Then
                .Row = e.row : .Col = grdheader.Invoice_No : strInvoiceNo = .Value
                If strInvoiceNo <> "" Then
                    If ctlcmd.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                        Call filldetail(strInvoiceNo)
                    Else
                        Call filldetailinviewmode(strInvoiceNo)
                    End If
                End If
            End If
        End With
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtassociatename_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtassociatename.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo errHandler
        Call validate_char(KeyAscii)
        GoTo EventExitSub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtdispatchpersonnel_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtdispatchpersonnel.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo errHandler
        Call validate_char(KeyAscii)
        GoTo EventExitSub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtDoc_No_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtdoc_no.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo errHandler
        If KeyCode = 112 Then
            Call CmdDocNo_Click(cmddocno, New System.EventArgs())
        End If
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtDoc_No_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtdoc_no.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo errHandler
        If txtdoc_no.Text = "" Then
            Call ClearControls()
        End If
        If KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{tab}")
        End If
        If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
        GoTo EventExitSub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtDoc_No_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtdoc_no.Leave
        On Error GoTo errHandler
        'If Trim(txtdoc_no.Text) <> "" Then
        '    Call CmdDocNo_Click(cmddocno, New System.EventArgs())
        'End If
        If LTrim(txtdoc_no.Text) = "" Then
            MsgBox("Record Not Found", MsgBoxStyle.Information, ResolveResString(100))
            Call ClearControls()
            Exit Sub
        Else
            FillDataInViewMode()
            Call setcontrolenableproperty((ctlcmd.Mode))
        End If

        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtDoc_No_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtdoc_no.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo errHandler
        Dim objexist As New ClsResultSetDB
        If Trim(txtdoc_no.Text) <> "" Then
            If Not IsNumeric(txtdoc_no.Text) Then
                MsgBox("Please Enter A Valid Gate Pass Number", MsgBoxStyle.Information, ResolveResString(100))
                txtdoc_no.Text = ""
            End If
        End If
        If Trim(txtdoc_no.Text) <> "" Then
            objexist.GetResult("select * from bar_gatepass_dtl where gatepassnum=" & txtdoc_no.Text & " and UNIT_CODE='" & gstrUNITID & "'")
            If objexist.EOFRecord Then
                MsgBox("Record Not Found", MsgBoxStyle.Information, ResolveResString(100))
                Call ClearControls()
            Else
                Call FillDataInViewMode()
                ctlcmd.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                'ctlcmd.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
            End If
        End If
        GoTo EventExitSub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub filldetail(ByRef strInvoiceNo As String)
        On Error GoTo errHandler
        Dim dblInvoiceQty As Double
        Dim dblpickedqty As Double
        Dim intCounter As Short
        Dim objfilldata As New ClsResultSetDB
        Dim strSQL As String
        Dim inti As Short
        strSQL = "select sd.item_code,scd.doc_no,sales_quantity,case when im.barcode_tracking=1 then isnull(bbsd.picked_qty,0) else isnull(sd.sales_quantity,0) end as picked_qty," & _
                 " im.barcode_tracking from saleschallan_dtl scd" & _
                 " inner join sales_dtl sd on scd.doc_no=sd.doc_no" & _
                 " and scd.location_code=sd.location_code" & _
                 " and scd.UNIT_CODE=sd.UNIT_CODE" & _
                 " inner join item_mst im on im.item_code=sd.item_code" & _
                 " and im.UNIT_CODE=sd.UNIT_CODE" & _
                 " left outer join (select isnull(sum(quantity),0) as picked_qty,invoice_no,item_alias,UNIT_CODE" & _
                 " from  BAR_BONDEDSTOCK_DTL where status_flag='L' and UNIT_CODE='" & gstrUNITID & "'" & _
                 " group by invoice_no,item_alias,UNIT_CODE) bbsd on (bbsd.invoice_no=scd.doc_no)" & _
                 " and bbsd.UNIT_CODE=scd.UNIT_CODE" & _
                 " AND BBSD.ITEM_ALIAS=im.ITm_itemalias and sd.item_code=im.item_code" & _
                 " where invoice_date between '" & Format(dtpfromdate.Value, "dd MMM yyyy") & "' and '" & Format(dtptodate.Value, "dd MMM yyyy") & "'" & _
                 " and scd.doc_no=" & strInvoiceNo & "" & _
                 " and scd.cancel_flag=0 and scd.bill_flag=1 and scd.UNIT_CODE='" & gstrUNITID & "'"
        objfilldata.GetResult(strSQL)
        If Not objfilldata.EOFRecord Then
            For intCounter = 1 To objfilldata.RowCount
                With fspdetail
                    Call .SetText(grddetail.Invoice_No, intCounter, objfilldata.GetValue("doc_no"))
                    Call .SetText(grddetail.invoice_qty, intCounter, objfilldata.GetValue("sales_quantity"))
                    Call .SetText(grddetail.Item_Code, intCounter, objfilldata.GetValue("item_code"))
                    Call .SetText(grddetail.picked_qty, intCounter, objfilldata.GetValue("picked_qty"))
                    Call .SetText(grddetail.barcode_track, intCounter, IIf(objfilldata.GetValue("barcode_tracking") = True, 1, 0))
                    Call .SetText(grddetail.check, intCounter, 0)
                    .MaxRows = intCounter + 1
                End With
                objfilldata.MoveNext()
            Next
        Else
            MsgBox("Record Not Found", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If
        fspdetail.MaxRows = fspdetail.MaxRows - 1
        With fspdetail
            .Row = 1
            .Row2 = .MaxRows
            .Col = grddetail.Invoice_No
            .Col2 = grddetail.picked_qty
            .BlockMode = True
            .Lock = True
            .BlockMode = False
        End With
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function isfullypicked(ByRef strInvoiceNo As Object) As Boolean
        On Error GoTo errHandler
        Dim objgetData As ClsResultSetDB
        Dim strSQL As String
        Dim inti As Short
        Dim dblpickedqty As Double
        Dim dblinvquantity As Double
        objgetData = New ClsResultSetDB
        strSQL = "select sd.item_code,sd.sales_quantity,case when barcode_tracking=1 then isnull(bbsd.picked_qty,0) else isnull(sd.sales_quantity,0) end as picked_qty" & _
                 " from sales_dtl sd inner join item_mst im on sd.item_code=im.item_code and sd.UNIT_CODE=im.UNIT_CODE left outer join(select isnull(sum(quantity),0) as picked_qty,invoice_no,UNIT_CODE," & _
                 " item_alias from  BAR_BONDEDSTOCK_DTL where status_flag='L' and UNIT_CODE='" & gstrUNITID & "' group by UNIT_CODE,invoice_no,item_alias) bbsd on bbsd.item_alias=im.itm_itemalias and" & _
                 " bbsd.UNIT_CODE=im.UNIT_CODE and sd.Item_Code = im.Item_Code And bbsd.Invoice_No = sd.Doc_No Where sd.Doc_No = " & strInvoiceNo & " and sd.UNIT_CODE='" & gstrUNITID & "'"
        objgetData.GetResult(strSQL)
        If Not objgetData.EOFRecord Then
            For inti = 1 To objgetData.RowCount
                dblpickedqty = dblpickedqty + objgetData.GetValue("picked_qty")
                dblinvquantity = dblinvquantity + objgetData.GetValue("sales_quantity")
                objgetData.MoveNext()
            Next
        End If
        If dblpickedqty = dblinvquantity Then
            isfullypicked = True
        Else
            isfullypicked = False
        End If
        Exit Function
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub insertdetail()
        On Error GoTo errHandler
        Dim strItemCode As String
        Dim dblInvoiceQty As Double
        Dim dblpickedqty As Double
        Dim lnginvoiceno As Integer
        Dim intcheck As Short
        Dim blnbarcode_tracking As Short
        Dim strdtldata As String
        Dim intCounter As Short
        With fspdetail
            For intCounter = 1 To .MaxRows
                .Row = intCounter : .Col = grddetail.Invoice_No : lnginvoiceno = Val(.Value)
                .Row = intCounter : .Col = grddetail.invoice_qty : dblInvoiceQty = Val(.Value)
                .Row = intCounter : .Col = grddetail.Item_Code : strItemCode = .Value
                .Row = intCounter : .Col = grddetail.picked_qty : dblpickedqty = Val(.Value)
                .Row = intCounter : .Col = grddetail.check : intcheck = Val(.Value)
                .Row = intCounter : .Col = grddetail.barcode_track : blnbarcode_tracking = Val(.Value)
                If lnginvoiceno <> Val("") Then
                    strdtldata = "insert into bar_gatepass_dtl(gatepassnum,invoiceno,itemcode,invoicequantity," & _
                        " pickedquantity,barcode_tracking,UNIT_CODE)" & _
                        " values(" & Trim(txtdoc_no.Text) & "," & lnginvoiceno & ",'" & strItemCode & "'," & _
                        " " & dblInvoiceQty & "," & dblpickedqty & "," & blnbarcode_tracking & ",'" & gstrUNITID & "')"
                    mP_Connection.Execute(strdtldata, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
            Next
        End With
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub filldetailinviewmode(ByRef strInvoiceNo As String)
        On Error GoTo errHandler
        Dim objdetaildata As New ClsResultSetDB
        Dim intCounter As Short
        With fspdetail
            objdetaildata.GetResult("select invoiceno,itemcode,invoicequantity,pickedquantity,barcode_tracking" & _
                " from bar_gatepass_dtl  where gatepassnum=" & Trim(txtdoc_no.Text) & " " & _
                " and invoiceno=" & strInvoiceNo & " and UNIT_CODE='" & gstrUNITID & "'")
            If Not objdetaildata.EOFRecord Then
                For intCounter = 1 To objdetaildata.RowCount
                    Call .SetText(grddetail.Invoice_No, intCounter, objdetaildata.GetValue("invoiceno"))
                    Call .SetText(grddetail.invoice_qty, intCounter, objdetaildata.GetValue("invoicequantity"))
                    Call .SetText(grddetail.Item_Code, intCounter, objdetaildata.GetValue("itemcode"))
                    Call .SetText(grddetail.picked_qty, intCounter, objdetaildata.GetValue("pickedquantity"))
                    Call .SetText(grddetail.barcode_track, intCounter, IIf(objdetaildata.GetValue("barcode_tracking") = True, 1, 0))
                    Call .SetText(grddetail.check, intCounter, 0)
                    .MaxRows = intCounter + 1
                    objdetaildata.MoveNext()
                Next
                .MaxRows = .MaxRows - 1
            End If
            .Row = 1
            .Row2 = .MaxRows
            .Col = grddetail.check
            .Col2 = grddetail.barcode_track
            .BlockMode = True
            .Lock = True
            .BlockMode = False
        End With
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub validate_char(ByRef KeyAscii As Short)
        On Error GoTo errHandler
        If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii >= 97 And KeyAscii <= 122 Then
        ElseIf KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{tab}")
        Else
            KeyAscii = 0
        End If
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtdrivername_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtdrivername.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo errHandler
        Call validate_char(KeyAscii)
        GoTo EventExitSub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtsecuritypersonnel_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtsecuritypersonnel.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo errHandler
        Call validate_char(KeyAscii)
        GoTo EventExitSub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtvehicleno_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtvehicleno.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo errHandler
        If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 45 Or KeyAscii = 40 Or KeyAscii = 41 Then
        ElseIf KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{tab}")
        Else
            KeyAscii = 0
        End If
        GoTo EventExitSub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub setcontrolenableproperty(ByRef mode As Short)
        On Error GoTo errHandler
        If ctlcmd.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            txtassociatename.Enabled = True
            txtdrivername.Enabled = True
            txtdispatchpersonnel.Enabled = True
            txtvehicleno.Enabled = True
            txtsecuritypersonnel.Enabled = True
        Else
            txtassociatename.Enabled = False
            txtdrivername.Enabled = False
            txtdispatchpersonnel.Enabled = False
            txtvehicleno.Enabled = False
            txtsecuritypersonnel.Enabled = False
            txtPODNo.Enabled = False
            ctlcmd.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
        End If
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtPODNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPODNo.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        On Error GoTo errHandler
        If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 45 Or KeyAscii = 40 Or KeyAscii = 41 Then
        ElseIf KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{tab}")
        Else
            KeyAscii = 0
        End If
        GoTo EventExitSub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub
    Private Sub ctlcmd_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles ctlcmd.ButtonClick
        On Error GoTo errHandler
        Dim strSQL As String
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                Call ClearControls()
                cmddocno.Enabled = False
                txtdoc_no.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                dtpfromdate.Value = GetServerDate()
                dtptodate.Value = GetServerDate()
                txtdoc_no.Enabled = False
                lblfromdate.Text = "Invoice From"
                cmdshow.Enabled = True
                Call setcontrolenableproperty((ctlcmd.Mode))
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                If ValidateData() = True Then
                    If ctlcmd.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                        txtdoc_no.Text = CStr(Generate_GatePassNo())
                    End If
                    Call SaveData()
                    lblfromdate.Text = "Gate Pass From"
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                If ctlcmd.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION, 60095) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        ctlcmd.Revert()
                        ctlcmd.Enabled(1) = False
                        ctlcmd.Enabled(2) = False
                        ctlcmd.Enabled(5) = False
                        Call ClearControls()
                        txtdoc_no.Text = ""
                        txtdoc_no.Enabled = True
                        cmdshow.Enabled = False
                        dtpfromdate.Value = GetServerDate()
                        dtptodate.Value = GetServerDate()
                        cmddocno.Enabled = True
                        txtPODNo.Text = ""
                        txtdoc_no.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        txtPODNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        lblfromdate.Text = "Gate Pass From"
                    End If
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                'ctlcmd.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
                txtPODNo.Enabled = True
                txtPODNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
        End Select
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
End Class