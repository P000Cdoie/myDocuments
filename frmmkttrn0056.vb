Option Strict Off
Option Explicit On
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Friend Class frmMKTTRN0056
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
	' Revision History           -   disable fields driver name,associate name,security person
	'                                vehicle no,security person and these fields will not
	'                                updated from this form
    '------------------------------------------------------------------------
    'Modified By JY on 16-May-2011
    '   Modified to support MultiUnit functionality
    '****************************************************
    Dim mintIndex As Short
    Private Enum grddetail
        Invoice_No = 1
        Item_Code = 2
        invoice_qty = 3
        picked_qty = 4
    End Enum
    Private Sub cmdgatepassno_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdgatepassno.Click
        On Error GoTo errHandler
        Dim strHelp As String
        Dim strgatepassno() As String
        Dim strSQL As String
        Dim objhdrdata As New ClsResultSetDB
        strHelp = "select GatePassNum,Gatepasscreateddt from bar_gatepass_hdr where isnull(Authorize_flag,0)=0 and UNIT_CODE='" & gstrUNITID & "'"
        strgatepassno = ctlEMPHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, "Gate Pass No")
        If Not (UBound(strgatepassno) = -1) Then
            If Len(strgatepassno(0)) >= 1 And strgatepassno(0) = "0" Then
                MsgBox("Record Not Found", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            Else
                TxtGatePassNo.Text = strgatepassno(0)
                objhdrdata.GetResult("select drivername,vehicleno,associatename,securityperson,dispatchperson,PODNo from bar_gatepass_hdr where gatepassnum=" & Trim(TxtGatePassNo.Text) & " and UNIT_CODE='" & gstrUNITID & "'")
                If Not objhdrdata.EOFRecord Then
                    TxtDriverName.Text = objhdrdata.GetValue("drivername")
                    TxtVehicleNo.Text = objhdrdata.GetValue("vehicleno")
                    TxtAssociateName.Text = objhdrdata.GetValue("associatename")
                    TxtSecurityPersonnel.Text = objhdrdata.GetValue("securityperson")
                    TxtDispatchPersonnel.Text = objhdrdata.GetValue("dispatchperson")
                    txtPODNo.Text = objhdrdata.GetValue("PODNo")
                End If
                Call set_gridheader()
                Call NewfillGrid()
            End If
        End If
        If Not IsDBNull(objhdrdata) Then
            objhdrdata = Nothing
        End If
        Exit Sub
errHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdshowitems_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdshowitems.Click
        On Error GoTo errHandler
        If Trim(TxtGatePassNo.Text) <> "" Then
            Call fillData()
        Else
            MsgBox("Please Select Gate Pass No", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If
        Exit Sub
errHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdshowitems_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cmdshowitems.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo errHandler
        If KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{tab}")
        End If
        GoTo EventExitSub
errHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub frmMKTTRN0056_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo errHandler
        mdifrmMain.CheckFormName = mintIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0056_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo errHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0056_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo errHandler
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader.Tag)
        FitToClient(Me, framain, ctlFormHeader, ctlcmd, 500)
        'Call FillLabelFromResFile(Me)
        txtdate.Text = CStr(GetServerDate())
        txttime.Text = getservertime()
        txtdate.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txttime.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        TxtVehicleNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        TxtAssociateName.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        TxtDispatchPersonnel.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        TxtSecurityPersonnel.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        TxtDriverName.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtPODNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        TxtGatePassNo.Enabled = False
        TxtGatePassNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        cmdgatepassno.Image = My.Resources.ico111.ToBitmap
        ctlcmd.Enabled(0) = True
        ctlcmd.Enabled(1) = False
        ctlcmd.Enabled(2) = True
        ctlcmd.Enabled(3) = True
        Call NewfillGrid()
        Call set_gridheader()
        txtPODNo.Enabled = False
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0056_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo errHandler
        Me.Dispose()
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub fillData()
        On Error GoTo errHandler
        Dim strSQL As String
        Dim objdtldata As New ClsResultSetDB
        Dim lnginvoiceno As Integer
        Dim strItemCode As String
        Dim dblInvoiceQty As Double
        Dim dblpickedqty As Double
        Dim intCounter As Short
        objdtldata.GetResult("select * from bar_gatepass_dtl where gatepassnum=" & TxtGatePassNo.Text & " and UNIT_CODE='" & gstrUNITID & "'")
        With fspdetail
            If Not objdtldata.EOFRecord Then
                For intCounter = 1 To objdtldata.RowCount
                    lnginvoiceno = objdtldata.GetValue("invoiceno")
                    strItemCode = objdtldata.GetValue("itemcode")
                    dblInvoiceQty = objdtldata.GetValue("invoicequantity")
                    dblpickedqty = objdtldata.GetValue("pickedquantity")
                    Call .SetText(grddetail.Invoice_No, intCounter, lnginvoiceno)
                    Call .SetText(grddetail.Item_Code, intCounter, strItemCode)
                    Call .SetText(grddetail.invoice_qty, intCounter, dblInvoiceQty)
                    Call .SetText(grddetail.picked_qty, intCounter, dblpickedqty)
                    .MaxRows = intCounter + 1
                    objdtldata.MoveNext()
                Next
                .MaxRows = .MaxRows - 1
            Else
                MsgBox("Record Not Found", MsgBoxStyle.Information, ResolveResString(100))
                Call ClearControls()
                Exit Sub
            End If
        End With
        With fspdetail
            .Row = 1
            .Row2 = .MaxRows
            .Col = grddetail.Invoice_No
            .Col2 = grddetail.picked_qty
            .BlockMode = True
            .Lock = True
            .BlockMode = False
        End With
        System.Windows.Forms.SendKeys.Send("{tab}")
        If Not IsDBNull(objdtldata) Then
            objdtldata = Nothing
        End If
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub NewfillGrid()
        On Error GoTo errHandler
        With fspdetail
            .Row = -1
            .Col = grddetail.Invoice_No
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Value = ""
            .Col = grddetail.invoice_qty
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Value = ""
            .Col = grddetail.Item_Code
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Value = ""
            .Col = grddetail.picked_qty
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .Value = ""
        End With
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub set_gridheader()
        On Error GoTo errHandler
        With fspdetail
            .MaxCols = 4
            .MaxRows = 1
            .Row = 0
            .Col = grddetail.Invoice_No
            .set_ColWidth(grddetail.Invoice_No, 1000)
            .Text = "Invoice No"
            .Row = 0
            .Col = grddetail.invoice_qty
            .set_ColWidth(grddetail.invoice_qty, 1500)
            .Text = "Invoice Quantity"
            .Row = 0
            .Col = grddetail.Item_Code
            .set_ColWidth(grddetail.Item_Code, 3000)
            .Text = "Item Code"
            .Row = 0
            .Col = grddetail.picked_qty
            .set_ColWidth(grddetail.picked_qty, 1500)
            .Text = "Picked Quantity"
        End With
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SaveData()
        On Error GoTo errHandler
        Dim strMSG As String
        Dim strSQL As String
        txttime.Text = getservertime()
        strMSG = getcancelledinv(Trim(TxtGatePassNo.Text))
        If Trim(strMSG) <> "" Then
            MsgBox("Following Invoice Has Been Cancelled" & vbCrLf & strMSG, MsgBoxStyle.Information, ResolveResString(100))
        End If
        strSQL = "update bar_gatepass_hdr set " & _
                 " gatepassauthorizationtime='" & Trim(VB6.Format(txttime.Text, "hh:mm")) & "'," & _
                 " gatepassauthorizationdt='" & getDateForDB(GetServerDate()) & "',authorize_flag=1 , PODNO = '" & txtPODNo.Text & "'" & _
                 " where gatepassnum=" & TxtGatePassNo.Text & " and UNIT_CODE='" & gstrUNITID & "'"
        mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        MsgBox("Gate Pass Authorized Successfully", MsgBoxStyle.Information, ResolveResString(100))
        Call ClearControls()
        txttime.Text = getservertime()
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
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
        End With
        TxtGatePassNo.Text = ""
        TxtAssociateName.Text = ""
        TxtDispatchPersonnel.Text = ""
        TxtDriverName.Text = ""
        TxtSecurityPersonnel.Text = ""
        TxtVehicleNo.Text = ""
        txtPODNo.Text = ""
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub fspdetail_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_KeyPressEvent)
        On Error GoTo errHandler
        If eventArgs.keyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{tab}")
        End If
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub TxtGatePassNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TxtGatePassNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo errHandler
        If KeyCode = 112 Then
            Call cmdgatepassno_Click(cmdgatepassno, New System.EventArgs())
        End If
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub TxtGatePassNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtGatePassNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo errHandler
        If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
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
    Private Sub txtgatepassnum_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPODNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo errHandler
        Call validate_num(KeyAscii)
        GoTo EventExitSub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Function validate_data() As Boolean
        On Error GoTo errHandler
        Dim objdata As New ClsResultSetDB
        validate_data = True
        If Trim(TxtGatePassNo.Text) = "" Then
            validate_data = False
            MsgBox("Gate Pass Number Cannot Be Left Blank", MsgBoxStyle.Information, ResolveResString(100))
            Exit Function
        End If
        objdata.GetResult("select * from bar_gatepass_hdr where gatepassnum=" & TxtGatePassNo.Text & " and UNIT_CODE='" & gstrUNITID & "'")
        If objdata.EOFRecord Then
            MsgBox("Gate Pass No Does Not Exists", MsgBoxStyle.Information, ResolveResString(100))
            validate_data = False
            Call ClearControls()
            If Not IsDBNull(objdata) Then
                objdata = Nothing
            End If
            Exit Function
        End If
        Exit Function
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function getservertime() As String
        On Error GoTo errHandler
        Dim objtime As New ClsResultSetDB
        Dim strTime As String
        objtime.GetResult("select convert(varchar(5),getdate(),108) as time")
        If Not objtime.EOFRecord Then
            strTime = objtime.GetValue("time")
        End If
        getservertime = strTime
        objtime = Nothing
        Exit Function
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function getcancelledinv(ByRef strgatepassno As String) As String
        On Error GoTo errHandler
        Dim objcancelledinv As New ClsResultSetDB
        Dim strinvoicelst As String
        Dim intCounter As Short
        objcancelledinv.GetResult("select distinct invoiceno from bar_gatepass_dtl bgd inner join saleschallan_dtl scd on bgd.invoiceno=scd.doc_no and bgd.UNIT_CODE=scd.UNIT_CODE where bgd.gatepassnum=" & strgatepassno & " and cancel_flag=1 and bgd.UNIT_CODE='" & gstrUNITID & "'")
        If Not objcancelledinv.EOFRecord Then
            For intCounter = 1 To objcancelledinv.RowCount
                strinvoicelst = strinvoicelst & objcancelledinv.GetValue("invoiceno") & ","
                objcancelledinv.MoveNext()
            Next
            strinvoicelst = Mid(strinvoicelst, 1, Len(strinvoicelst) - 1)
        End If
        getcancelledinv = strinvoicelst
        Exit Function
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function validategatepass() As Boolean
        On Error GoTo errHandler
        Dim objdata As New ClsResultSetDB
        validategatepass = True
        TxtGatePassNo.Enabled = True
        If Trim(TxtGatePassNo.Text) = "" Then
            validategatepass = False
            MsgBox("Enter Gate Pass Number To Print", MsgBoxStyle.Information, ResolveResString(100))
            If txtPODNo.Enabled = True Then txtPODNo.Focus()
            Exit Function
        End If
        objdata.GetResult("select * from bar_gatepass_hdr where gatepassnum=" & TxtGatePassNo.Text & " and Authorize_Flag=1 and UNIT_CODE='" & gstrUNITID & "'")
        If objdata.EOFRecord Then
            MsgBox("Either Gate Pass No Does Not Exists" & vbCrLf & "Or It Is Not Authorized", MsgBoxStyle.Information, ResolveResString(100))
            validategatepass = False
            txtPODNo.Text = ""
            If txtPODNo.Enabled = True Then txtPODNo.Focus()
            Exit Function
        End If
        If Not IsDBNull(objdata) Then
            objdata = Nothing
        End If
        Exit Function
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub validate_num(ByRef KeyAscii As Short)
        On Error GoTo errHandler
        If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlcmd_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.cmdGrpAuthorise.ButtonClickEventArgs) Handles ctlcmd.ButtonClick
        On Error GoTo errHandler
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_AUTHORIZE
                If validate_data() = True Then
                    Call SaveData()
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
                'With crrptgatepass
                '    If validategatepass() = True Then
                '        Call UpdateRegistryDSNProperties(gstrCONNECTIONDSN, gstrCONNECTIONDATABASE, gstrCONNECTIONSERVER)
                '        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
                '        .Reset()
                '        .DiscardSavedData = True
                '        .Connect = gstrREPORTCONNECT
                '        .WindowShowPrintSetupBtn = True
                '        .WindowShowSearchBtn = True
                '        .WindowShowPrintBtn = True
                '        .WindowShowExportBtn = True
                '        .WindowState = Crystal.WindowStateConstants.crptMaximized
                '        .WindowTitle = "Gate Pass Authorization"
                '        .set_Formulas(0, "fmlcompany='" & gstrCOMPANY & "'")
                '        .set_Formulas(1, "fmladdress1='" & gstr_WRK_ADDRESS1 & "'")
                '        .set_Formulas(2, "fmladdress2='" & gstr_WRK_ADDRESS2 & "'")
                '        .ReportFileName = My.Application.Info.DirectoryPath & "\reports\rptgatepassentry.rpt"
                '        .SelectionFormula = "{bar_gatepass_hdr.gatepassnum}=" & Trim(TxtGatePassNo.Text) & ""
                '        .Action = 1
                '        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                '    End If
                'End With
                If validategatepass() = True Then
                    CreateCrystalReport()
                End If
        End Select
        Exit Sub
errHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub CreateCrystalReport()
        On Error GoTo ErrHandler
        Dim strSelectionFormula As String 'Holds the Selection Formula to be passed to Report
        Dim strReportName As String ' Holds the name of the report
        Dim RepDoc As ReportDocument
        Dim RepPath As String
        Dim FrmCRViewer As New eMProCrystalReportViewer
        RepDoc = FrmCRViewer.GetReportDocument()
        FrmCRViewer.ReportHeader = ctlFormHeader.HeaderString()
        FrmCRViewer.ShowPrintButton = True
        FrmCRViewer.ShowTextSearchButton = True
        FrmCRViewer.ShowZoomButton = True
        With RepDoc
            'strReportName = My.Application.Info.DirectoryPath & "\reports\rptgatepassentry.rpt"
            'If CheckFile(strReportName) = False Then
            strReportName = "\reports\rptgatepassentry.rpt"
            'End If
            RepPath = My.Application.Info.DirectoryPath & strReportName
            .Load(RepPath)
            .DataDefinition.FormulaFields("fmlcompany").Text = "'" & gstrCOMPANY & "'"
            .DataDefinition.FormulaFields("fmladdress1").Text = "'" & gstr_WRK_ADDRESS1 & "'"
            .DataDefinition.FormulaFields("fmladdress2").Text = "'" & gstr_WRK_ADDRESS2 & "'"
            strSelectionFormula = "{bar_gatepass_hdr.gatepassnum}=" & Trim(TxtGatePassNo.Text) & " AND {bar_gatepass_hdr.UNIT_CODE} =  '" & gstrUNITID & "'"
            .RecordSelectionFormula = strSelectionFormula
        End With
        FrmCRViewer.Show()
        Exit Sub
ErrHandler:
        If Err.Number = 20545 Then
            Resume Next
        Else
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
            Exit Sub
        End If
    End Sub
End Class