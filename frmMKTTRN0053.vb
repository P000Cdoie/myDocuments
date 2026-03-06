Option Strict Off
Option Explicit On
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Friend Class frmMKTTRN0053
	Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
	' File Name         :   frmMKTTRN0053.frm
	' Function          :   Used to Print & View Packing List details
	' Created By        :   Manoj Kr. Vaish
	' Created On        :   02 May, 2007
	'===================================================================================
	'***********************************************************************************
	'Revised By      : Manoj Kr. Vaish
	'Issue ID        : 20739
	'Revision Date   : 31 July 2007
	'History         : To get the Packing List Detail for an Invoice for Packing List Report
    '***********************************************************************************
    '----------------------------------------------------
    'Modified by    :   Virendra Gupta
    'Modified ON    :   02/06/2011
    'Modified to support MultiUnit functionality
    '-----------------------------------------------------------------------
	
    Dim mresult As ClsResultSetDB
	Dim mintIndex As Short
    Dim WindowHnd As Integer
    Private Enum Col_Header
        paletteno = 1
        CustItemCode = 2
        CustItemDesc = 3
        Quantity = 4
        NoOfBox = 5
        PerPaletteWt = 6
        PerBoxWt = 7
        NetWeight = 8
    End Enum
    Private Structure PaletteDetail
        Dim paletteno As Short
        Dim PaletteWt As Decimal
    End Structure
    Dim marrPaletteDetail() As PaletteDetail
    Dim mCurInvoiceQuantity As Decimal
    Dim mblnRecordExist As Boolean
    Dim mStrAccountCode As String
    Dim mintTotalBox As Short
    Dim mintTotalPalette As Short
    Dim mcurGrossWeight As Decimal
    Dim mcurTotalNetWeight As Decimal
    Dim mcurTotalBoxWeight As Decimal
    Dim mcurTotalPaletteWt As Decimal
    Dim mblnvalidBeforeNewRow As Boolean
    Dim mblnZeroPalleteWt As Boolean
    Const MaxHdrGridCols As Short = 8
    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
        Dim strHelp As Object
        On Error GoTo Err_Handler
        With Me.txtInvoice
            fsitemList.MaxRows = 0
            If Len(Trim(txtLocationCode.Text)) = 0 Then
                MsgBox("Select Location Code First", MsgBoxStyle.Information, ResolveResString(100))
                CmdLocCodeHelp.Focus()
                Exit Sub
            End If
            If Len(Trim(.Text)) = 0 Then
                'calling help to display permanent invoice numbers
                strHelp = ShowList(1, .MaxLength, "", "Doc_No", DateColumnNameInShowList("Invoice_Date") & "As Invoice_Date", "SalesChallan_dtl", " and print_flag=1 and bill_flag=1 and location_code='" & Trim(txtLocationCode.Text) & "'")
                If Val(strHelp) = -1 Then ' No record found
                    'resource Message
                    Call ConfirmWindow(10512, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                    Call FillPaletteDetail()
                    .Focus()
                Else
                    .Text = strHelp 'displaying value in text box
                    Call FillPaletteDetail()
                End If
            Else
                'calling help to display permanent invoice numbers
                strHelp = ShowList(1, .MaxLength, .Text, "Doc_No", DateColumnNameInShowList("Invoice_Date") & "As Invoice_Date", "SalesChallan_dtl", " and print_flag=1 and bill_flag=1 and location_code='" & Trim(txtLocationCode.Text) & "'")
                If Val(strHelp) = -1 Then ' No record found
                    'resource Message
                    Call ConfirmWindow(10512, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                    Call FillPaletteDetail()
                    .Focus()
                Else
                    .Text = strHelp 'displaying value in text box
                    Call FillPaletteDetail()
                End If
            End If
        End With
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdInvoice_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As UCActXCtl.cmdGrpAuthorise.ButtonClickEventArgs) Handles Cmdinvoice.ButtonClick
        'Dim rssaledtl As ClsResultSetDB
        Dim Address As String
        Dim strsql_sel As String
        Dim strRptName As String
        Dim strQuantityCheck As String
        On Error GoTo Err_Handler
        If Cmdinvoice.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE Then
            Me.Close()
            Exit Sub
        ElseIf Len(Trim(txtInvoice.Text)) = 0 Then
            MsgBox("Select Invoice No.", MsgBoxStyle.Information, ResolveResString(100))
            txtInvoice.Focus()
            Exit Sub
        End If
        '<<<<CR11 Code Starts>>>>
        Dim objRpt As ReportDocument
        Dim frmReportViewer As New eMProCrystalReportViewer
        objRpt = frmReportViewer.GetReportDocument()
        frmReportViewer.ShowPrintButton = True
        frmReportViewer.ShowTextSearchButton = True
        frmReportViewer.ShowZoomButton = True
        frmReportViewer.ReportHeader = Me.ctlFormHeader1.HeaderString()
        '<<<<CR11 Code Ends>>>>
        strRptName = GetPlantName()
        strRptName = "\Reports\EXPORT_packing" & "_" & strRptName & ".rpt"
        If Not CheckFile(strRptName) Then
            strRptName = "\Reports\EXPORT_packing.rpt"
        End If
        With objRpt
            'load the report
            .Load(My.Application.Info.DirectoryPath & strRptName)
            .DataDefinition.FormulaFields("Comp_name").Text = "'" & gstrCOMPANY & "'"
            .DataDefinition.FormulaFields("comp_add").Text = "'" & gstr_WRK_ADDRESS1 & gstr_WRK_ADDRESS2 & "'"
        End With
        strsql_sel = "{SalesChallan_Dtl.Doc_No}=" & txtInvoice.Text & " and {SalesChallan_Dtl.Location_Code}='" & Trim(txtLocationCode.Text) & "' and {SalesChallan_Dtl.invoice_type}='EXP' and {SalesChallan_Dtl.sub_category}='E'"
        objRpt.RecordSelectionFormula = strsql_sel & " and {SalesChallan_Dtl.UNIT_CODE} = '" & gstrUNITID & "'"
        Select Case eventArgs.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_AUTHORIZE 'For Save Data
                strQuantityCheck = CheckQuantity()
                If Len(Trim(strQuantityCheck)) > 0 Then
                    If strQuantityCheck = "False" Then
                        Exit Sub
                    End If
                End If
                Call InsertPaletteDetail() 'Save Palette Detail
                Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                fsitemList.MaxRows = 0
                txtLocationCode.Text = ""
                txtInvoice.Text = ""
                Cmdinvoice.Enabled(0) = False
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE 'For Close Form
                Me.Close()
                Exit Sub
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_REFRESH 'For Print Priview
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                frmReportViewer.Show()
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
                frmReportViewer.SetReportDocument()
                objRpt.PrintToPrinter(1, False, 0, 0)
        End Select
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub CmdLocCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdLocCodeHelp.Click
        '*******************************************************************************
        'Author             :   Manoj Kr. Vaish
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Display Help From Location Master
        'Comments           :   NA
        'Creation Date      :   01-Aug-2007
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strHelp As String
        If Len(Trim(Me.txtLocationCode.Text)) = 0 Then 'To check if There is No Text Then Show All Help
            strHelp = ShowList(1, (txtLocationCode.MaxLength), "", "s.Location_Code", "l.Description", "Location_mst l,SaleConf s", "and s.Location_Code=l.Location_Code and s.Unit_Code = l.Unit_Code and s.Unit_Code = '" & gstrUNITID & "' and (s.fin_start_date <= getdate() and s.fin_end_date >= getdate())", , , , , , "s.Unit_Code")
            If strHelp = "-1" Then 'If No Record Exists In The Table
                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Exit Sub
            Else
                txtLocationCode.Text = strHelp
            End If
        Else
            'To Display All Possible Help Starting With Text in TextField
            strHelp = ShowList(1, (txtLocationCode.MaxLength), txtLocationCode.Text, "s.Location_Code", "l.Description", "Location_mst l,SaleConf s", "and s.Location_Code=l.Location_Code and s.Unit_Code=l.Unit_Code and (s.fin_start_date <= getdate() and s.fin_end_date >= getdate())", , , , , , "s.Unit_Code")
            If strHelp = "-1" Then 'If No Record Exists In The Table
                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Exit Sub
            Else
                txtLocationCode.Text = strHelp
            End If
        End If
        'Procedure Call To Select The Location Code Description
        Call SelectDescriptionForField("Description", "Location_Code", "Location_Mst", lblLocCodeDes, (txtLocationCode.Text))
        txtLocationCode.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub ctlFormHeader1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ctlFormHeader1.Click
        '-----------------------------------------------------------------------
        'Author              - Manoj Kr. Vaish
        'Create Date         - 02/05/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Show Empower Help on Control Header Button Click
        '-----------------------------------------------------------------------
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
    End Sub
    Private Sub frmMKTTRN0053_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '-----------------------------------------------------------------------
        'Author              - Manoj Kr. Vaish
        'Create Date         - 02/05/2007
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call empower help on F4 Click
        '-----------------------------------------------------------------------
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then Call ctlFormHeader1_ClickEvent(ctlFormHeader1, New System.EventArgs()) : Exit Sub
    End Sub
    Private Sub fsitemList_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles fsitemList.Enter
        ''-------------------------------------------------------------------------------------------------
        '' Author        : Manoj Kr. Vaish
        '' Arguments     : NIL
        '' Return Value  : NIL
        '' Function      : To set the focus on palette number cell
        '' Datetime      : 31-July-2007
        ''----------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        With fsitemList
            .Row = 1
            .Col = Col_Header.PerPaletteWt
            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
            .Focus()
        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub fsitemList_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles fsitemList.KeyPressEvent
        '-------------------------------------------------------------------------------------------
        ' Author        : Manoj Kr. Vaish
        ' Arguments     : NIL
        ' Return Value  : NIL
        ' Function      : Restrict keys
        ' Datetime      : 14 June 2007
        'Issue ID       : 19992
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Select Case eventArgs.keyAscii
            Case System.Windows.Forms.Keys.Up, System.Windows.Forms.Keys.Down
                eventArgs.keyAscii = 0
            Case 39, 34, 96, 45
                eventArgs.keyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub fsitemList_LeaveCell(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles fsitemList.LeaveCell
        '-------------------------------------------------------------------------------------------
        ' Author        : Manoj Kr. Vaish
        ' Arguments     : NIL
        ' Return Value  : NIL
        ' Function      : Show Customer Part Description on mouse click
        ' Datetime      : 01 Aug 2007
        'Issue ID       : 20739
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim varItemCode As Object
        Dim curPalleteWt As Decimal
        Dim curBoxWt As Decimal
        Dim strRetVal As String
        Dim rwCount As Integer
        Dim colCount As Integer
        With fsitemList
            If eventArgs.newRow = -1 Then Exit Sub
            rwCount = .ActiveRow
            colCount = .ActiveCol
            If Not ValidateRowData(rwCount, colCount) Then
                eventArgs.cancel = True
                .Row = .ActiveRow
                .Col = .ActiveCol
                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                .Focus()
            End If
            If mblnRecordExist = False Then
                With fsitemList
                    If .ActiveRow = 1 And .ActiveCol = Col_Header.PerPaletteWt Then
                        .Row = 1
                        .Col = Col_Header.PerPaletteWt : curPalleteWt = Val(.Value)
                        While .Row <= .MaxRows
                            Call .SetText(Col_Header.PerPaletteWt, .Row, CDbl(curPalleteWt))
                            .Row = .Row + 1
                        End While
                    ElseIf .ActiveRow = 1 And .ActiveCol = Col_Header.PerBoxWt Then
                        .Row = 1
                        .Col = Col_Header.PerBoxWt : curBoxWt = Val(.Value)
                        While .Row <= .MaxRows
                            Call .SetText(Col_Header.PerBoxWt, .Row, CDbl(curBoxWt))
                            .Row = .Row + 1
                        End While
                    End If
                End With
            End If
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtInvoice_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtInvoice.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo Err_Handler
        If KeyAscii = 13 Then
            Call txtinvoice_Validating(txtInvoice, New System.ComponentModel.CancelEventArgs(False))
            With fsitemList
                .Row = 1
                .Col = Col_Header.paletteno
                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                .Focus()
            End With
        End If
        If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
            KeyAscii = KeyAscii
        Else
            KeyAscii = 0
        End If
        GoTo EventExitSub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtInvoice_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtInvoice.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*******************************************************************************
        'Author             :   Manoj Kr. Vaish
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   01-Aug-2007
        '*******************************************************************************
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If cmdHelp.Enabled Then Call cmdHelp_Click(cmdHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtinvoice_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtInvoice.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim strSql As String
        On Error GoTo Err_Handler
        If Len(Trim(txtInvoice.Text)) = 0 Then
            '   MsgBox "Please Enter Invoive No.", vbInformation + vbSystemModal, "empower"
            txtInvoice.Focus()
            GoTo EventExitSub
        Else
            strSql = " Select Doc_no,account_code from saleschallan_dtl where doc_no= '" & txtInvoice.Text & "' and bill_flag=1 and print_flag=1 and Unit_Code = '" & gstrUNITID & "'"
            'mresult.ResultSetClose()
            mresult = New ClsResultSetDB
            mresult.GetResult(strSql)
            If Not (mresult.GetNoRows > 0) Then
                MsgBox("Invoice No Does not Exist", MsgBoxStyle.Information, ResolveResString(100))
                txtInvoice.Text = ""
                txtInvoice.Focus()
                GoTo EventExitSub
            Else
                mStrAccountCode = mresult.GetValue("Account_Code")
            End If
            mresult.ResultSetClose()
        End If
        GoTo EventExitSub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTTRN0053_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub frmMKTTRN0053_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo Err_Handler
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FitToClient(Me, fraInvoice, ctlFormHeader1, Cmdinvoice) 'To fit the form in the MDI
        Cmdinvoice.Caption(0) = "Save"
        Cmdinvoice.Caption(1) = "Preview"
        Cmdinvoice.Picture(0) = My.Resources.resEmpower.ico229.ToBitmap
        Cmdinvoice.Picture(1) = My.Resources.resEmpower.ico232.ToBitmap   '00rr00 not found in .net ?????
        Call AddGridHeaders() 'Add Headers on Grid
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0053_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo Err_Handler
        'Removing the form name from list
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        'Setting the corresponding node's tag
        frmModules.NodeFontBold(Tag) = False
        'Closing the recordset
        Me.Dispose()
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0053_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex
        CmdLocCodeHelp.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0053_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub AddBlankRow()
        ''-------------------------------------------------------------------------------------------------
        '' Author        : Manoj Kr. vaish
        '' Arguments     : NIL
        '' Return Value  : NIL
        '' Function      : To add a blank row in the Grid
        '' Datetime      : 31-July-2007
        ''--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim Intcounter As Short
        With fsitemList
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .set_RowHeight(.Row, 300)
            .Row = .MaxRows
            .Col = Col_Header.paletteno
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Lock = True
            .Col = Col_Header.CustItemCode
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Lock = True
            .Col = Col_Header.CustItemDesc
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Lock = True
            .Col = Col_Header.Quantity
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Lock = True
            .Col = Col_Header.NoOfBox
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Lock = True
            .Col = Col_Header.PerPaletteWt
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = 4
            .TypeFloatMin = 0.0#
            .TypeFloatMax = CDbl("99999999.9999")
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Text = "0.0000"
            .Col = Col_Header.PerBoxWt
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatDecimalPlaces = 4
            .TypeFloatMin = 0.0#
            .TypeFloatMax = CDbl("99999999.9999")
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Text = "0.0000"
            .Col = Col_Header.NetWeight
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Lock = True
        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub AddGridHeaders()
        ''-------------------------------------------------------------------------------------------------
        '' Author        : Manoj Kr. vaish
        '' Arguments     : NIL
        '' Return Value  : NIL
        '' Function      : To add a blank row in the Grid
        '' Datetime      : 31-July-2007
        ''--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        With fsitemList
            .MaxRows = 0
            .MaxCols = MaxHdrGridCols
            .Row = 0
            .Font = Me.Font
            .UnitType = FPSpreadADO.UnitTypeConstants.UnitTypeTwips
            .Col = Col_Header.paletteno : .Text = "Palette No." : .set_ColWidth(Col_Header.paletteno, 600)
            .Col = Col_Header.CustItemCode : .Text = "Cust Item Code" : .set_ColWidth(Col_Header.CustItemCode, 2000)
            .Col = Col_Header.CustItemDesc : .Text = "Cust Item Desc" : .set_ColWidth(Col_Header.CustItemDesc, 2700) : .UserResizeCol = FPSpreadADO.UserResizeConstants2.UserResizeOn
            .Col = Col_Header.Quantity : .Text = "Quantity" : .set_ColWidth(Col_Header.Quantity, 800)
            .Col = Col_Header.NoOfBox : .Text = "No.Of Box" : .set_ColWidth(Col_Header.NoOfBox, 800)
            .Col = Col_Header.PerPaletteWt : .Text = "Palette Wt/Unit(Kg)" : .set_ColWidth(Col_Header.PerPaletteWt, 1100)
            .Col = Col_Header.PerBoxWt : .Text = "Box Wt/Unit(Kg)" : .set_ColWidth(Col_Header.PerBoxWt, 900)
            .Col = Col_Header.NetWeight : .Text = "Net Weight(Kg)" : .set_ColWidth(Col_Header.NetWeight, 900)
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLocationCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLocationCode.TextChanged
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Set The Values of Related control on Change of Location Code
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If Len(Trim(txtLocationCode.Text)) = 0 Then
            fsitemList.MaxRows = 0
            txtInvoice.Text = ""
            lblLocCodeDes.Text = ""
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLocationCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLocationCode.Enter
        '*******************************************************************************
        'Author             :   Manoj Kr. Vaish
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   to Show Selected Text
        'Comments           :   NA
        'Creation Date      :   01-Aug-2007
        '*******************************************************************************
        On Error GoTo ErrHandler
        Me.txtLocationCode.SelectionStart = 0
        Me.txtLocationCode.SelectionLength = Len(Me.txtLocationCode.Text)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLocationCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLocationCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Manoj Kr. Vaish
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   01-Aug-2007
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                If Len(txtLocationCode.Text) > 0 Then
                    Call txtLocationCode_Validating(txtLocationCode, New System.ComponentModel.CancelEventArgs(False))
                Else
                    txtInvoice.Focus()
                End If
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtLocationCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtLocationCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*******************************************************************************
        'Author             :   Manoj Kr. Vaish
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   01-Aug-2007
        '*******************************************************************************
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdLocCodeHelp.Enabled Then Call CmdLocCodeHelp_Click(CmdLocCodeHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLocationCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtLocationCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*******************************************************************************
        'Author             :   Manoj Kr. Vaish
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   Check Validity Of Location Code In The Location_Mst
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strSql As String
        If Len(txtLocationCode.Text) > 0 Then
            If CheckExistanceOfFieldData((txtLocationCode.Text), "Location_Code", "SaleConf", "(fin_start_date <= getdate() and fin_end_date >= getdate())") Then
                txtInvoice.Focus()
            Else
                Call ConfirmWindow(10411, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtLocationCode.Text = ""
                txtLocationCode.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Function CheckExistanceOfFieldData(ByRef pstrFieldText As String, ByRef pstrColumnName As String, ByRef pstrTableName As String, Optional ByRef pstrCondition As String = "") As Boolean
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   pstrFieldText - Field Text,pstrColumnName - Column Name
        '                       pstrTableName - Table Name,pstrCondition - Optional Parameter For Condition
        'Return Value       :   NA
        'Function           :   To Check Validity Of Field Data Whethet it Exists In The
        '                       Database Or Not
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        CheckExistanceOfFieldData = False
        Dim strTableSql As String 'Declared To Make Select Query
        Dim rsExistData As ClsResultSetDB
        strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and Unit_Code = '" & gstrUNITID & "'"
        If Len(Trim(pstrCondition)) > 0 Then
            strTableSql = strTableSql & " AND " & pstrCondition
        End If
        rsExistData = New ClsResultSetDB
        rsExistData.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsExistData.GetNoRows > 0 Then
            CheckExistanceOfFieldData = True
        Else
            CheckExistanceOfFieldData = False
        End If
        rsExistData.ResultSetClose()
        rsExistData = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub SelectDescriptionForField(ByRef pstrFieldName1 As String, ByRef pstrFieldName2 As String, ByRef pstrTableName As String, ByRef pContrName As System.Windows.Forms.Control, ByRef pstrControlText As String)
        '*******************************************************************************
        'Author             :
        'Argument(s)if any  :   pstrFieldName1 - Field Name1,pstrFieldName2 - Field Name2,pstrTableName - Table Name
        '                       pContName - Name Of The Control where Caption Is To Be Set
        '                       pstrControlText - Field Text
        'Return Value       :   NA
        'Function           :   To Select The Field Description In The Description Labels
        'Comments           :   NA
        'Creation Date      :
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strDesSql As String 'Declared to make Select Query
        Dim rsDescription As ClsResultSetDB
        strDesSql = "Select " & Trim(pstrFieldName1) & " from " & Trim(pstrTableName) & " where " & Trim(pstrFieldName2) & "='" & Trim(pstrControlText) & "' and Unit_code = '" & gstrUNITID & "'"
        rsDescription = New ClsResultSetDB
        rsDescription.GetResult(strDesSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsDescription.GetNoRows > 0 Then
            pContrName.Text = rsDescription.GetValue(Trim(pstrFieldName1))
        End If
        rsDescription.ResultSetClose()
        rsDescription = Nothing
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function ValidateRowData(ByVal Row As Integer, Optional ByRef Col As Integer = 0) As Boolean
        '-------------------------------------------------------------------------------------------
        ' Author        : Manoj Kr. Vaish
        ' Arguments     : Active row and column
        ' Return Value  : Boolean
        ' Function      : Validate the grid while entering the Packing List Details
        ' Datetime      : 01 Aug 2007
        'Issue ID       : 20739
        '---------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim varBoxNo As Object
        ValidateRowData = True
        With fsitemList
            .Row = Row
            If .ActiveCol = Col_Header.PerPaletteWt Then
                .Row = Row
                .Col = Col_Header.PerPaletteWt
                If CheckPalleteWeight(.ActiveRow, (Col_Header.PerPaletteWt)) = False And mblnZeroPalleteWt = False Then
                    MsgBox("Pallete Wt. can't be zero", MsgBoxStyle.Information, ResolveResString(100))
                    ValidateRowData = False
                    Exit Function
                ElseIf CheckPalleteWeight(.ActiveRow, (Col_Header.PerPaletteWt)) = False And mblnZeroPalleteWt = True Then
                    MsgBox("Pallte Wt. can't be different for same pallete no.", MsgBoxStyle.Information, ResolveResString(100))
                    ValidateRowData = False
                    Exit Function
                End If
            ElseIf .ActiveCol = Col_Header.PerBoxWt Then
                .Row = Row
                .Col = Col_Header.PerBoxWt
                If Val(.Text) = 0 Then
                    MsgBox("Box Wt. can't be zero", MsgBoxStyle.Information, ResolveResString(100))
                    ValidateRowData = False
                    Exit Function
                End If
            Else
                ValidateRowData = True
                Exit Function
            End If
        End With
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub FillPaletteDetail()
        '-------------------------------------------------------------------------------------------
        ' Author        : Manoj Kr. Vaish
        ' Arguments     : NIL
        ' Return Value  : NIL
        ' Function      : Fill palette detail if exist
        ' Datetime      : 02 Aug 2007
        ' Issue ID      : 20739
        '--------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rspalette As ClsResultSetDB
        Dim strQry As String
        Dim intRowCount As Short
        Dim Intcounter As Short
        Dim Com As ADODB.Command
        Dim rspltdetail As ADODB.Recordset
        rspalette = New ClsResultSetDB
        If Len(Trim(txtInvoice.Text)) > 0 Then
            Call txtinvoice_Validating(txtInvoice, New System.ComponentModel.CancelEventArgs(False)) 'Validate selected invoice number
            strQry = "select a.*,b.* from MKT_PALETTE_DTL a,MKT_PALETTE_HDR b where b.doc_no=" & Trim(txtInvoice.Text) & " and a.doc_no=b.doc_no and a.Unit_Code = b.Unit_Code and  a.Unit_Code = '" & gstrUNITID & "'"
            Call rspalette.GetResult(strQry, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            fsitemList.MaxRows = 0
            mblnRecordExist = False
            If rspalette.GetNoRows > 0 Then
                intRowCount = rspalette.GetNoRows
                With fsitemList
                    rspalette.MoveFirst()
                    For Intcounter = 1 To intRowCount Step 1
                        Call AddBlankRow()
                        Call .SetText(Col_Header.paletteno, Intcounter, rspalette.GetValue("Pallete_No"))
                        Call .SetText(Col_Header.CustItemCode, Intcounter, rspalette.GetValue("Cust_Item_Code"))
                        Call .SetText(Col_Header.CustItemDesc, Intcounter, rspalette.GetValue("Cust_Item_Desc"))
                        Call .SetText(Col_Header.Quantity, Intcounter, System.Math.Round(Convert.ToDouble(rspalette.GetValue("Quantity")), 0))
                        Call .SetText(Col_Header.NoOfBox, Intcounter, rspalette.GetValue("Boxes"))
                        Call .SetText(Col_Header.PerPaletteWt, Intcounter, rspalette.GetValue("PaletteWt_PerPalette"))
                        Call .SetText(Col_Header.PerBoxWt, Intcounter, rspalette.GetValue("BoxWt_PerBox"))
                        Call .SetText(Col_Header.NetWeight, Intcounter, rspalette.GetValue("NetWt_PerBox"))
                        rspalette.MoveNext()
                    Next
                    mblnRecordExist = True
                    Cmdinvoice.Enabled(0) = True
                    Cmdinvoice.Enabled(1) = True
                    Cmdinvoice.Enabled(2) = True
                End With
            Else
                '******************Fill Pallete Detail from store proc*********************
                Com = New ADODB.Command
                rspltdetail = New ADODB.Recordset
                With Com
                    .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    .CommandText = "PRC_FILLPACKINGDETAIL"
                    .Parameters.Append(Com.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(gstrUNITID)))
                    .Parameters.Append(Com.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, Trim(txtInvoice.Text)))
                    .Parameters.Append(Com.CreateParameter("@LOC_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(txtLocationCode.Text)))
                    .Parameters.Append(Com.CreateParameter("@ERR", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 100))
                    .let_ActiveConnection(mP_Connection)
                    rspltdetail = .Execute
                End With
                If Len(Com.Parameters(3).Value) > 0 Then
                    MsgBox(Com.Parameters(3).Value, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    GoTo ErrHandler
                    Com = Nothing
                    Exit Sub
                End If
                With fsitemList
                    Do While Not rspltdetail.EOF
                        Call AddBlankRow()
                        Call .SetText(Col_Header.paletteno, .MaxRows, rspltdetail.Fields("palleteno"))
                        Call .SetText(Col_Header.CustItemCode, .MaxRows, rspltdetail.Fields("Item_Code"))
                        Call .SetText(Col_Header.CustItemDesc, .MaxRows, rspltdetail.Fields("Item_Desc"))
                        Call .SetText(Col_Header.Quantity, .MaxRows, CShort(rspltdetail.Fields("Quantity").Value))
                        Call .SetText(Col_Header.NoOfBox, .MaxRows, rspltdetail.Fields("no_box"))
                        Call .SetText(Col_Header.PerPaletteWt, .MaxRows, "")
                        Call .SetText(Col_Header.PerBoxWt, .MaxRows, "")
                        Call .SetText(Col_Header.NetWeight, .MaxRows, rspltdetail.Fields("netwt"))
                        rspltdetail.MoveNext()
                    Loop
                End With
                If rspltdetail.State = ADODB.ObjectStateEnum.adStateOpen Then rspltdetail.Close()
                rspltdetail = Nothing
                Com = Nothing
                mblnRecordExist = False
                Cmdinvoice.Enabled(0) = True
            End If
            rspalette.ResultSetClose()
            rspalette = Nothing
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function CheckQuantity() As String
        '-------------------------------------------------------------------------------------------
        ' Author        : Manoj Kr. Vaish
        ' Arguments     : NIL
        ' Return Value  : String
        ' Function      : Verify Invoice Quanity from Sales_Dtl table
        ' Datetime      : 02 Aug 2007
        ' Issue ID      : 20739
        '--------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim intCount As Short
        Dim VarDelete As Object
        Dim varItemCode As Object
        Dim strMSG As String
        Dim intcol As Integer
        With fsitemList
            '*********************Check Blank Row *****************************
            For intCount = 1 To .MaxRows
                For intcol = 1 To .MaxCols
                    .Col = intcol
                    .Row = intCount
                    If (.Col = Col_Header.CustItemCode) Then
                        If Len(Trim(.Text)) = 0 Then
                            CheckQuantity = "False"
                            Call ConfirmWindow(10316, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            .Row = intCount : .Col = intcol : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                            CheckQuantity = "False"
                            Exit Function
                        End If
                    End If
                    'Check for Pallete Wt
                    If .Col = Col_Header.PerPaletteWt Then
                        If CheckPalleteWeight(.Row, (Col_Header.PerPaletteWt)) = False And mblnZeroPalleteWt = False Then
                            Call ConfirmWindow(10419, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            .Row = intCount : .Col = intcol : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                            CheckQuantity = "False"
                            Exit Function
                        ElseIf CheckPalleteWeight(.Row, (Col_Header.PerPaletteWt)) = False And mblnZeroPalleteWt = True Then
                            MsgBox("Same Pallete should be of same pallete wt.", MsgBoxStyle.Information, ResolveResString(100))
                            .Row = intCount : .Col = intcol : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                            CheckQuantity = "False"
                            Exit Function
                        End If
                    End If
                    If (.Col = Col_Header.paletteno) Or (.Col = Col_Header.Quantity) Or (.Col = Col_Header.NoOfBox) Or (.Col = Col_Header.PerBoxWt) Then
                        If (Val(Trim(.Text)) = 0) Then
                            CheckQuantity = "False"
                            Call ConfirmWindow(10419, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            .Row = intCount : .Col = intcol : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                            CheckQuantity = "False"
                            Exit Function
                        End If
                    End If
                Next intcol
            Next intCount
        End With
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub InsertPaletteDetail()
        '-------------------------------------------------------------------------------------------
        ' Author        : Manoj Kr. Vaish
        ' Arguments     : NIL
        ' Return Value  : NIL
        ' Function      : Insert Palette Detail
        ' Datetime      : 02 Aug 2007
        ' Issue ID      : 20739
        '-------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strDelete As String
        Dim strinsertdtl As String
        Dim strinserthdr As String
        Dim Intcounter As Short
        Dim VarDelete As Object
        Dim varPalette As Object
        Dim varCustItemCode As Object
        Dim varcustpartdesc As Object
        Dim varQuantity As Object
        Dim varBoxNo As Object
        Dim varpalettewt As Object
        Dim varBoxWt As Object
        Dim varnetwt As Object
        Dim varItemCode As Object
        strinsertdtl = ""
        strinserthdr = ""
        strDelete = ""
        '******************************Delete Old Row If Exist************************************
        If mblnRecordExist = True Then
            strDelete = Trim(strDelete) & "delete from mkt_palette_hdr where doc_no='" & Trim(txtInvoice.Text) & "' AND UNIT_CODE = '" & gstrUNITID & "'"
            strDelete = strDelete & "delete from mkt_palette_dtl where doc_no='" & Trim(txtInvoice.Text) & "' AND UNIT_CODE = '" & gstrUNITID & "'"
        End If
        Call CalculateNetAndGrossWt() 'Calculate Net & Gross Wt
        '***********************Insert header Row************************************************
        strinserthdr = Trim(strinserthdr) & "insert into mkt_palette_hdr(Loc_no,doc_no,Account_Code"
        strinserthdr = strinserthdr & ",No_Of_Box,No_Of_Pallet,GrossWeight,NetWeight,"
        strinserthdr = strinserthdr & "Ent_UserId,Ent_Dt,Upd_Dt ,Unit_Code) values('" & Trim(txtLocationCode.Text) & "'"
        strinserthdr = strinserthdr & ",'" & Trim(txtInvoice.Text) & "','" & Trim(mStrAccountCode) & "'"
        strinserthdr = strinserthdr & "," & mintTotalBox & "," & mintTotalPalette & "," & mcurGrossWeight
        strinserthdr = strinserthdr & "," & mcurTotalNetWeight & ",'" & mP_User & "'," & " getdate(), getdate(), '" & gstrUNITID & "') "
        '******************************************************************************************
        '***********************Insert Detail Row*************************************************
        With fsitemList
            For Intcounter = 1 To .MaxRows Step 1
                varPalette = Nothing
                varCustItemCode = Nothing
                varcustpartdesc = Nothing
                varQuantity = Nothing
                varBoxNo = Nothing
                varpalettewt = Nothing
                varBoxWt = Nothing
                varnetwt = Nothing
                Call .GetText(Col_Header.paletteno, Intcounter, varPalette)
                Call .GetText(Col_Header.CustItemCode, Intcounter, varCustItemCode)
                Call .GetText(Col_Header.CustItemDesc, Intcounter, varcustpartdesc)
                Call .GetText(Col_Header.Quantity, Intcounter, varQuantity)
                Call .GetText(Col_Header.NoOfBox, Intcounter, varBoxNo)
                Call .GetText(Col_Header.PerPaletteWt, Intcounter, varpalettewt)
                Call .GetText(Col_Header.PerBoxWt, Intcounter, varBoxWt)
                Call .GetText(Col_Header.NetWeight, Intcounter, varnetwt)
                strinsertdtl = Trim(strinsertdtl) & "insert into mkt_palette_dtl(doc_no,Pallete_No"
                strinsertdtl = strinsertdtl & ",Cust_Item_Code,Cust_Item_Desc,Boxes,Quantity,NetWt_PerBox,"
                strinsertdtl = strinsertdtl & "PaletteWt_PerPalette,BoxWt_PerBox, Unit_Code) values('" & Trim(txtInvoice.Text) & "'," & varPalette
                strinsertdtl = strinsertdtl & ",'" & Trim(varCustItemCode) & "','" & Trim(varcustpartdesc) & "'"
                strinsertdtl = strinsertdtl & "," & CDec(varBoxNo) & "," & CDec(varQuantity) & "," & CDec(varnetwt)
                strinsertdtl = strinsertdtl & "," & CDec(varpalettewt) & "," & CDec(varBoxWt) & ", '" & gstrUNITID & "')" & vbCrLf
            Next
        End With
        mP_Connection.BeginTrans()
        If Len(strDelete) > 0 Then mP_Connection.Execute(strDelete, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        If Len(strinserthdr) > 0 Then mP_Connection.Execute(strinserthdr, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        If Len(strinsertdtl) > 0 Then mP_Connection.Execute(strinsertdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.CommitTrans()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub CalculateNetAndGrossWt()
        '-------------------------------------------------------------------------------------------
        ' Author        : Manoj Kr. Vaish
        ' Arguments     : NIL
        ' Return Value  : NIL
        ' Function      : Calculate Net & Gross Weight
        ' Datetime      : 02 Aug 2007
        ' Issue ID      : 20739
        '-------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim intCount As Short
        Dim varBoxNo As Object
        Dim varBoxWt As Object
        Dim varnetwt As Object
        Dim varPaletteNo As Object
        mintTotalBox = 0
        mcurGrossWeight = 0
        mcurTotalPaletteWt = 0
        mcurTotalNetWeight = 0
        mcurTotalBoxWeight = 0
        Call FillDistinctPaletteInArray()
        With fsitemList
            For intCount = 1 To .MaxRows
                .Row = intCount
                varBoxNo = Nothing
                varBoxWt = Nothing
                varnetwt = Nothing
                Call .GetText(Col_Header.NoOfBox, intCount, varBoxNo)
                Call .GetText(Col_Header.PerBoxWt, intCount, varBoxWt)
                Call .GetText(Col_Header.NetWeight, intCount, varnetwt)
                mintTotalBox = mintTotalBox + Val(varBoxNo)
                mcurTotalNetWeight = mcurTotalNetWeight + CDec(varnetwt)
                mcurTotalBoxWeight = mcurTotalBoxWeight + (Val(varBoxNo) * CDec(varBoxWt))
            Next
            If UBound(marrPaletteDetail) > 0 Then
                mintTotalPalette = UBound(marrPaletteDetail)
                For intCount = 1 To UBound(marrPaletteDetail) Step 1
                    mcurTotalPaletteWt = mcurTotalPaletteWt + marrPaletteDetail(intCount).PaletteWt
                Next
            End If
            mcurGrossWeight = mcurTotalNetWeight + mcurTotalPaletteWt + mcurTotalBoxWeight
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub FillDistinctPaletteInArray()
        '-------------------------------------------------------------------------------------------
        ' Author        : Manoj Kr. Vaish
        ' Arguments     : NIL
        ' Return Value  : NIL
        ' Function      : Fill the distinct palette & its weight in array
        ' Datetime      : 02 Aug 2007
        ' Issue ID      : 20739
        '---------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim intArrIndex As Short
        Dim intRow As Short
        Dim Intcounter As Short
        Dim blnExist As Boolean
        Dim varPaletteNo As Object
        Dim varpalettewt As Object
        blnExist = False
        With fsitemList
            ReDim marrPaletteDetail(0)
            For intRow = 1 To .MaxRows Step 1
                varPaletteNo = Nothing
                varpalettewt = Nothing
                Call .GetText(Col_Header.paletteno, intRow, varPaletteNo)
                Call .GetText(Col_Header.PerPaletteWt, intRow, varpalettewt)
                If UBound(marrPaletteDetail) > 0 Then
                    For Intcounter = 1 To UBound(marrPaletteDetail)
                        If StrComp(varPaletteNo, CStr(marrPaletteDetail(Intcounter).paletteno), CompareMethod.Text) = 0 Then
                            blnExist = True
                            Exit For
                        Else
                            blnExist = False
                        End If
                    Next
                End If
                If blnExist = False Or UBound(marrPaletteDetail) = 0 Then
                    ReDim Preserve marrPaletteDetail(UBound(marrPaletteDetail) + 1)
                    intArrIndex = UBound(marrPaletteDetail)
                    marrPaletteDetail(intArrIndex).paletteno = CShort(varPaletteNo)
                    marrPaletteDetail(intArrIndex).PaletteWt = CDec(varpalettewt)
                End If
            Next
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function CheckPalleteWeight(ByVal Row As Integer, Optional ByRef Col As Integer = 0) As Boolean
        '-------------------------------------------------------------------------------------------
        ' Author        : Manoj Kr. Vaish
        ' Arguments     : NIL
        ' Return Value  : Boolean
        ' Function      : check the palltete wt of the Pallete
        ' Datetime      : 03 Aug 2007
        ' Issue ID      : 20739
        '---------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim curTempQty As Decimal
        Dim varPalleteNo As Object
        Dim intRowCount As Short
        Dim varPalleteWt As Object
        Dim intCount As Short
        On Error GoTo ErrHandler
        intCount = 0
        CheckPalleteWeight = True
        mblnZeroPalleteWt = True
        With fsitemList
            .Row = Row
            .Col = Col_Header.PerPaletteWt
            If Row > 0 Then
                .Col = Col_Header.paletteno
                varPalleteNo = Trim(.Text)
                .Col = Col_Header.PerPaletteWt
                varPalleteWt = Trim(.Text)
                For intRowCount = 1 To .MaxRows
                    .Row = intRowCount
                    .Col = Col_Header.paletteno
                    If StrComp(varPalleteNo, Trim(.Text), CompareMethod.Text) = 0 Then
                        .Col = Col_Header.PerPaletteWt
                        intCount = intCount + 1
                        If Val(.Text) <> 0 Then
                            If Val(varPalleteWt) <> Val(.Text) And Val(varPalleteWt) > 0 Then
                                CheckPalleteWeight = False
                                Exit For
                            End If
                        ElseIf Val(varPalleteWt) = 0 And Val(.Text) = 0 And intCount = 1 Then
                            intCount = 1
                            Exit For
                        Else
                            CheckPalleteWeight = True
                        End If
                    End If
                Next
                If intCount = 1 And Val(varPalleteWt) = 0 Then
                    CheckPalleteWeight = False
                    mblnZeroPalleteWt = False
                End If
            End If
        End With
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
End Class