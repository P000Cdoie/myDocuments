Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Friend Class frmMKTTRN0033_mtm
    Inherits System.Windows.Forms.Form
    '===================================================================================
    ' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
    ' File Name         :   FRMMKTTRN0033.frm
    ' Function          :   Used to Range Printing of Bin card
    ' Created By        :   Arshad Ali
    ' Created On        :   25 May, 2004
    '---------------------------------------------------------------------------------------------------------
    ' Revised By        :   Davinder Singh
    ' Revision Date     :   05/09/2005
    ' Revision History  :   Code added according to issue ID:-14999 to print the bin cards when multiple items associated
    '                       with single invoice no.
    '---------------------------------------------------------------------------------------------------------
    ' Revised By        :   Davinder Singh
    ' Revision Date     :   31/03/2006
    ' Issue ID          :   17445
    ' Revision History  :   To print or not to print the Invoice No. prefix according to the PrintPrefix flag in Sales_Parameter
    '---------------------------------------------------------------------------------------------------------
    ' Revised By        :   Davinder Singh
    ' Revision Date     :   01/05/2006
    ' Issue ID          :   17707
    ' Revision History  :   To pick the Bin Qty. from Sales_dtl instead of Custitem_mst
    '---------------------------------------------------------------------------------------------------------
    ' Revision By       :   Davinder Singh
    ' Revision Date     :   09/06/2006
    ' Issue ID          :   18043
    ' Revision History  :   To provide the DTP on the form so that user can select the Date he wants to print on the Bin Card
    '---------------------------------------------------------------------------------------------------------
    ' Revision By       :   Davinder Singh
    ' Revision Date     :   28/07/2006
    ' Issue ID          :   18378
    ' Revision History  :   While printing Bin Cards System Hangs after printing some cards
    '---------------------------------------------------------------------------------------------------------

    Dim mintFormIndex As Double
    Dim mBoxQuantity As Double
    Dim mSalesQuantity As Double
    Dim mCtlHdrInvoiceNo As System.Windows.Forms.ColumnHeader
    Dim mCtlHdrInvoiceDate As System.Windows.Forms.ColumnHeader
    Dim mCtlHdrItemCode As System.Windows.Forms.ColumnHeader
    Dim mCtlHdrBinQty As System.Windows.Forms.ColumnHeader
    Dim mlvwInvoice As System.Windows.Forms.ListViewItem
    Dim mblnPrintDate As Boolean
    Dim mlngWaitingTime As Integer

    Private Sub cmdHelpInvoice_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelpInvoice.Click
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Show Help Form
        '----------------------------------------------------
        Dim varHelp As Object
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        'UPGRADE_WARNING: Couldn't resolve default property of object varHelp. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        varHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT Doc_no,Invoice_date FROM Saleschallan_dtl WHERE UNIT_CODE = '" & gstrUNITID & "'  AND Bill_flag =1 and Cancel_flag =0 and Location_code = '" & Trim(txtUnitCode.Text) & "' and Doc_no like '" & Trim(txtinvoice.Text) & "%'")
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(varHelp) <> -1 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object varHelp(0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If varHelp(0) <> "0" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object varHelp(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtInvoice.Text = Trim(varHelp(0))
                txtInvoice.Focus()
            Else
                MsgBox("No Record Available", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub cmdHelpInvoice2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelpInvoice2.Click
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Show Help Form
        '----------------------------------------------------
        Dim varHelp As Object
        On Error GoTo ErrHandler
        If Trim(txtInvoice.Text) = "" Then MsgBox("Select From Invoice No.", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, ResolveResString(100)) : txtInvoice.Focus() : Exit Sub
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        'UPGRADE_WARNING: Couldn't resolve default property of object varHelp. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        varHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT Doc_no,Invoice_date FROM Saleschallan_dtl WHERE UNIT_CODE = '" & gstrUNITID & "'  AND Bill_flag =1 and Cancel_flag =0 and doc_no >= " & Val(txtinvoice.Text) & " and Location_code = '" & Trim(txtUnitCode.Text) & "' and Doc_no like '" & Trim(txtInvoiceTo.Text) & "%'")
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(varHelp) <> -1 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object varHelp(0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If varHelp(0) <> "0" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object varHelp(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtInvoiceTo.Text = Trim(varHelp(0))
                txtInvoiceTo.Focus()
            Else
                MsgBox("No Record Available", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub cmdInvoice_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As UCActXCtl.UCfraRepCmd.ButtonClickEventArgs) Handles Cmdinvoice.ButtonClick
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To call Button Click Events
        '---------------------------------------------------------------------------------------------------------
        ' Revised By        :   Davinder Singh
        ' Revision Date     :   28/07/2006
        ' Issue ID          :   18378
        ' Revision History  :   1) Used the Sleep function to give the print command after specified interval
        '                       2) Put the unwanted Queries out of the loop
        '---------------------------------------------------------------------------------------------------------

        Dim NoOfLabels As Integer
        Dim mFrom As Short
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim strLable As String
        Dim strDate As String
        Dim strLocation As String
        Dim strDINo As String
        Dim strsql As String
        Dim strReportFileName As String
        Dim dblQuantity As Double
        Dim intCount As Integer
        Dim strLastInvNo As String
        Dim strInvNo As String
        Dim rsSalesChallnDtl As ClsResultSetDB

        On Error GoTo ErrHandler

        mFrom = 1

        If eventArgs.Button <> UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE Then
            If eventArgs.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_FILE Then
                MsgBox("Export is not available", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
        End If

        Dim intTotalselected As Short
        Select Case eventArgs.button

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_WINDOW

                'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

                intTotalselected = 0
                For intCount = 0 To lvwInvoice.Items.Count - 1
                    'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If lvwInvoice.Items.Item(intCount).Checked = True Then
                        intTotalselected = intTotalselected + 1
                    End If
                Next
                If intTotalselected > 1 Then
                    MsgBox("Please select only one Invoice at once to preview.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    Exit Sub
                End If

                If ValidatebeforeSave() = False Then
                    'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    Exit Sub
                End If

                strReportFileName = GetFileName()
                If InStr(1, strReportFileName, ".") > 0 Then
                    strReportFileName = Mid(strReportFileName, 1, InStr(1, strReportFileName, ".") - 1)
                End If
                If strReportFileName = "" Then
                    MsgBox("No Report File Name Defined in Sales Parameter Table.", MsgBoxStyle.Information, ResolveResString(100))
                    'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    Exit Sub
                Else
                    strReportFileName = My.Application.Info.DirectoryPath & "\Reports\" & strReportFileName & ".rpt"
                End If

                For intCount = 0 To lvwInvoice.Items.Count - 1
                    'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If lvwInvoice.Items.Item(intCount).Checked = True Then
                        Call RefreshValues()
                        'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                        'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                        If ValidateInvoice((lvwInvoice.Items.Item(intCount).Text), lvwInvoice.Items.Item(intCount).SubItems(2).Text) Then
                            rsSalesChallnDtl = New ClsResultSetDB
                            'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                            'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                            rsSalesChallnDtl.GetResult("SELECT A.INVOICE_DATE, B.SRVDINO, B.SRVLOCATION FROM SALESCHALLAN_DTL A INNER JOIN SALES_DTL B ON A.DOC_NO=B.DOC_NO AND A.UNIT_CODE = B.UNIT_CODE AND A.UNIT_CODE = '" & gstrUNITID & "' where a.doc_no = " & lvwInvoice.Items.Item(intCount).Text & " And B.cust_item_code='" & lvwInvoice.Items.Item(intCount).SubItems(2).Text & "'")

                            If rsSalesChallnDtl.GetNoRows > 0 Then
                                rsSalesChallnDtl.MoveFirst()
                                'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesChallnDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                strDate = rsSalesChallnDtl.GetValue("Invoice_date")
                                'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesChallnDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                strDINo = rsSalesChallnDtl.GetValue("SRVDINO")
                                'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesChallnDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                strLocation = rsSalesChallnDtl.GetValue("SRVLocation")
                            End If
                            rsSalesChallnDtl.ResultSetClose()
                            'UPGRADE_NOTE: Object rsSalesChallnDtl may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                            rsSalesChallnDtl = Nothing
                            'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                            'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                            strsql = "{Sales_Dtl.UNIT_CODE} = '" & gstrUNITID & "' and {Sales_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {Sales_Dtl.Doc_No} =" & Trim(lvwInvoice.Items.Item(intCount).Text) & " and {Sales_Dtl.cust_item_code}='" & Trim(lvwInvoice.Items.Item(intCount).SubItems(2).Text) & "'"

                            NoOfLabels = Fix(mSalesQuantity / mBoxQuantity)
                            If (mSalesQuantity - (NoOfLabels * mBoxQuantity)) > 0 Then
                                NoOfLabels = NoOfLabels + 1
                            End If

                            Dim objRpt As ReportDocument

                            Dim frmReportViewer As New eMProCrystalReportViewer
                            objRpt = frmReportViewer.GetReportDocument()
                            frmReportViewer.ShowPrintButton = True
                            frmReportViewer.ShowTextSearchButton = True
                            frmReportViewer.ShowZoomButton = True
                            frmReportViewer.ReportHeader = Me.ctlFormHeader1.HeaderString()

                            With objRpt
                                'load the report
                                .Load(strReportFileName)

                                strLable = mFrom & " OF " & NoOfLabels

                                .DataDefinition.FormulaFields("PageCount").Text = "'" & strLable & "'"
                                .DataDefinition.FormulaFields("InvoiceNo").Text = "'" & Print_InvoiceNumber(lvwInvoice.Items.Item(intCount).Text) & "'"
                                .DataDefinition.FormulaFields("InvoiceDate").Text = "'" & strDate & "'"
                                .DataDefinition.FormulaFields("SRVLocation").Text = "'" & strLocation & "'"
                                .DataDefinition.FormulaFields("SRVDINO").Text = "'" & strDINo & "'"
                                .DataDefinition.FormulaFields("NoofBin").Text = "'" & NoOfLabels & "'"
                                .DataDefinition.FormulaFields("BoxQuantity").Text = "'" & mBoxQuantity & "'"

                                If mblnPrintDate Then
                                    .DataDefinition.FormulaFields("DateofPrinting").Text = "'" & VB6.Format(DTPprintdate.Value, gstrDateFormat) & "'"
                                End If

                                .RecordSelectionFormula = strsql & " AND {Sales_Dtl.UNIT_CODE} = '" & gstrUNITID & "'"

                                frmReportViewer.Show()

                                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                                Exit Sub
                            End With

                            'With rptbinPrinting
                            '    .Reset()
                            '    .DiscardSavedData = True
                            '    .WindowState = 2
                            '    .WindowShowPrintBtn = True
                            '    .WindowShowRefreshBtn = True
                            '    .WindowShowExportBtn = True
                            '    .WindowShowPrintSetupBtn = True
                            '    .WindowShowZoomCtl = True
                            '    .WindowShowCloseBtn = True
                            '    .WindowParentHandle = mdifrmMain.Handle.ToInt32
                            '    .WindowState = 2
                            '    .WindowTitle = ctlFormHeader1.HeaderString()
                            '    .Connect = gstrREPORTCONNECT
                            '    strLable = mFrom & " OF " & NoOfLabels
                            '    .set_Formulas(0, "PageCount = '" & strLable & "'")
                            '    'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                            '    .set_Formulas(1, "InvoiceNo = '" & Print_InvoiceNumber(lvwInvoice.Items.Item(intCount).Text) & "'")
                            '    .set_Formulas(2, "InvoiceDate = '" & strDate & "'")
                            '    .set_Formulas(3, "SRVLocation = '" & strLocation & "'")
                            '    .set_Formulas(4, "SRVDINO = '" & strDINo & "'")
                            '    .set_Formulas(5, "NoofBin = '" & NoOfLabels & "'")
                            '    .set_Formulas(6, "BoxQuantity = '" & mBoxQuantity & "'")
                            '    If mblnPrintDate Then
                            '        .set_Formulas(7, "DateofPrinting = '" & VB6.Format(DTPprintdate.Value, "dd/mm/yyyy") & "'")
                            '    End If
                            '    '.set_Formulas(8, "SF_LocCode = '" & Trim(txtUnitCode.Text) & "'")
                            '    '.set_Formulas(9, "SF_DocNo = " & Trim(lvwInvoice.Items.Item(intCount).Text))
                            '    '.set_Formulas(10, "SF_ItemCode = '" & Trim(lvwInvoice.Items.Item(intCount).SubItems(2).Text) & "'")
                            '    '.set_Formulas(11, "SF_UnitCode = '" & gstrUNITID & "'")
                            '    .ReportFileName = strReportFileName
                            '    .SelectionFormula = strsql
                            '    '.Destination = Crystal.DestinationConstants.crptToWindow
                            '    .Destination = 0
                            '    .Action = 1
                            '    'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                            '    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                            '    Exit Sub
                            'End With
                        End If
                    End If
                Next

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_PRINTER

                'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

                If ValidatebeforeSave() = False Then
                    'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    Exit Sub
                End If

                strReportFileName = GetFileName()
                If InStr(1, strReportFileName, ".") > 0 Then
                    strReportFileName = Mid(strReportFileName, 1, InStr(1, strReportFileName, ".") - 1)
                End If

                If strReportFileName = "" Then
                    MsgBox("No Report File Name Defined in Sales Parameter Table.", MsgBoxStyle.Information, ResolveResString(100))
                    'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    Exit Sub
                Else
                    strReportFileName = My.Application.Info.DirectoryPath & "\Reports\" & strReportFileName & ".rpt"
                End If

                For intCount = 0 To lvwInvoice.Items.Count - 1
                    'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If lvwInvoice.Items.Item(intCount).Checked = True Then

                        Call RefreshValues()

                        'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                        If StrComp(strLastInvNo, Trim(lvwInvoice.Items.Item(intCount).Text)) <> 0 Then
                            'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                            strLastInvNo = Trim(lvwInvoice.Items.Item(intCount).Text)
                            'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                            strInvNo = Print_InvoiceNumber(lvwInvoice.Items.Item(intCount).Text)
                        End If

                        'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                        'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                        If ValidateInvoice((lvwInvoice.Items.Item(intCount).Text), lvwInvoice.Items.Item(intCount).SubItems(2).Text) = False Then
                            'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                            Exit Sub
                        End If

                        rsSalesChallnDtl = New ClsResultSetDB
                        'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                        'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                        rsSalesChallnDtl.GetResult("SELECT A.INVOICE_DATE, B.SRVDINO, B.SRVLOCATION FROM SALESCHALLAN_DTL A INNER JOIN SALES_DTL B ON A.DOC_NO=B.DOC_NO AND A.UNIT_CODE = B.UNIT_CODE AND A.UNIT_CODE = '" & gstrUNITID & "' where a.doc_no = " & lvwInvoice.Items.Item(intCount).Text & " And B.cust_item_code='" & lvwInvoice.Items.Item(intCount).SubItems(2).Text & "'")

                        If rsSalesChallnDtl.GetNoRows > 0 Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesChallnDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            strDate = rsSalesChallnDtl.GetValue("Invoice_date")
                            'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesChallnDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            strDINo = rsSalesChallnDtl.GetValue("SRVDINO")
                            'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesChallnDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            strLocation = rsSalesChallnDtl.GetValue("SRVLocation")
                        End If
                        rsSalesChallnDtl.ResultSetClose()
                        'UPGRADE_NOTE: Object rsSalesChallnDtl may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                        rsSalesChallnDtl = Nothing

                        'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                        'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                        strsql = "{Sales_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {Sales_Dtl.Doc_No} =" & Trim(lvwInvoice.Items.Item(intCount).Text) & " and {Sales_Dtl.cust_item_code} ='" & Trim(lvwInvoice.Items.Item(intCount).SubItems(2).Text) & "' AND {Sales_Dtl.UNIT_CODE} = '" & gstrUNITID & "'"

                        NoOfLabels = Fix(mSalesQuantity / mBoxQuantity)

                        If (mSalesQuantity - (NoOfLabels * mBoxQuantity)) > 0 Then
                            NoOfLabels = NoOfLabels + 1
                        End If

                        'With rptbinPrinting
                        '    .Reset()
                        '    .DiscardSavedData = True
                        '    .Connect = gstrREPORTCONNECT
                        '    .ReportFileName = strReportFileName
                        '    .SelectionFormula = strsql
                        '    .set_Formulas(1, "InvoiceNo = '" & strInvNo & "'")
                        '    .set_Formulas(2, "InvoiceDate = '" & strDate & "'")
                        '    .set_Formulas(3, "SRVLocation = '" & strLocation & "'")
                        '    .set_Formulas(4, "SRVDINO = '" & strDINo & "'")
                        '    .set_Formulas(5, "NoofBin = '" & NoOfLabels & "'")
                        '    .set_Formulas(6, "BoxQuantity = '" & mBoxQuantity & "'")
                        '    If mblnPrintDate Then
                        '        .set_Formulas(7, "DateofPrinting = '" & VB6.Format(DTPprintdate.Value, "dd/mm/yyyy") & "'")
                        '    End If
                        '    '.set_Formulas(8, "SF_LocCode = '" & Trim(txtUnitCode.Text) & "'")
                        '    '.set_Formulas(9, "SF_DocNo = " & Trim(lvwInvoice.Items.Item(intCount).Text))
                        '    '.set_Formulas(10, "SF_ItemCode = '" & Trim(lvwInvoice.Items.Item(intCount).SubItems(2).Text) & "'")
                        '    '.set_Formulas(11, "SF_UnitCode = '" & gstrUNITID & "'")
                        '    .DiscardSavedData = True
                        '    intMaxLoop = NoOfLabels
                        '    For intLoopCounter = mFrom To intMaxLoop
                        '        strLable = intLoopCounter & " OF " & NoOfLabels
                        '        .set_Formulas(0, "PageCount = '" & strLable & "'")
                        '        If intLoopCounter = intMaxLoop Then
                        '            dblQuantity = mSalesQuantity - (intLoopCounter - 1) * mBoxQuantity
                        '            .set_Formulas(6, "BoxQuantity = '" & CStr(dblQuantity) & "'")
                        '        Else
                        '            .set_Formulas(6, "BoxQuantity = '" & CStr(mBoxQuantity) & "'")
                        '        End If

                        '        .Destination = 1
                        '        .Action = 1
                        '        Sleep((mlngWaitingTime))
                        '    Next
                        'End With
                        Dim objRpt As ReportDocument
                        Dim frmReportViewer As New eMProCrystalReportViewer
                        objRpt = frmReportViewer.GetReportDocument()
                        frmReportViewer.ShowPrintButton = True
                        frmReportViewer.ShowTextSearchButton = True
                        frmReportViewer.ShowZoomButton = True
                        frmReportViewer.ReportHeader = Me.ctlFormHeader1.HeaderString()
                        '<<<<CR11 Code Ends>>>>

                        With objRpt
                            'LOAD THE REPORT
                            .Load(strReportFileName)

                            .DataDefinition.FormulaFields("InvoiceNo").Text = "'" & strInvNo & "'"
                            .DataDefinition.FormulaFields("InvoiceDate").Text = "'" & strDate & "'"
                            .DataDefinition.FormulaFields("SRVLocation").Text = "'" & strLocation & "'"
                            .DataDefinition.FormulaFields("SRVDINO").Text = "'" & strDINo & "'"
                            .DataDefinition.FormulaFields("NoofBin").Text = "'" & NoOfLabels & "'"
                            .DataDefinition.FormulaFields("BoxQuantity").Text = "'" & mBoxQuantity & "'"

                            If mblnPrintDate Then
                                .DataDefinition.FormulaFields("DateofPrinting").Text = "'" & VB6.Format(DTPprintdate.Value, gstrDateFormat) & "'"
                            End If
                            .RecordSelectionFormula = strsql & " AND {Sales_Dtl.UNIT_CODE} = '" & gstrUNITID & "'"
                            intMaxLoop = NoOfLabels
                            For intLoopCounter = mFrom To intMaxLoop
                                strLable = intLoopCounter & " OF " & NoOfLabels
                                .DataDefinition.FormulaFields("PageCount").Text = "'" & strLable & "'"
                                If intLoopCounter = intMaxLoop Then
                                    dblQuantity = mSalesQuantity - (intLoopCounter - 1) * mBoxQuantity
                                    .DataDefinition.FormulaFields("BoxQuantity").Text = "'" & CStr(dblQuantity) & "'"
                                Else
                                    .DataDefinition.FormulaFields("BoxQuantity").Text = "'" & CStr(mBoxQuantity) & "'"
                                End If
                                frmReportViewer.SetReportDocument()
                                objRpt.PrintToPrinter(1, False, 0, 0)

                                Sleep((mlngWaitingTime))
                            Next
                        End With
                    End If
                Next
        End Select
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Exit Sub
ErrHandler:
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdUnitCodeList_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUnitCodeList.Click
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Show Help Form
        '----------------------------------------------------
        Dim varHelp As Object
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        'UPGRADE_WARNING: Couldn't resolve default property of object varHelp. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        varHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT Unt_CodeID,unt_unitname FROM Gen_UnitMaster WHERE Unt_Status=1 AND Unt_CodeID = '" & gstrUNITID & "'")
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(varHelp) <> -1 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object varHelp(0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If varHelp(0) <> "0" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object varHelp(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                txtUnitCode.Text = Trim(varHelp(0))
                txtUnitCode.Focus()
            Else
                MsgBox("No Record Available", MsgBoxStyle.Information, ResolveResString(100))
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub frmMKTTRN0033_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape And Shift = 0 Then Me.Close()
    End Sub

    Private Sub lvwInvoice_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lvwInvoice.Click
        'Dim intCount As Integer
        'If OptAll.value = True Then
        '    For intCount = 1 To lvwInvoice.ListItems.count
        '        If lvwInvoice.ListItems.Item(intCount).Checked = False Then
        '            lvwInvoice.ListItems.Item(intCount).Checked = True
        '        End If
        '    Next
        'End If
    End Sub

    'UPGRADE_WARNING: Event optAll.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optAll_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAll.CheckedChanged
        If eventSender.Checked Then
            '----------------------------------------------------
            'Author              - Arshad Ali
            'Create Date         - 25/05/2003
            'Arguments           - None
            'Return Value        - None
            '----------------------------------------------------
            On Error GoTo ErrHandler
            Dim intCount As Short
            If OptAll.Checked = True Then
                For intCount = 0 To lvwInvoice.Items.Count - 1
                    'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If lvwInvoice.Items.Item(intCount).Checked = False Then
                        'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                        lvwInvoice.Items.Item(intCount).Checked = True
                    End If
                Next
            Else
                For intCount = 0 To lvwInvoice.Items.Count - 1
                    'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If lvwInvoice.Items.Item(intCount).Checked = True Then
                        'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                        lvwInvoice.Items.Item(intCount).Checked = False
                    End If
                Next
            End If
            Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    'UPGRADE_WARNING: Event optSelected.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optSelected_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSelected.CheckedChanged
        If eventSender.Checked Then
            '----------------------------------------------------
            'Author              - Arshad Ali
            'Create Date         - 25/05/2003
            'Arguments           - None
            'Return Value        - None
            '----------------------------------------------------
            On Error GoTo ErrHandler
            Dim intCount As Short
            If OptAll.Checked = True Then
                For intCount = 0 To lvwInvoice.Items.Count - 1
                    'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If lvwInvoice.Items.Item(intCount).Checked = False Then
                        'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                        lvwInvoice.Items.Item(intCount).Checked = True
                    End If
                Next
            Else
                For intCount = 0 To lvwInvoice.Items.Count - 1
                    'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If lvwInvoice.Items.Item(intCount).Checked = True Then
                        'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                        lvwInvoice.Items.Item(intCount).Checked = False
                    End If
                Next
            End If
            Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub

    Private Sub txtInvoice_Change(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles txtInvoice.Change
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Clear Related Data
        '----------------------------------------------------
        txtInvoice.Text = Replace(txtInvoice.Text, "'", "")
        txtInvoiceTo.Text = ""
        lvwInvoice.Items.Clear()
    End Sub

    Private Sub txtInvoice_KeyDown(ByVal Sender As System.Object, ByVal e As CtlGeneral.KeyDownEventArgs) Handles txtInvoice.KeyDown
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.Shift
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            cmdHelpInvoice_Click(cmdHelpInvoice, New System.EventArgs())
        End If
    End Sub

    Private Sub txtInvoiceTo_Change(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles txtInvoiceTo.Change
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            -
        '----------------------------------------------------
        On Error GoTo ErrHandler
        txtInvoiceTo.Text = Replace(txtInvoiceTo.Text, "'", "")
        lvwInvoice.Items.Clear()
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtInvoiceTo_KeyDown(ByVal Sender As System.Object, ByVal e As CtlGeneral.KeyDownEventArgs) Handles txtInvoiceTo.KeyDown
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.Shift
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            cmdHelpInvoice2_Click(cmdHelpInvoice2, New System.EventArgs())
        End If

    End Sub

    Private Sub txtInvoiceTo_KeyPress(ByVal Sender As System.Object, ByVal e As CtlGeneral.KeyPressEventArgs) Handles txtInvoiceTo.KeyPress
        Dim KeyAscii As Short = e.KeyAscii
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To set focus on next control
        '----------------------------------------------------
        On Error GoTo Err_Handler
        If KeyAscii = 13 Then
            OptSelected.Checked = True
            Call FillInvoicesToList()
            System.Windows.Forms.SendKeys.Send(vbTab)
        End If
        If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
            KeyAscii = KeyAscii
        Else
            KeyAscii = 0
        End If
        DirectCast(Sender, CtlGeneral).KeyPressKeyascii = KeyAscii
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtInvoiceTo_LostFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtInvoiceTo.LostFocus
        Call FillInvoicesToList()
    End Sub

    'UPGRADE_WARNING: Event txtUnitCode.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtUnitCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnitCode.TextChanged
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To cleare Date
        '----------------------------------------------------
        On Error GoTo ErrHandler
        txtUnitCode.Text = Replace(txtUnitCode.Text, "'", "")
        txtInvoice.Text = ""
        txtInvoiceTo.Text = ""
        lvwInvoice.Items.Clear()
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtUnitCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnitCode.Enter
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Selected Set Focus
        '----------------------------------------------------
        On Error GoTo ErrHandler
        'Selecting the text on focus
        With txtUnitCode
            .SelectionStart = 0 : .SelectionLength = Len(Trim(.Text))
        End With
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtUnitCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtUnitCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call help on F1 Press
        '----------------------------------------------------
        On Error GoTo ErrHandler
        'If Ctrl/Alt/Shift is also pressed
        If Shift <> 0 Then Exit Sub
        'Show the help form when user pressed F1
        If KeyCode = System.Windows.Forms.Keys.F1 Then Call cmdUnitCodeList_Click(cmdUnitCodeList, New System.EventArgs())
        Exit Sub 'This is to avoid the execution of the error handler
        'UPGRADE_ISSUE: Constant vbEnter was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
        If KeyCode = Keys.Enter Then System.Windows.Forms.SendKeys.Send("{TAB}")
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtUnitCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtUnitCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Set Focus on Next Control
        '----------------------------------------------------
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{TAB}")
            'Supressing ¬ ¤ ¦ » characters since these are being used as string delimiters
        ElseIf KeyAscii = 187 Or KeyAscii = 166 Or KeyAscii = 164 Or KeyAscii = 172 Or KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 96 Then
            KeyAscii = 0
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtUnitCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtUnitCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Validate selected unit code
        '----------------------------------------------------
        On Error GoTo ErrHandler
        Dim strUnitDesc As String
        Dim mobjGLTrans As New prj_GLTransactions.cls_GLTransactions(gstrUNITID, GetServerDate)
        'Populate the details

        If Trim(txtUnitCode.Text) = "" Then GoTo EventExitSub
        strUnitDesc = mobjGLTrans.GetUnit(Trim(txtUnitCode.Text), ConnectionString:=gstrCONNECTIONSTRING)

        If CheckString(strUnitDesc) <> "Y" Then
            MsgBox(CheckString(strUnitDesc), MsgBoxStyle.Critical, ResolveResString(100))
            txtUnitCode.Text = ""
            Cancel = True
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub ctlFormHeader1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ctlFormHeader1.Click
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call Empower help
        '----------------------------------------------------
        MsgBox("No Help Attached to This Form", MsgBoxStyle.Information, ResolveResString(100))
    End Sub
    Private Sub txtInvoice_KeyPress(ByVal Sender As System.Object, ByVal e As CtlGeneral.KeyPressEventArgs) Handles txtInvoice.KeyPress
        Dim KeyAscii As Short = e.KeyAscii
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To set focus on next control
        '----------------------------------------------------
        On Error GoTo Err_Handler
        If KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send(vbTab)
        End If
        If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
            KeyAscii = KeyAscii
        Else
            KeyAscii = 0
        End If
        DirectCast(Sender, CtlGeneral).KeyPressKeyascii = KeyAscii
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtInvoice_KeyUp(ByVal Sender As System.Object, ByVal e As CtlGeneral.KeyUpEventArgs) Handles txtInvoice.KeyUp
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.Shift
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Set focus on Next Control
        '----------------------------------------------------
        On Error GoTo Err_Handler
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    'UPGRADE_WARNING: Form event frmMKTTRN0033.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    Private Sub frmMKTTRN0033_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
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
    'UPGRADE_WARNING: Form event frmMKTTRN0033.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    Private Sub frmMKTTRN0033_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
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
    Private Sub frmMKTTRN0033_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To display empower help on click of F4 in empower
        '----------------------------------------------------
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then Call ctlFormHeader1_ClickEvent(ctlFormHeader1, New System.EventArgs()) : Exit Sub
    End Sub
    Private Sub frmMKTTRN0033_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To initialise required data
        '----------------------------------------------------
        On Error GoTo Err_Handler
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        'Call FillLabelFromResFile(Me)   'To Fill label description from Resource file
        Call FitToClient(Me, fraInvoice, ctlFormHeader1, Cmdinvoice, 500) 'To fit the form in the MDI
        ''----Changes made By Davinder on 05-09-2005 according to issue ID:-14999 to select the selected option button by default
        OptAll.Checked = False : OptSelected.Checked = True

        ''----Added by Davinder on 09/06/2006 (Issue ID:-18043 To provide the DTP on the form so that user can select the date he wants to print on the Bin Card)
        DTPprintdate.value = GetServerDate
        cmdUnitCodeList.Image = My.Resources.ico111.ToBitmap
        cmdHelpInvoice.Image = My.Resources.ico111.ToBitmap
        cmdHelpInvoice2.Image = My.Resources.ico111.ToBitmap
        ''----Changes by Davinder end's here

        Call AddColumnsInListView()
        Call ReadSalesParameter()

        gblnCancelUnload = False
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0033_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
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
        'UPGRADE_NOTE: Object frmMKTTRN0033 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        'Me = Nothing
        Me.Dispose()
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function ValidatebeforeSave() As Boolean
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - Boolean
        'Function            - To Check Valid Feilds Value
        '-----------------------------------------------------------------------
        ' Revision By       :Davinder Singh
        ' Revision Date     :22/03/2006
        ' Issue ID          :17707
        ' Revision History  :To pick the Bin Qty. from Sales_dtl instead of Custitem_mst
        '-----------------------------------------------------------------------
        Dim strErrMsg As String
        Dim ctlBlank As System.Windows.Forms.Control
        Dim lNo As Integer
        Dim blnInvalidData As Boolean
        Dim intCount As Integer
        Dim blnSelected As Boolean
        On Error GoTo Err_Handler

        ValidatebeforeSave = True
        lNo = 1
        strErrMsg = ResolveResString(10059) & vbCrLf

        If Trim(txtUnitCode.Text) = "" Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Location Code"
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = txtUnitCode
        End If

        For intCount = 0 To lvwInvoice.Items.Count - 1
            'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
            If lvwInvoice.Items.Item(intCount).Checked Then
                blnSelected = True
                'UPGRADE_WARNING: Lower bound of collection lvwInvoice.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                If Trim(lvwInvoice.Items.Item(intCount).Text) = "" Then
                    blnInvalidData = True
                    strErrMsg = strErrMsg & vbCrLf & lNo & "." & ResolveResString(60373)
                    lNo = lNo + 1
                    If ctlBlank Is Nothing Then ctlBlank = lvwInvoice
                End If
            End If
        Next

        If Not blnSelected Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & ". No Invoice Selected."
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = lvwInvoice
        End If

        strErrMsg = VB.Left(strErrMsg, Len(strErrMsg) - 1)
        strErrMsg = strErrMsg & "."
        lNo = lNo + 1

        If blnInvalidData = True Then
            ValidatebeforeSave = False
            gblnCancelUnload = True
            Call MsgBox(strErrMsg, MsgBoxStyle.Information, ResolveResString(100))
            ctlBlank.Focus()
            Exit Function
        End If

        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub RefreshValues()
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To intialise Values
        '----------------------------------------------------
        mBoxQuantity = 0
        mSalesQuantity = 0
    End Sub

    Private Sub AddColumnsInListView()
        '***********************************
        'To add Columns Headers in the ListView in the form load
        '***********************************
        '----Changes made By Davinder on 05-09-2005 according to issue ID:-14999 by adding the code to add a third column named 'Item Code' in the list view
        On Error GoTo ErrHandler
        With Me.lvwInvoice
            'UPGRADE_WARNING: Add method in .Net does not produce a value Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C676271D-1A65-46CA-8403-027613C5E565"'
            mCtlHdrInvoiceNo = .Columns.Add("")
            mCtlHdrInvoiceNo.Text = "Invoice No"
            mCtlHdrInvoiceNo.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwInvoice.Width) / 2 - 1000)
            'UPGRADE_WARNING: Add method in .Net does not produce a value Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C676271D-1A65-46CA-8403-027613C5E565"'
            mCtlHdrInvoiceDate = .Columns.Add("")
            mCtlHdrInvoiceDate.Text = "Invoice Date"
            mCtlHdrInvoiceDate.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwInvoice.Width) / 2 - 1200)
            'UPGRADE_WARNING: Add method in .Net does not produce a value Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="C676271D-1A65-46CA-8403-027613C5E565"'
            mCtlHdrItemCode = .Columns.Add("")
            mCtlHdrItemCode.Text = "Item Code"
            mCtlHdrItemCode.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwInvoice.Width) / 2 - 300)
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Public Sub FillInvoicesToList()
        '----------------------------------------------------
        'Author              - Arshad Ali
        'Create Date         - 25/05/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Fill Invoices into the Listview
        '----------------------------------------------------
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim rsInvoice As New ClsResultSetDB
        If Trim(txtInvoice.Text) <> "" And Trim(txtInvoiceTo.Text) <> "" Then
            strsql = "SELECT a.Doc_no,a.Invoice_date,b.cust_item_code FROM Saleschallan_dtl a,sales_dtl b WHERE a.unit_code = b.unit_code and a.unit_code = '" & gstrUNITID & "' and a.Bill_flag =1 and a.Cancel_flag =0 and a.Location_code = '" & Trim(txtUnitCode.Text) & "' and a.Doc_no >= " & Val(txtinvoice.Text) & " and a.doc_no <= " & Val(txtInvoiceTo.Text) & " and a.doc_no=b.doc_no order by a.Doc_no"
            rsInvoice.GetResult(strsql)
            If rsInvoice.GetNoRows > 0 Then
                lvwInvoice.Items.Clear()
                rsInvoice.MoveFirst()
                While Not rsInvoice.EOFRecord
                    'UPGRADE_WARNING: Couldn't resolve default property of object rsInvoice.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    mlvwInvoice = Me.lvwInvoice.Items.Add(rsInvoice.GetValue("doc_no"))
                    'UPGRADE_WARNING: Lower bound of collection mlvwInvoice has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object rsInvoice.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If mlvwInvoice.SubItems.Count > 1 Then
                        mlvwInvoice.SubItems(1).Text = rsInvoice.GetValue("invoice_date")
                    Else
                        mlvwInvoice.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsInvoice.GetValue("invoice_date")))
                    End If
                    '----Changes made By Davinder on 05-09-2005 according to issue ID:-14999 to fill the third column of list view with Item code
                    'UPGRADE_WARNING: Lower bound of collection mlvwInvoice has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object rsInvoice.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If mlvwInvoice.SubItems.Count > 2 Then
                        mlvwInvoice.SubItems(2).Text = rsInvoice.GetValue("cust_item_code")
                    Else
                        mlvwInvoice.SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsInvoice.GetValue("cust_item_code")))
                    End If
                    rsInvoice.MoveNext()
                End While
            End If
            rsInvoice.ResultSetClose()
            'UPGRADE_NOTE: Object rsInvoice may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            rsInvoice = Nothing
        End If
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
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
        Dim Rs As New ADODB.Recordset
        Rs = New ADODB.Recordset
        Rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        Rs.Open(strField, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
        If Rs.RecordCount > 0 Then
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            If IsDbNull(Rs.Fields(0).Value) = False Then
                Find_Value = Rs.Fields(0).Value
            Else
                Find_Value = ""
            End If
        Else
            Find_Value = ""
        End If
        Rs.Close()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Function ValidateInvoice(ByRef pstrInvoiceNo As String, ByRef pstrItemCode As String) As Boolean
        '----------------------------------------------------
        ' Author             : Arshad Ali
        ' Create Date        : 25/05/2003
        ' Arguments          : None
        ' Return Value       : None
        ' Function           : To Validate invoice No
        '---------------------------------------------------------------------------------------------------------
        ' Revised By         : Davinder Singh
        ' Revision Date      : 05/09/2005
        ' Issue ID           : 17055
        ' Revision History   : function is mordified according to issue ID:-14999 by making it able to take second argument in the form of item code
        '-----------------------------------------------------------------------
        ' Revision By        : Davinder Singh
        ' Revision Date      : 01/05/2006
        ' Issue ID           : 17707
        ' Revision History   : To pick the Bin Qty. from Sales_dtl instead of Custitem_mst
        '-----------------------------------------------------------------------

        Dim strsql As String
        Dim rsInvoice As ClsResultSetDB
        Dim StrItemCode As String
        Dim boxQuantity As Double

        On Error GoTo Err_Handler

        ValidateInvoice = True

        If Len(pstrInvoiceNo) = 0 Then
            ValidateInvoice = False
            Exit Function
        End If

        strsql = "SELECT B.ITEM_CODE, B.SALES_QUANTITY, ISNULL(B.BINQUANTITY,0) BINQUANTITY FROM SALESCHALLAN_DTL A INNER JOIN SALES_DTL B "
        strsql = strsql & " ON A.LOCATION_CODE = B.LOCATION_CODE"
        strsql = strsql & " AND A.DOC_NO = B.DOC_NO"
        strsql = strsql & " AND A.UNIT_CODE = B.UNIT_CODE AND A.UNIT_CODE = '" & gstrUNITID & "'"
        strsql = strsql & " WHERE A.DOC_NO = '" & Trim(pstrInvoiceNo) & "'"
        strsql = strsql & " AND A.LOCATION_CODE ='" & Trim(txtUnitCode.Text) & " '"
        strsql = strsql & " AND B.CUST_ITEM_CODE = '" & pstrItemCode & "'"
        strsql = strsql & " AND BILL_FLAG = 1 AND CANCEL_FLAG =0 "

        rsInvoice = New ClsResultSetDB
        rsInvoice.GetResult(strsql)

        If rsInvoice.GetNoRows > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object rsInvoice.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            StrItemCode = rsInvoice.GetValue("Item_code")
            'UPGRADE_WARNING: Couldn't resolve default property of object rsInvoice.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mSalesQuantity = Val(rsInvoice.GetValue("Sales_Quantity"))
            'UPGRADE_WARNING: Couldn't resolve default property of object rsInvoice.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            boxQuantity = Val(rsInvoice.GetValue("BinQuantity"))
            If boxQuantity > 0 Then
                mBoxQuantity = boxQuantity
            Else
                MsgBox("No Bin Quantity is Defined for the item: " & StrItemCode & " of Invoice No: " & pstrInvoiceNo, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, ResolveResString(100))
                ValidateInvoice = False
            End If
        End If

        rsInvoice.ResultSetClose()
        'UPGRADE_NOTE: Object rsInvoice may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rsInvoice = Nothing

        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Function GetFileName() As String
        On Error GoTo Err_Handler
        Dim RSDB As ClsResultSetDB
        GetFileName = ""
        RSDB = New ClsResultSetDB
        RSDB.GetResult("Select isnull(BinCardFileName,'') BinCardFileName from Sales_Parameter WHERE UNIT_CODE = '" & gstrUNITID & "'")
        If RSDB.GetNoRows > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object RSDB.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Trim(RSDB.GetValue("BinCardFileName")) <> "" Then
                'UPGRADE_WARNING: Couldn't resolve default property of object RSDB.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                GetFileName = RSDB.GetValue("BinCardFileName")
            End If
        End If
        RSDB.ResultSetClose()
        RSDB = Nothing
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Sub ReadSalesParameter()
        On Error GoTo Err_Handler
        Dim RSDB As ClsResultSetDB
        RSDB = New ClsResultSetDB
        RSDB.GetResult("Select Isnull(DatePrintOnBinPrinting,0) as DatePrintOnBinPrinting, BINCARDWAITINGTIME from Sales_parameter WHERE UNIT_CODE = '" & gstrUNITID & "'")

        If RSDB.GetNoRows > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object RSDB.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mblnPrintDate = IIf(RSDB.GetValue("DatePrintOnBinPrinting") = "True", True, False)
            'UPGRADE_WARNING: Couldn't resolve default property of object RSDB.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mlngWaitingTime = RSDB.GetValue("BinCardWaitingtime")
        End If
        RSDB.ResultSetClose()
        'UPGRADE_NOTE: Object RSDB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        RSDB = Nothing
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
End Class