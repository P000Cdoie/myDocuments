Option Strict Off
Option Explicit On
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Friend Class frmMKTTRN0047
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------
	'Copyright(c)                               - MIND
	'Form Name (Physical Name)                  - FRMMKTTRN0047.frm
	'Created by                                  - Sourabh Khatri
	'Created Date                                - 24-02-2006
	'Form Description                            - Form will be used to upload information scanned by bar code device.
    '                                            - It will be used by SMIEL.
    '----------------------------------------------------
    'Modified by    :   Virendra Gupta
    'Modified ON    :   01/06/2011
    'Modified to support MultiUnit functionality
    '-----------------------------------------------------------------------
    '***********************************************************************************
    'Form Level Declarations
    Enum GridColumn
        check = 1
        FileLocation = 2
        status = 3
        inv_no = 4
    End Enum

    Dim mlngFormTag As Integer
    Dim mRsobject As New ClsResultSetDB
    Dim mDirFileUploading As Object 'Variable use to save location of file to be uploaded
    Dim mDirFileTransfer As Object 'Variable used to save location of file to be moved after uploading
    Dim FileObject As New Scripting.FileSystemObject
    Dim cmdObject As New ADODB.Command

    Private Sub cmdGrpMain_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.cmdGrpAuthorise.ButtonClickEventArgs) Handles cmdGrpMain.ButtonClick
        On Error GoTo Errorhandler
        Dim intCounter As Short
        Dim strInvNo As String
        Dim blnFlag As Boolean
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_AUTHORIZE
                Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                With Me.MainGrid
                    intCounter = 1
                    cmdObject.let_ActiveConnection(mP_Connection)
                    cmdObject.CommandTimeout = 0
                    cmdObject.CommandType = ADODB.CommandTypeEnum.adCmdText
                    blnFlag = False
                    While intCounter <= .MaxRows
                        .Row = intCounter
                        .Col = GridColumn.status
                        .TypeButtonText = ""
                        .Row = intCounter
                        .Col = GridColumn.check
                        If Val(.Text) = 1 Then
                            .Row = intCounter
                            .Col = GridColumn.FileLocation
                            TransferData(.Text, intCounter)
                            blnFlag = True
                        End If
                        intCounter = intCounter + 1
                    End While
                    If blnFlag = False Then
                        Call MsgBox("Please check at least one file for transfer", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Else
                        Call MsgBox("Transfer successfully", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    End If
                    Me.Cursor = System.Windows.Forms.Cursors.Default
                End With
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_REFRESH
                Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                SetGrid()
                UpdateGrid()
                Me.Cursor = System.Windows.Forms.Cursors.Default
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
                With Me.MainGrid
                    Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                    If .MaxRows > 0 Then
                        intCounter = 1
                        While intCounter <= .MaxRows
                            .Row = intCounter
                            .Col = GridColumn.check
                            If Val(.Text) = 1 Then
                                .Row = intCounter
                                .Col = GridColumn.inv_no
                                strInvNo = strInvNo & .Text & ","
                                intCounter = intCounter + 1
                            End If
                        End While
                        strInvNo = Mid(strInvNo, 1, Len(strInvNo) - 1)
                        Call ShowReport(strInvNo)
                    End If
                    Me.Cursor = System.Windows.Forms.Cursors.Default
                End With
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
        End Select
        Exit Sub
Errorhandler:
        Me.Cursor = System.Windows.Forms.Cursors.Default
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub frmMKTTRN0047_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
        'Call Initialize_controls
        gblnCancelUnload = False
        gblnFormAddEdit = False
        mlngFormTag = mdifrmMain.AddFormNameToWindowList(Me.ctlFormHeader1.HeaderString())
        'Fit in client window
        Call FitToClient(Me, frmMain, ctlFormHeader1, cmdGrpMain, 300)
        ' Show tool tips
        Call ShowToolTips()
        ' Set Grin Layout
        Call SetGrid()
        'Change caption of button
        Me.cmdGrpMain.Caption(0) = "Transfer"
        mRsobject.GetResult("Select isnull(FileUploadingLoc,'') FileUploadingLoc ,isnull(FileTransferLoc,'') FileTransferLoc from Sales_parameter where Unit_Code = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        mDirFileUploading = mRsobject.GetValue("FileUploadingLoc")
        mDirFileTransfer = mRsobject.GetValue("FileTransferLoc")
        Call UpdateGrid()
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub frmMKTTRN0047_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mlngFormTag
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub frmMKTTRN0047_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub

    Private Sub frmMKTTRN0047_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Err_Handler
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub frmMKTTRN0047_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                frmMKTTRN0047_FormClosed(Me, New System.Windows.Forms.FormClosedEventArgs(False))
        End Select
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub frmMKTTRN0047_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '-------------------------------------------------------------------------------------------
        ' Author        : Sourabh Khatri
        ' Arguments     : Cancel as Integer
        ' Return Value  : NIL
        ' Function      : RemoveFormNameFromWindowList
        ' Release form Object Memory from Database.
        ' Created On    :24 Feb 2006
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        mdifrmMain.RemoveFormNameFromWindowList = mlngFormTag
        Me.Dispose()
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub ShowToolTips()
        '-------------------------------------------------------------------------------------------
        ' Author        : Sourabh Khatri
        ' Arguments     : Cancel as Integer
        ' Return Value  : NIL
        ' Function      : RemoveFormNameFromWindowList
        ' Release form Object Memory from Database.
        ' Created On    :24 Feb 2006
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub SetGrid()
        '-------------------------------------------------------------------------------------------
        ' Author        : Sourabh Khatri
        ' Return Value  : NIL
        ' Function      : Set grid property
        ' Created On    : 24 Feb 2006
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        With Me.MainGrid
            .maxRows = 0
            .Row = .maxRows
            .MaxCols = GridColumn.inv_no
            .Col = GridColumn.check : .Text = "Check"
            .TypeTextWordWrap = True
            .Col = GridColumn.FileLocation : .Text = "File Name"
            .set_ColWidth(GridColumn.FileLocation, 27)
            .Col = GridColumn.status : .Text = "Status"
            .set_ColWidth(GridColumn.status, 16)
            .Col = GridColumn.inv_no
            .ColHidden = True
            .set_RowHeight(0, 18)
        End With
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub AddRowInGrid(ByVal strFileName As String)
        '-------------------------------------------------------------------------------------------
        ' Author        : Sourabh Khatri
        ' Return Value  : NIL
        ' Function      : Set grid property
        ' Created On    : 24 Feb 2006
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        With Me.MainGrid
            .maxRows = .maxRows + 1
            .Row = .maxRows
            .Col = GridColumn.check
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
            .value = 1
            .TypeCheckCenter = True
            .Col = GridColumn.FileLocation
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Text = strFileName
            .Col = GridColumn.status
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
            .Col = GridColumn.inv_no
            .ColHidden = True
            Me.cmdGrpMain.Enabled(2) = True
            .set_RowHeight(.MaxRows, 15)
        End With
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Public Sub UpdateGrid()
        '-------------------------------------------------------------------------------------------
        ' Author        : Sourabh Khatri
        ' Return Value  : NIL
        ' Function      : Set Grid property
        ' Created On    : 24 Feb 2006
        '--------------------------------------------------------------------------------------------------
        Dim sFile As String
        On Error GoTo ErrHandler
        Dim Flag As Boolean
        Flag = False
        If Len(mDirFileUploading) > 0 Then
            If FileObject.FolderExists(mDirFileUploading) Then
                sFile = Dir(mDirFileUploading & "\*.*")
                While sFile <> ""
                    If sFile <> "." And sFile <> ".." Then
                        AddRowInGrid((sFile))
                        Flag = True
                    End If
                    sFile = Dir()
                End While
            End If
            If Flag Then
                With Me.cmdGrpMain
                    .Enabled(0) = True
                    .Enabled(1) = True
                End With
            End If
        Else
            Call MsgBox("Folder location specified in sales parameter does not exist physicaly.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            Exit Sub
        End If
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub optAll_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAll.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo Errorhandler
            Dim intCounter As Short
            With Me.MainGrid
                If .maxRows > 0 Then
                    intCounter = 1
                    While intCounter <= .maxRows
                        .Row = intCounter
                        .Col = GridColumn.check
                        .value = 1
                        intCounter = intCounter + 1
                    End While
                End If
            End With
            Exit Sub 'To prevent the execution of errhandler
Errorhandler:  'The Error Handling Code Starts here
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        End If
    End Sub

    Private Sub optUncheckAll_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optUncheckAll.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo Errorhandler
            Dim intCounter As Short
            With Me.MainGrid
                If .maxRows > 0 Then
                    intCounter = 1
                    While intCounter <= .maxRows
                        .Row = intCounter
                        .Col = GridColumn.check
                        .value = 0
                        intCounter = intCounter + 1
                    End While
                End If
            End With
            Exit Sub 'To prevent the execution of errhandler
Errorhandler:  'The Error Handling Code Starts here
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        End If
    End Sub

    Public Function TransferData(ByVal strFileName As String, ByVal intCounter As Short) As Object
        On Error GoTo Errorhandler
        Dim arr() As String
        Dim strm As Scripting.TextStream
        Dim strsql As String
        Dim finDate As String
        Dim strInvNo As String
        If FileObject.FileExists(mDirFileUploading & "\" & strFileName) Then
            strm = FileObject.OpenTextFile(mDirFileUploading & "\" & strFileName, Scripting.IOMode.ForReading, False)
            mP_Connection.Execute("Delete BarCodeDispatchData_tmp where Unit_Code = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            mRsobject.GetResult("Select Fin_Year_Notation  From Financial_Year_tb Where getdate() >=  Fin_Start_date And getdate()  <= Fin_End_date and Unit_Code = '" & gstrUNITID & "'")
            finDate = mRsobject.GetValue("Fin_Year_Notation")
            While Not strm.AtEndOfStream
                strsql = strm.ReadLine
                If Len(Trim(strsql)) > 0 Then
                    arr = Split(strsql, "*")
                    strsql = "Insert into BarCodeDispatchData_tmp(InvoiceNo,PartCode,Qty,LotBagNo,Unit_Code)"
                    strsql = strsql & "values('" & finDate & New String("0", 6 - Len(arr(0))) & arr(0) & "','" & arr(1) & "'," & arr(2) & ",'" & arr(4) & arr(3) & "', '" & gstrUNITID & "')"
                    cmdObject.CommandText = strsql
                    cmdObject.Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
            End While
            strm.Close()
            'To insert exception list into exception table
            cmdObject.CommandText = "Insert into BarCodeDispatchData_Exception Select a.* from BarCodeDispatchData_tmp a inner join sales_dtl b on a.invoiceno = b.doc_no and a.Unit_Code = b.Unit_Code where invoiceno + partcode <> doc_no + cust_Item_Code and a.unit_Code = '" & gstrUNITID & "'"
            cmdObject.Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            'To delete exception record from temp table
            cmdObject.CommandText = "delete from BarCodeDispatchData_tmp where not exists(Select * from sales_dtl A where a.doc_no = BarCodeDispatchData_tmp.invoiceno and a.cust_Item_Code = BarCodeDispatchData_tmp.partcode and a.Unit_Code = '" & gstrUNITID & "') and Unit_Code =  '" & gstrUNITID & "'"
            cmdObject.Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            'To check duplicate data
            mRsobject.GetResult("Select a.* from BarCodeDispatchData a inner join BarCodeDispatchData_tmp b on a.invoiceno = b.invoiceno and a.partcode = b.partcode and a.lotbagno = b.lotbagno and a.Unit_Code = b.Unit_Code and a.Unit_Code = '" & gstrUNITID & "'")
            If mRsobject.RowCount > 0 Then
                If MsgBox("Scanned data available in file " & strFileName & " already exist in Record.Do you want to save again ?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                    'Insert record into main table
                    cmdObject.CommandText = "Insert into BarCodeDispatchData Select a.PartCode,a.Qty,b.Item_Code,a.LotBagNo,a.InvoiceNo,Unit_Code from BarCodeDispatchData_tmp a inner join Sales_dtl b on b.cust_item_Code = a.PartCode and b.doc_no = a.invoiceno and a.Unit_Code = b.Unit_Code and a.Unit_Code = '" & gstrUNITID & "'"
                Else
                    cmdObject.CommandText = "Insert into BarCodeDispatchData Select a.PartCode,a.Qty,b.Item_Code,a.LotBagNo,a.InvoiceNo,Unit_Code from BarCodeDispatchData_tmp a inner join Sales_dtl b on b.cust_item_Code = a.PartCode and b.doc_no = a.invoiceno and a.Unit_Code = b.Unit_Code and a.Unit_Code = '" & gstrUNITID & "' and not exists ( Select * from BarCodeDispatchData where a.invoiceno = BarCodeDispatchData.invoiceno and a.partcode = BarCodeDispatchData.partcode and a.lotbagno = BarCodeDispatchData.lotbagno and  a.Unit_Code = '" & gstrUNITID & "')"
                End If
            Else
                cmdObject.CommandText = "Insert into BarCodeDispatchData Select a.PartCode,a.Qty,b.Item_Code,a.LotBagNo,a.InvoiceNo,Unit_Code from BarCodeDispatchData_tmp a inner join Sales_dtl b on b.cust_item_Code = a.PartCode and b.doc_no = a.invoiceno and a.Unit_Code = b.Unit_Code and a.Unit_Code = '" & gstrUNITID & "'"
            End If
            cmdObject.Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            mRsobject.GetResult("Select distinct invoiceno from BarCodeDispatchData_tmp where Unit_Code = '" & gstrUNITID & "'")
            While Not mRsobject.EOFRecord
                strInvNo = strInvNo & mRsobject.GetValue("invoiceno") & ","
                mRsobject.MoveNext()
            End While
            strInvNo = Mid(strInvNo, 1, Len(strInvNo) - 1)
            With Me.MainGrid
                .Row = intCounter
                .Col = GridColumn.inv_no
                .Text = strInvNo
            End With
            strsql = "Select invoiceno,partcode,Sales_quantity ,sum(Qty) as LotQty from Sales_dtl a inner join BarCodeDispatchData b on a.cust_item_Code = b.PartCode and a.doc_no = b.invoiceno and a.Unit_Code = b.Unit_Code  where  a.Unit_Code = '" & gstrUNITID & "' and b.invoiceno in (Select distinct invoiceno from BarCodeDispatchData_tmp where Unit_Code = '" & gstrUNITID & "') group by invoiceno,partcode,Sales_quantity having sales_quantity <> sum(Qty) "
            Call mRsobject.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If Not mRsobject.EOFRecord Then
                With Me.MainGrid
                    .Row = intCounter
                    .Col = GridColumn.status
                    .TypeButtonText = "Incomplete File"
                End With
            Else
                With Me.MainGrid
                    .Row = intCounter
                    .Col = GridColumn.status
                    .TypeButtonText = "Complete File"
                End With
            End If
            Call FileObject.MoveFile(mDirFileUploading & "\" & strFileName, mDirFileTransfer & "\" & strFileName)
        Else
            Call MsgBox("File " & strFileName & " doesn't exist at location " & mDirFileUploading, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
        End If
        Exit Function 'To prevent the execution of errhandler
Errorhandler:  'The Error Handling Code Starts here
        If Err.Number = 58 Then Resume Next
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function

    Private Sub ShowReport(ByRef strInvNo As String)
        On Error GoTo Errorhandler
        Dim strSelectionFormula As Object
        '<<<<CR11 Code Starts>>>>
        Dim objRpt As ReportDocument
        Dim frmReportViewer As New eMProCrystalReportViewer
        objRpt = frmReportViewer.GetReportDocument()
        frmReportViewer.ShowPrintButton = True
        frmReportViewer.ShowTextSearchButton = True
        frmReportViewer.ShowZoomButton = True
        frmReportViewer.ReportHeader = Me.ctlFormHeader1.HeaderString()
        '<<<<CR11 Code Ends>>>>
        With objRpt
            'load the report
            .Load(My.Application.Info.DirectoryPath & "\Reports\rptBarCode-SalesWise.rpt")
            .DataDefinition.FormulaFields("CompanyName").Text = "'" & gstrCOMPANY & "'"
            .DataDefinition.FormulaFields("Address").Text = "'" & gstr_WRK_ADDRESS1 & gstr_WRK_ADDRESS2 & "'"
            .DataDefinition.FormulaFields("fromdate").Text = "'01/JAN/2005'"
            .DataDefinition.FormulaFields("InvoiceNo").Text = "[" & strInvNo & "]"
            .DataDefinition.FormulaFields("LotNo like").Text = " '%'"
            .DataDefinition.FormulaFields("Todate").Text = "'" & GetServerDate() & "'"
        End With
        strSelectionFormula = "{BarCodeDispatchData.InvoiceNo} IN [" & strInvNo & "]"
        objRpt.RecordSelectionFormula = strSelectionFormula & " and {BarCodeDispatchData.UNIT_CODE} = '" & gstrUNITID & "'"
        frmReportViewer.Zoom = 120
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
Errorhandler:
        If Err.Number = 20545 Then
            Resume Next
        Else
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub

    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        Exit Sub    'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub MainGrid_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles MainGrid.ButtonClicked
        On Error GoTo Errorhandler
        With Me.MainGrid
            If e.col = GridColumn.status Then
                .Row = e.row : .Col = GridColumn.inv_no
                If Trim(.Text) <> "" Then
                    Call ShowReport(.Text)
                End If
            End If
        End With
        Exit Sub
Errorhandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
End Class