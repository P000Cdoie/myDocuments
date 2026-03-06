Option Strict Off
Option Explicit On
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO
Friend Class frmMKTTRN0046
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------
	'Copyright(c)                               - MIND
	'Form Name (Physical Name)     - FRMADMMST0002.frm
	'Created by                                  - Sourabh Khatri
	'Created Date                              - 18-01-2006
	'Form Description                        - 57F4 Challan
    '----------------------------------------------------
    'Modified by    :   Virendra Gupta
    'Modified ON    :   20/05/2011
    'Modified to support MultiUnit functionality
    '-----------------------------------------------------------------------
	'Form Level Declarations
    Dim mlngFormTag As Integer
    Private Enum EnumGrid
        CheckBox = 1
        ItemCode = 2
        ItemDescription = 3
        CUSTDRGNO = 4
        ReceiveQty = 5
        DispatchQty = 6
        StockQty = 7
        BalanceQty = 8
        InvoiceQty = 9
        KanbanNo = 10
    End Enum
    Private Enum KANBANGrid1
        KanbanNo = 1
        SchDate = 2
        SChTime = 3
        UNLoc = 4
        USLoc = 5
        BalanceQty = 6
        SelectedQty = 7
    End Enum
    Dim mblnNagare As Boolean
    Dim rsKanBan As ADODB.Recordset
    Dim objInvoicePrint As New prj_InvoicePrinting.clsInvoicePrinting(gstrDateFormat)
    Private Sub CmdButtons_57F4_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles CmdButtons_57F4.ButtonClick
        Dim strMasterString As String
        Dim strDetailString As String
        Dim strDocNo As String
        Dim cmdObject As New ADODB.Command
        Dim Intcounter As Short
        Dim BlnTrans As Boolean
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                Call Initialize_controls()
                Me.txtGrinNo.Focus()
                GenerateDisRecordSet()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                If Val(SelectDataFromTable("Invoice_Lock", "mkt_57F4Challan_Hdr", " doc_no = '" & Me.txtDocumentNo.Text & "' and Unit_Code = '" & gstrUNITID & "'")) = 1 Then
                    MsgBox(" 57F4 Challan is locked.You can not edit it", MsgBoxStyle.Information, ResolveResString(100))
                    Me.CmdButtons_57F4.Revert()
                    Exit Sub
                End If
                Call EnableControls(True, Me)
                Me.txtGrinNo.Enabled = False : Me.cmdGrinHelp.Enabled = False : Me.dt57F4Date.Enabled = False
                Me.dtDocDate.Enabled = False : Me.txtDocumentNo.Enabled = False : Me.cmdDOcHelp.Enabled = False
                Me.frmOption.Enabled = False
                ShowItemsInGrid((Me.txtGrinNo.Text))
                Me.MainGrid.Row = 1
                Me.MainGrid.Row2 = Me.MainGrid.MaxRows
                Me.MainGrid.Col = EnumGrid.CheckBox
                Me.MainGrid.Col2 = EnumGrid.CheckBox
                Me.MainGrid.BlockMode = True
                Me.MainGrid.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                Me.MainGrid.ColHidden = False
                Me.MainGrid.BlockMode = False
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                'Set false into transaction boolean variable
                BlnTrans = False
                If Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                    If ValidDataBeforeSave() Then
                        'Genarete Document number and open transaction
                        mP_Connection.BeginTrans()
                        BlnTrans = True
                        If Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                            strDocNo = GenerateDocNO(109, Me.dtDocDate.Text, eMPowerFunctions.DocTypeEnum.Doc_ECN, False, True)
                            If strDocNo = "" Then
                                MsgBox("Please define document series in document master", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                                mP_Connection.RollbackTrans()
                                Exit Sub
                            End If
                            Me.txtDocumentNo.Text = strDocNo
                        Else
                            strDocNo = Me.txtDocumentNo.Text
                        End If
                        If Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                            strMasterString = " Insert into Mkt_57F4Challan_Hdr (Doc_Type,Doc_No,Doc_Date,Grin_No," & " Location_Code,Against_Nagare,Cancel_Flag,NatureOfProc,Trans_Name,Truck_No,Ent_dt,Ent_UserId,Upd_dt, Upd_UserId, Unit_Code)  Values (109,'" & strDocNo & "','" & getDateForDB(Me.dtDocDate.Value) & "'," & " '" & Me.txtGrinNo.Text & "','" & Me.txtLocationCode.Text & "'," & IIf(Me.optNagare.Checked, 1, 0) & ",0,'" & Me.txtNatureOfProc.Text & "','" & Me.txtTransporter.Text & "','" & Me.txtTruckNo.Text & "',GETDATE(),'" & mP_User & "',GETDATE(),'" & mP_User & "','" & gstrUNITID & "')"
                        Else
                            strMasterString = " Update Mkt_57F4Challan_Hdr Set Location_Code = '" & Me.txtLocationCode.Text & "',NatureOfProc = '" & Me.txtNatureOfProc.Text & "',Trans_Name = '" & Me.txtTransporter.Text & "',Truck_No = '" & Me.txtTruckNo.Text & "',Upd_dt = GETDATE()," & " Upd_UserId = '" & mP_User & "' where doc_no = '" & Me.txtDocumentNo.Text & "' and doc_type = 109 and Unit_Code = '" & gstrUNITID & "'"
                        End If
                        With cmdObject
                            .let_ActiveConnection(mP_Connection)
                            .CommandType = ADODB.CommandTypeEnum.adCmdText
                            .CommandText = strMasterString
                            .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End With
                        With Me.MainGrid
                            For Intcounter = 1 To .MaxRows
                                .Row = Intcounter
                                .Col = EnumGrid.CheckBox
                                If System.Math.Abs(Val(.Value)) = 1 Then
                                    strDetailString = ""
                                    If Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                                        .Col = EnumGrid.ItemCode
                                        If SelectDataFromTable("item_Code", "mkt_57f4challan_dtl", " item_Code = '" & Trim(.Text) & "' and doc_No = '" & Me.txtDocumentNo.Text & "' and Unit_code = '" & gstrUNITID & "'") <> "" Then
                                            .Col = EnumGrid.InvoiceQty
                                            strDetailString = "Update Mkt_57F4Challan_dtl set Invoice_Qty = " & Val(Trim(.Text)) & ",Upd_dt = GETDATE(),Upd_UserId = '" & mP_User & "'"
                                            .Col = EnumGrid.ItemCode
                                            strDetailString = strDetailString & " where item_Code = '" & Trim(.Text) & "' and doc_no = '" & Me.txtDocumentNo.Text & "' and doc_type = 109 and Unit_Code = '" & gstrUNITID & "'"
                                        End If
                                    End If
                                    If Trim(strDetailString) = "" Then
                                        strDetailString = "Insert into Mkt_57F4Challan_dtl (Doc_Type,Doc_No,Item_Code," & " Cust_Drgno,Balance_Qty,Invoice_Qty,Ent_dt,Ent_UserId," & " Upd_dt,Upd_UserId, Unit_Code) Values (109,'" & strDocNo & "',"
                                        'Item Code
                                        .Col = EnumGrid.ItemCode
                                        strDetailString = strDetailString & "'" & Trim(.Text) & "',"
                                        'Customer drawing No
                                        .Col = EnumGrid.CUSTDRGNO
                                        strDetailString = strDetailString & "'" & Trim(.Text) & "',"
                                        'Balance Quantity
                                        .Col = EnumGrid.BalanceQty
                                        strDetailString = strDetailString & "'" & Trim(.Text) & "',"
                                        'Invoice Quantity
                                        .Col = EnumGrid.InvoiceQty
                                        strDetailString = strDetailString & "'" & Trim(.Text) & "',"
                                        'User id and date time
                                        strDetailString = strDetailString & "GETDATE(),'" & mP_User & "',GETDATE(),'" & mP_User & "','" & gstrUNITID & "')"
                                    End If
                                    With cmdObject
                                        .CommandText = strDetailString
                                        .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    End With
                                    If Me.optNagare.Checked = True Then
                                        .Col = EnumGrid.ItemCode
                                        rsKanBan.Filter = "ItemCode= '" & Trim(.Text) & "'"
                                        If Not rsKanBan.EOF Then
                                            mP_Connection.Execute("Delete from Mkt_57F4ChallanKanBan_Dtl where doc_no = '" & Me.txtDocumentNo.Text & "' and item_Code = '" & Trim(.Text) & "' and doc_type = 109 and Unit_Code = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                            While Not rsKanBan.EOF
                                                If rsKanBan.Fields("kanbanQty").Value > 0 Then
                                                    strDetailString = "insert into Mkt_57F4ChallanKanBan_Dtl(Doc_Type,Doc_No,Item_Code,KanBan_No,Quantity,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,Unit_Code) " & " values (109,'" & Me.txtDocumentNo.Text & "','" & Trim(.Text) & "','" & rsKanBan.Fields("kanbanno").Value & "','" & rsKanBan.Fields("kanbanqty").Value & "',GETDATE(),'" & mP_User & "',GETDATE(),'" & mP_User & "','" & gstrUNITID & "')"
                                                    With cmdObject
                                                        .CommandText = strDetailString
                                                        .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                    End With
                                                End If
                                                rsKanBan.MoveNext()
                                            End While
                                        End If
                                        rsKanBan.Filter = ADODB.FilterGroupEnum.adFilterNone
                                    End If
                                End If
                            Next
                            If Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                                MsgBox("Transaction completed successfully with document No " & Me.txtDocumentNo.Text, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                            Else
                                MsgBox("Transaction completed successfully ", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                            End If
                            mP_Connection.CommitTrans()
                            BlnTrans = False
                            rsKanBan = Nothing
                        End With
                        strDocNo = Me.txtDocumentNo.Text
                        Me.CmdButtons_57F4.Revert()
                        Call EnableControls(False, Me)
                        Initialize_controls()
                        Me.txtDocumentNo.Text = strDocNo
                        ShowDataInViewMode()
                    End If
                Else
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                Call frmMKTTRN0046_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
                If Val(SelectDataFromTable("Invoice_Lock", "mkt_57F4Challan_Hdr", " doc_no = '" & Me.txtDocumentNo.Text & "' and Unit_Code = '" & gstrUNITID & "'")) = 1 Then
                    MsgBox(" 57F4 Challan is locked.You can not delete it", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If
                If ConfirmWindow(10054, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    mP_Connection.Execute("Delete from mkt_57f4challan_hdr where doc_no = '" & Me.txtDocumentNo.Text & "' and doc_type = 109 and Unit_Code = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    mP_Connection.Execute("Delete from mkt_57f4challan_dtl where doc_no = '" & Me.txtDocumentNo.Text & "' and doc_type = 109 and Unit_Code = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    mP_Connection.Execute("Delete from Mkt_57F4ChallanKanBan_Dtl where doc_no = '" & Me.txtDocumentNo.Text & "' and doc_type = 109 and Unit_Code = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    MsgBox("Challan has been deleted successfully", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                    Initialize_controls()
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
                If Trim(Me.txtDocumentNo.Text) <> "" Then
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                    If Val(SelectDataFromTable("Invoice_Lock", "mkt_57F4Challan_Hdr", " doc_no = '" & Me.txtDocumentNo.Text & "' and Unit_Code = '" & gstrUNITID & "'")) = 0 Then
                        If MsgBox("Do you want to lock challan", MsgBoxStyle.YesNo + MsgBoxStyle.Information, ResolveResString(100)) = MsgBoxResult.Yes Then
                            If Not LockInvoice() Then
                                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                                Exit Sub
                            Else
                                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                                MsgBox("Challan has been locked successfully", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                                Me.lblYesNo.Text = "Yes"
                            End If
                        End If
                    End If
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                    If UCase(SelectDataFromTable("Unt_Codeid", "gen_unitmaster", " Unt_Codeid like '%' ")) = "SUN" Then
                        Me.SSTab1.SelectedIndex = 1
                        Call PrintingInvoice()
                    End If
                    '<<<<CR11 Code Starts>>>>
                    Dim objRpt As ReportDocument
                    Dim frmReportViewer As New eMProCrystalReportViewer
                    objRpt = frmReportViewer.GetReportDocument()
                    frmReportViewer.ShowPrintButton = True
                    frmReportViewer.ShowTextSearchButton = True
                    frmReportViewer.ShowZoomButton = True
                    frmReportViewer.ReportHeader = Me.ctlFormHeader_57F4.HeaderString()
                    '<<<<CR11 Code Ends>>>>
                    If Val(CStr(Me.chkPieceMeal.CheckState)) = 1 Then
                        With objRpt
                            .Load(My.Application.Info.DirectoryPath & "\reports\57F4CustRegister.rpt")
                            .DataDefinition.FormulaFields("compname").Text = "'" & gstrCOMPANY & "'"
                            .DataDefinition.FormulaFields("CompAdd1").Text = "'" & gstr_WRK_ADDRESS1 & "'"
                            .DataDefinition.FormulaFields("Address").Text = "'" & gstr_WRK_ADDRESS2 & "'"
                            .RecordSelectionFormula = " {grn_hdr.doc_no} = " & Me.txtGrinNo.Text & " and {grn_hdr.UNIT_CODE} = '" & gstrUNITID & "'"
                            frmReportViewer.Zoom = 150
                            frmReportViewer.Show()
                        End With
                    End If
                    If Val(CStr(Me.chkAnnexure.CheckState)) = 1 Then
                        With objRpt
                            frmReportViewer.ReportHeader = "ANNEXURE VI"
                            .Load(My.Application.Info.DirectoryPath & "\reports\57F4CustAnnexure.rpt")
                            .DataDefinition.FormulaFields("Cname").Text = "'" & gstr_WRK_ADDRESS1 & "'"
                            .DataDefinition.FormulaFields("Caddress1").Text = "'" & gstr_WRK_ADDRESS1 & "'"
                            .DataDefinition.FormulaFields("Caddress2").Text = "'" & gstr_WRK_ADDRESS2 & "'"
                            .DataDefinition.FormulaFields("RemoveDateTime").Text = "'" & dtpRemoval.Text & "  " & Me.dtpRemovalTime.Text & "'"
                            .SetParameterValue(0, Val(Me.txtDocumentNo.Text))
                            .RecordSelectionFormula = " {vw_57F4Annexure.doc_no} = " & Me.txtDocumentNo.Text & " and {vw_57F4Annexure.UNIT_CODE} = '" & gstrUNITID & "'"
                            frmReportViewer.Zoom = 150
                            frmReportViewer.Show()
                        End With
                    End If
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                End If
        End Select
        Exit Sub
Errorhandler:
        If BlnTrans = True Then
            mP_Connection.RollbackTrans()
        End If
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        Me.rtbInvoicePreview.Text = ""
    End Sub
    Private Sub cmdDOcHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDOcHelp.Click
        On Error GoTo errohandler
        Dim strDocNo As String
        strDocNo = ShowList(1, (Me.txtDocumentNo.MaxLength), "", "doc_No", "Doc_Date", "Mkt_57F4Challan_hdr", "")
        If strDocNo = "-1" Then
            Call MsgBox("No document no. exists", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        Else
            Me.txtDocumentNo.Text = strDocNo
            TxtDocumentNo_Validating(txtDocumentNo, New System.ComponentModel.CancelEventArgs((False)))
        End If
        Exit Sub
errohandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdGRINHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGrinHelp.Click
        On Error GoTo errohandler
        Dim strGRINNo As String
        strGRINNo = ShowList(1, (Me.txtGrinNo.MaxLength), "", "doc_No", "Vendor_Name", "Mkt_57F4GrinNo('" & gstrUNITID & "')", "")
        If strGRINNo = "-1" Then
            Call MsgBox("Grin No does not exist", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        Else
            Me.txtGrinNo.Text = strGRINNo
            txtGRINNo_Validating(txtGrinNo, New System.ComponentModel.CancelEventArgs((False)))
        End If
        Exit Sub
errohandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdKANBANCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdKANBANCancel.Click
        On Error GoTo ErrHandler
        Me.frmKANBAN.Visible = False
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdKANBANOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdKANBANOK.Click
        On Error GoTo ErrHandler
        Dim Intcounter As Short
        If Me.CmdButtons_57F4.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            'Validation for 0 value
            If Val(Me.lblKANBANTotal.Text) <= 0 Then
                Call MsgBox("Kanban Quantity can not be zero", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, ResolveResString(100))
                Me.KANBANGrid.Row = 1
                Me.KANBANGrid.Col = KANBANGrid1.SelectedQty
                Me.KANBANGrid.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                Me.KANBANGrid.Focus()
                Exit Sub
            End If
            With Me.KANBANGrid
                rsKanBan.Filter = "itemCode = '" & Me.lblItem.Text & "'"
                If Not rsKanBan.EOF Then
                    For Intcounter = 1 To .MaxRows
                        rsKanBan.MoveFirst()
                        While Not rsKanBan.EOF
                            .Row = Intcounter
                            .Col = KANBANGrid1.KanbanNo
                            If StrComp(rsKanBan.Fields("kanbanno").Value, Trim(.Text), CompareMethod.Text) = 0 Then
                                .Col = KANBANGrid1.SelectedQty
                                rsKanBan.Fields("kanbanqty").Value = Val(.Text)
                                rsKanBan.Update()
                            End If
                            rsKanBan.MoveNext()
                        End While
                    Next
                    rsKanBan.Filter = ADODB.FilterGroupEnum.adFilterNone
                Else
                    rsKanBan.Filter = ADODB.FilterGroupEnum.adFilterNone
                    For Intcounter = 1 To .MaxRows
                        rsKanBan.AddNew()
                        .Row = Intcounter
                        rsKanBan.Fields("ItemCode").Value = Me.lblItem.Text
                        .Col = KANBANGrid1.KanbanNo
                        rsKanBan.Fields("kanbanno").Value = .Text
                        .Col = KANBANGrid1.SchDate
                        rsKanBan.Fields("schDate").Value = getDateForDB(.Text)
                        .Col = KANBANGrid1.SChTime
                        rsKanBan.Fields("SchTime").Value = IIf(.Text = "", "00:00", .Text)
                        .Col = KANBANGrid1.UNLoc
                        rsKanBan.Fields("unLoc").Value = .Text
                        .Col = KANBANGrid1.USLoc
                        rsKanBan.Fields("usLoc").Value = .Text
                        .Col = KANBANGrid1.BalanceQty
                        rsKanBan.Fields("BalanceQty").Value = Val(.Text)
                        .Col = KANBANGrid1.SelectedQty
                        rsKanBan.Fields("KanbanQty").Value = Val(.Text)
                        rsKanBan.Update()
                    Next
                End If
            End With
            With Me.MainGrid
                If .ActiveRow <> .MaxRows Then
                    .Row = .ActiveRow + 1
                    .Col = EnumGrid.InvoiceQty
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                Else
                    Me.CmdButtons_57F4.Focus()
                End If
            End With
        End If
        Me.frmKANBAN.Visible = False
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdlocationhelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdLocationHelp.Click
        On Error GoTo ErrHandler
        Dim strLoc As String
        strLoc = ShowList(1, (Me.txtLocationCode.MaxLength), "", "Location_Code", "Description", "Location_mst", " and Loc_Type in ('M','P','S')")
        If strLoc = "-1" Then
            MsgBox("Location does not exist", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        Else
            Me.txtLocationCode.Text = strLoc
        End If
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrint.Click
        On Error GoTo ErrHandler
        Dim intCount As Short
        Dim varTemp As Object
        Dim strfilename As String
        Dim intNoCopies As Integer
        Dim dblWaitingTime As Double
        'Kill App.Path & "\TypeToPrn.bat"
        Kill(gstrLocalCDrive & "EmproInv\TypeToPrn.bat")
        If Len(objInvoicePrint.FileName) > 0 Then
            strfilename = objInvoicePrint.FileName
        End If
        If intNoCopies = 0 Then intNoCopies = 1
        dblWaitingTime = Val(SelectDataFromTable("waitingTime", "sales_parameter", " waitingTime is not null and Unit_code = '" & gstrUNITID & "'"))
        If dblWaitingTime = 0 Then
            dblWaitingTime = 5000
        End If
TypeFileNotFoundCreateRetry:
        For intCount = 1 To intNoCopies
            varTemp = Shell("cmd.exe /c " & gstrLocalCDrive & "EmproInv\TypeToPrn.bat " & strfilename, AppWinStyle.Hide)
            Sleep(dblWaitingTime)
            Call printBarCode(objInvoicePrint.BCFileName)
            Sleep(dblWaitingTime)
            varTemp = Shell("cmd.exe /c " & gstrLocalCDrive & "EmproInv\TypeToPrn.bat " & gstrLocalCDrive & "EmproInv\PageFeed.txt", AppWinStyle.Hide)
        Next
        Exit Sub
ErrHandler:
        If Err.Number = 53 Then
            'Open App.Path & "\" & "TypeToPrn.bat" For Append As #1
            FileOpen(1, gstrLocalCDrive & "EmproInv\TypeToPrn.bat", OpenMode.Append)
            PrintLine(1, "Type %1> prn") '& Printer.Port
            FileClose(1)
            GoTo TypeFileNotFoundCreateRetry
        End If
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0046_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mlngFormTag
        Me.SSTab1.SelectedIndex = 0
        If Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            Me.txtDocumentNo.Focus()
        Else
            Me.txtGrinNo.Focus()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0046_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0046_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Err_Handler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            'Call ctlFormHeader1_Click
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0046_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                'If user press the ESC Key ,the Form will be in View Mode
                If Me.CmdButtons_57F4.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Me.CmdButtons_57F4.Revert()
                        Call EnableControls(False, Me)
                        Initialize_controls()
                    End If
                End If
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
    Private Sub frmMKTTRN0046_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
        Call Initialize_controls()
        gblnCancelUnload = False
        gblnFormAddEdit = False
        mlngFormTag = mdifrmMain.AddFormNameToWindowList(Me.ctlFormHeader_57F4.HeaderString())
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        'Fit in client window
        Call FitToClient(Me, frmMain, ctlFormHeader_57F4, CmdButtons_57F4, 300)
        ' Show tool tips
        Call ShowToolTips()
        ' Set Grin Layout
        Me.optNagare.Checked = True
        mblnNagare = True
        'Call SetGrid
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        If Directory.Exists(gstrLocalCDrive + "EmproInv") = False Then
            Directory.CreateDirectory(gstrLocalCDrive + "EmproInv")
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub Initialize_controls()
        '-------------------------------------------------------------------------------------------
        ' Author        : Sourabh Khatri
        ' Arguments     : Cancel as Integer
        ' Return Value  : NIL
        ' Function      : RemoveFormNameFromWindowList
        ' Release form Object Memory from Database.
        ' Created On    :18 jan 2006
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Me.cmdDOcHelp.Image = My.Resources.ico111.ToBitmap
        Me.cmdGrinHelp.Image = My.Resources.ico111.ToBitmap
        Me.cmdLocationHelp.Image = My.Resources.ico111.ToBitmap
        Call EnableControls(False, Me)
        Me.frmKANBAN.Visible = False
        Me.dt57F4Date.Format = DateTimePickerFormat.Custom
        Me.dt57F4Date.CustomFormat = gstrDateFormat
        Me.dtDocDate.Format = DateTimePickerFormat.Custom
        Me.dtDocDate.CustomFormat = gstrDateFormat
        If Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            Me.txtGrinNo.Enabled = True : Me.txtLocationCode.Enabled = True : Me.dtDocDate.Enabled = True : Me.MainGrid.Enabled = True
            Me.cmdGrinHelp.Enabled = True : Me.cmdLocationHelp.Enabled = True : Me.frmOption.Enabled = True
            Me.Lbl57F4No.Text = "" : Me.lblCustName.Text = "" : Me.dt57F4Date.Value = GetServerDate()
            Me.txtGrinNo.Text = "" : Me.txtDocumentNo.Text = "" : Me.txtDocumentNo.Enabled = False
            Me.txtDocumentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : Me.txtLocationCode.Text = "" : Me.lblLocDesc.Text = ""
            Me.txtGrinNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : Me.txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            Me.optNagare.Enabled = True : Me.optwithOutNagare.Enabled = True : Me.dtDocDate.Value = GetServerDate()
            Me.lblLocked.Visible = False : Me.lblYesNo.Visible = False : Me.lblYesNo.Text = ""
            Me.txtNatureOfProc.Text = "" : Me.txtTransporter.Text = "" : Me.txtTruckNo.Text = ""
            Me.txtNatureOfProc.Enabled = True : Me.txtTransporter.Enabled = True : Me.txtTruckNo.Enabled = True
            Me.txtNatureOfProc.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : Me.txtTransporter.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : Me.txtTruckNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            Me.chkAnnexure.Enabled = False : Me.frmOption.Enabled = True : Me.optNagare.Checked = True : lblPendingDays.Text = ""
        Else
            Me.txtGrinNo.Text = "" : Me.txtDocumentNo.Text = "" : Me.txtLocationCode.Text = ""
            Me.lblLocDesc.Text = "" : Me.Lbl57F4No.Text = "" : Me.lblCustName.Text = ""
            SetGrid()
            Me.txtDocumentNo.Enabled = True : Me.cmdDOcHelp.Enabled = True : Me.MainGrid.Enabled = True
            Me.txtDocumentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : Me.dtDocDate.Enabled = False
            Me.chkPieceMeal.Enabled = True : Me.dtpRemovalTime.Enabled = True : Me.dtpRemoval.Enabled = True : Me.chkRemoval.Enabled = True
            Me.lblLocked.Visible = True : Me.lblYesNo.Visible = True : Me.lblYesNo.Text = ""
            Me.txtNatureOfProc.Text = "" : Me.txtTransporter.Text = "" : Me.txtTruckNo.Text = ""
            Me.txtNatureOfProc.Enabled = False : Me.txtTransporter.Enabled = False : Me.txtTruckNo.Enabled = False
            Me.txtNatureOfProc.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : Me.txtTransporter.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : Me.txtTruckNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Me.CmdButtons_57F4.Enabled(2) = False
            Me.CmdButtons_57F4.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False : lblPendingDays.Text = ""
            Me.CmdButtons_57F4.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
            Me.dt57F4Date.Value = GetServerDate()
            Me.dtDocDate.Value = GetServerDate() : Me.chkAnnexure.Enabled = True
            Me.chkAnnexure.CheckState = False : Me.chkPieceMeal.CheckState = False
        End If
        'To set grid property
        SetGrid()
        SetKanbanGrid()
        Me.cmdKANBANCancel.Enabled = True : Me.cmdKANBANOK.Enabled = True
        Me.KANBANGrid.Enabled = True : Me.optKanban.Enabled = True : Me.optKanbanDate.Enabled = True
        Me.dtpRemoval.Format = DateTimePickerFormat.Custom
        Me.dtpRemoval.CustomFormat = gstrDateFormat
        Me.dtpRemoval.Value = GetServerDate()
        'Me.dtpRemovalTime.Value = VB6.Format(TimeOfDay, "MM:hh")
        Me.dtpRemovalTime.Text = TimeOfDay
        ' Code end here
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0046_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        On Error GoTo ErrHandler
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
            If Me.CmdButtons_57F4.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                enmValue = ConfirmWindow(10055, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Or enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call CmdButtons_57F4_ButtonClick(CmdButtons_57F4, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE))
                    ElseIf enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Then
                        gblnCancelUnload = False
                        gblnFormAddEdit = False
                    End If
                Else
                    gblnCancelUnload = True
                    gblnFormAddEdit = True
                    Me.CmdButtons_57F4.Focus()
                End If
            Else
                Me.Dispose()
                Exit Sub
            End If
        End If
        'Checking The Status
        If gblnCancelUnload = True Then eventArgs.Cancel = 1
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTTRN0046_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '-------------------------------------------------------------------------------------------
        ' Author        : Sourabh Khatri
        ' Arguments     : Cancel as Integer
        ' Return Value  : NIL
        ' Function      : RemoveFormNameFromWindowList
        ' Release form Object Memory from Database.
        ' Created On    :18 jan 2006
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        objInvoicePrint = Nothing
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
        ' Created On    :18 Jan 2006
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Me.ToolTip1.SetToolTip(Me.txtDocumentNo, "Document No")
        Me.ToolTip1.SetToolTip(Me.txtGrinNo, "Grin no")
        Me.ToolTip1.SetToolTip(Me.txtLocationCode, "Location code of type 'M'  or 'P'")
        Me.ToolTip1.SetToolTip(Me.dt57F4Date, "57F4 Challan Date")
        Me.ToolTip1.SetToolTip(Me.dtDocDate, " Document Date (Default current date)")
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SetGrid()
        '-------------------------------------------------------------------------------------------
        ' Author        : Sourabh Khatri
        ' Return Value  : NIL
        ' Function      : Set grid property
        ' Created On    :18 Jan 2006
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim blnMode As Boolean
        With Me.MainGrid
            .MaxRows = 0
            .Row = .MaxRows
            .MaxCols = EnumGrid.KanbanNo
            If Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                blnMode = True
            Else
                blnMode = False
            End If
            .Col = EnumGrid.CheckBox : .Text = "Check"
            .TypeTextWordWrap = True
            .Col = EnumGrid.ItemCode : .Text = "Item Code"
            .set_ColWidth(EnumGrid.ItemCode, 15)
            .Col = EnumGrid.ItemDescription : .Text = "Item Description"
            .set_ColWidth(EnumGrid.ItemDescription, 25)
            .Col = EnumGrid.CUSTDRGNO : .Text = "Customer Part Code"
            .set_ColWidth(EnumGrid.CUSTDRGNO, 15)
            .Col = EnumGrid.ReceiveQty : .Text = "Receive Quantity"
            .set_ColWidth(EnumGrid.ReceiveQty, 15)
            .Col = EnumGrid.DispatchQty : .Text = "Dispatch Quantity"
            .set_ColWidth(EnumGrid.DispatchQty, 15)
            .Col = EnumGrid.BalanceQty : .Text = "Balance Quantity"
            .set_ColWidth(EnumGrid.BalanceQty, 15)
            .Col = EnumGrid.StockQty : .Text = "Stock Quantity"
            .set_ColWidth(EnumGrid.BalanceQty, 15)
            .Col = EnumGrid.InvoiceQty : .Text = "Challan Quantity"
            .set_ColWidth(EnumGrid.InvoiceQty, 15)
            .Col = EnumGrid.KanbanNo : .Text = "Kanban No"
            .set_ColWidth(EnumGrid.KanbanNo, 20)
            .set_RowHeight(0, 18)
            If blnMode = False Then
                .Col = EnumGrid.CheckBox
                .ColHidden = True
                .Col = EnumGrid.ReceiveQty
                .ColHidden = True
                .Col = EnumGrid.DispatchQty
                .ColHidden = True
                .Col = EnumGrid.StockQty
                .ColHidden = True
            Else
                .Col = EnumGrid.CheckBox
                .ColHidden = False
                .Col = EnumGrid.ReceiveQty
                .ColHidden = False
                .Col = EnumGrid.DispatchQty
                .ColHidden = False
                .Col = EnumGrid.StockQty
                .ColHidden = False
            End If
            If Me.optNagare.Checked = False Then
                .Col = EnumGrid.KanbanNo
                .ColHidden = True
            Else
                .Col = EnumGrid.KanbanNo
                .ColHidden = False
            End If
            If blnMode = True Then
                Call AddBlankRow()
            End If
            .ColsFrozen = EnumGrid.ItemCode
        End With
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Sub AddBlankRow()
        '-------------------------------------------------------------------------------------------
        ' Author        : Sourabh Khatri
        ' Return Value  : NIL
        ' Function      : Add new row
        '
        ' Created On    :18 Jan 2006
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        With Me.MainGrid
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            If Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                .Col = EnumGrid.CheckBox
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                .TypeCheckCenter = True
                .Col = EnumGrid.ItemCode
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                .TypeMaxEditLen = 16
                .Col = EnumGrid.ItemDescription
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = EnumGrid.ReceiveQty
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatMax = "99999999.99"
                .TypeFloatMin = "0.00"
                .TypeFloatDecimalPlaces = 2
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Col = EnumGrid.DispatchQty
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatMax = "99999999.99"
                .TypeFloatMin = "0.00"
                .TypeFloatDecimalPlaces = 2
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Col = EnumGrid.BalanceQty
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatMax = "99999999.99"
                .TypeFloatMin = "0.00"
                .TypeFloatDecimalPlaces = 2
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Col = EnumGrid.StockQty
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatMax = "99999999.99"
                .TypeFloatMin = "0.00"
                .TypeFloatDecimalPlaces = 2
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Col = EnumGrid.InvoiceQty
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatMax = "99999999.99"
                .TypeFloatMin = "0.00"
                .TypeFloatDecimalPlaces = 2
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Col = EnumGrid.KanbanNo
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "Kanban No"
                If Me.optNagare.Checked = False Then
                    .Col = EnumGrid.KanbanNo
                    .ColHidden = True
                End If
            ElseIf Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                .Col = EnumGrid.ItemCode
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                .TypeMaxEditLen = 16
                .Col = EnumGrid.ItemDescription
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = EnumGrid.BalanceQty
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatMax = "99999999.99"
                .TypeFloatMin = "0.00"
                .TypeFloatDecimalPlaces = 2
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Col = EnumGrid.InvoiceQty
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatMax = "99999999.99"
                .TypeFloatMin = "0.00"
                .TypeFloatDecimalPlaces = 2
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                If Me.optNagare.Checked = True Then
                    .Col = EnumGrid.KanbanNo
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                    .TypeButtonText = "Kanban No"
                End If
            End If
            .set_RowHeight(.MaxRows, 15)
        End With
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub optKanban_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optKanban.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            With Me.KANBANGrid
                .SortBy = FPSpreadADO.SortByConstants.SortByRow
                'Converted Comment                .SortKey(1) = 1
                'Converted Comment                .SortKeyOrder(1) = FPSpreadADO.SortKeyOrderConstants.SortKeyOrderAscending
                .Row = 1
                .Row2 = .MaxRows
                .Col = 1
                .Col2 = .MaxCols
                .Action = FPSpreadADO.ActionConstants.ActionSort
            End With
            Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        End If
    End Sub
    Private Sub optKanbanDate_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optKanbanDate.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            With Me.KANBANGrid
                .SortBy = FPSpreadADO.SortByConstants.SortByRow
                'Converted Comment            .SortKey(1) = 2
                'Converted Comment            .SortKeyOrder(1) = FPSpreadADO.SortKeyOrderConstants.SortKeyOrderAscending
                .Row = 1
                .Row2 = .MaxRows
                .Col = 1
                .Col2 = .MaxCols
                .Action = FPSpreadADO.ActionConstants.ActionSort
            End With
            Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        End If
    End Sub
    Private Sub optNagare_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optNagare.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            If Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                mblnNagare = True
                With Me.MainGrid
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = EnumGrid.KanbanNo
                    .Col2 = EnumGrid.KanbanNo
                    .BlockMode = True
                    .ColHidden = False
                    .BlockMode = False
                    .Row = 1
                    .Col = EnumGrid.CheckBox
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                End With
            End If
            Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        End If
    End Sub
    Private Sub OptNagare_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles optNagare.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If (Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT) And KeyAscii = System.Windows.Forms.Keys.Return Then
            mblnNagare = True
            With Me.MainGrid
                .Row = 1
                .Row2 = .MaxRows
                .Col = EnumGrid.KanbanNo
                .Col2 = EnumGrid.KanbanNo
                .BlockMode = True
                .ColHidden = False
                .BlockMode = False
                .Row = 1
                .Col = EnumGrid.CheckBox
                .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
            End With
        End If
        GoTo EventExitSub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub optwithOutNagare_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optwithOutNagare.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            If Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                mblnNagare = False
                With Me.MainGrid
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = EnumGrid.KanbanNo
                    .Col2 = EnumGrid.KanbanNo
                    .BlockMode = True
                    .ColHidden = True
                    .BlockMode = False
                    .Row = 1
                    .Col = EnumGrid.CheckBox
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                End With
            End If
            Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        End If
    End Sub
    Private Sub optwithOutNagare_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles optwithOutNagare.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If (Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT) And KeyAscii = System.Windows.Forms.Keys.Return Then
            mblnNagare = False
            With Me.MainGrid
                .Row = 1
                .Row2 = .MaxRows
                .Col = EnumGrid.KanbanNo
                .Col2 = EnumGrid.KanbanNo
                .BlockMode = True
                .ColHidden = True
                .BlockMode = False
                .Row = 1
                .Col = EnumGrid.CheckBox
                .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
            End With
        End If
        GoTo EventExitSub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtDocumentNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDocumentNo.TextChanged
        On Error GoTo ErrHandler
        If Len(Me.txtDocumentNo.Text) = 0 Then
            Call Initialize_controls()
        Else
            Call GenerateBeep(Len(Trim(Me.txtDocumentNo.Text)), 8)
        End If
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtDocumentNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtDocumentNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            Call cmdDOcHelp_Click(cmdDOcHelp, New System.EventArgs())
        End If
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub TxtDocumentNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDocumentNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Call KeyPressValidation(KeyAscii, 1)
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            TxtDocumentNo_Validating(txtDocumentNo, New System.ComponentModel.CancelEventArgs(False))
        End If
        GoTo EventExitSub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtDocumentNo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDocumentNo.Leave
        On Error GoTo ErrHandler
        'Call TxtDocumentNo_Validate(False)
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub TxtDocumentNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtDocumentNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim rsobject As New ClsResultSetDB
        If txtDocumentNo.Text <> "" Then
            rsobject.GetResult("Select Doc_No from Mkt_57F4Challan_Hdr where doc_no = '" & Trim(Me.txtDocumentNo.Text) & "' and Unit_Code = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsobject.RowCount > 0 Then
                Call ShowDataInViewMode()
                Me.CmdButtons_57F4.Focus()
            Else
                MsgBox("Invalid Document Number", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, ResolveResString(100))
                Me.txtDocumentNo.Text = ""
                Me.txtDocumentNo.Focus()
                GoTo EventExitSub
            End If
        Else
            Me.CmdButtons_57F4.Focus()
        End If
        rsobject = Nothing
        GoTo EventExitSub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtGRINNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtGrinNo.TextChanged
        On Error GoTo ErrHandler
        If Trim(Me.txtGrinNo.Text) = "" Then
            Call Initialize_controls()
        Else
            Call GenerateBeep(Len(Trim(Me.txtGrinNo.Text)), 8)
        End If
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtGRINNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtGrinNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            Call cmdGRINHelp_Click(cmdGrinHelp, New System.EventArgs())
        End If
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtGRINNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtGrinNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Call KeyPressValidation(KeyAscii, 1)
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            Call txtGRINNo_Validating(txtGrinNo, New System.ComponentModel.CancelEventArgs(False))
        End If
        GoTo EventExitSub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtGRINNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtGrinNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim rsobject As New ClsResultSetDB
        Cancel = False
        If Trim(Me.txtGrinNo.Text) <> "" Then
            rsobject.GetResult("Select doc_No,vendor_name from Mkt_57F4GrinNo('" & gstrUNITID & "') where doc_No = '" & Me.txtGrinNo.Text & "' and Unit_Code = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsobject.RowCount > 0 Then
                Me.lblCustName.Text = rsobject.GetValue("Vendor_Name")
                rsobject.GetResult("Select isnull(invoice_No,0) invoice_No,invoice_Date,vendor_Code from grn_hdr where doc_no = '" & Me.txtGrinNo.Text & "' and Unit_Code = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rsobject.RowCount > 0 Then
                    Me.Lbl57F4No.Text = rsobject.GetValue("Invoice_No")
                    Me.dt57F4Date.Value = rsobject.GetValue("invoice_Date")
                    Me.lblCustCode.Text = rsobject.GetValue("vendor_Code")
                    lblPendingDays.Text = CStr(180 - DateDiff(Microsoft.VisualBasic.DateInterval.Day, Me.dt57F4Date.Value, Me.dtDocDate.Value))
                End If
                Me.txtLocationCode.Focus()
            Else
                MsgBox("Invalid Grin No", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                Cancel = True
            End If
            rsobject.ResultSetClose()
        End If
        rsobject = Nothing
        GoTo EventExitSub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtLocationCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLocationCode.TextChanged
        On Error GoTo ErrHandler
        If Trim(Me.txtLocationCode.Text) = "" Then
            Call Initialize_controls()
        Else
            Call GenerateBeep(Len(Trim(Me.txtLocationCode.Text)), 4)
        End If
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLocationCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtLocationCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
                Call cmdlocationhelp_Click(cmdLocationHelp, New System.EventArgs())
            End If
        End If
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLocationCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLocationCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Call KeyPressValidation(KeyAscii, 2)
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            Call txtLocationCode_Validating(txtLocationCode, New System.ComponentModel.CancelEventArgs(False))
        End If
        GoTo EventExitSub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub KeyPressValidation(ByRef KeyAscii As Short, ByRef ALLOWMODE As Short)
        On Error GoTo ErrHandler
        If ALLOWMODE = 1 Then
            If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Then
                KeyAscii = KeyAscii
            Else
                KeyAscii = 0
            End If
        Else
            If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Then
                KeyAscii = KeyAscii
            Else
                KeyAscii = 0
            End If
        End If
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub GenerateBeep(ByRef intlen As Short, ByRef intValidLen As Short)
        On Error GoTo ErrHandler
        If intlen > intValidLen Then
            Beep()
        End If
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLocationCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtLocationCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim rsobject As New ClsResultSetDB
        If Trim(Me.txtLocationCode.Text) <> "" Then
            strsql = "Select location_Code,description from location_mst where location_code = '" & Trim(Me.txtLocationCode.Text) & "' and loc_type in ('M','P','S') and Unit_Code = '" & gstrUNITID & "'"
            rsobject.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsobject.RowCount <= 0 Then
                Call MsgBox(ResolveResString(10236), MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            Else
                Me.lblLocDesc.Text = rsobject.GetValue("Description")
                Call ShowItemsInGrid((Me.txtGrinNo.Text))
                Me.txtNatureOfProc.Focus()
            End If
            rsobject.ResultSetClose()
        End If
        rsobject = Nothing
        GoTo EventExitSub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub ShowItemsInGrid(ByRef strGRINNo As String)
        On Error GoTo ErrHandler
        Dim rsobject As New ClsResultSetDB
        Dim Intcounter As Short
        Dim AcceptQty As Double
        Dim DisQty As Double
        Dim intC As Short
        rsobject.GetResult("Select A.Item_Code,description,accepted_Quantity,Isnull(DisQty,0) DisQty,Cur_Bal,isnull(Cust_drgno,'') as Cust_Item_Code from Mkt_57F4GrinItemDtl(" & Val(strGRINNo) & " , '" & gstrUNITID & "') A inner join cust_ord_hdr C on  a.UNIT_CODE = c.UNIT_CODE and c.account_Code = '" & Me.lblCustCode.Text & "' and A.Unit_Code = '" & gstrUNITID & "' and active_Flag = 'A' inner join cust_ord_dtl D on C.Account_Code = D.Account_Code and C.Cust_ref = D.Cust_Ref and C.Amendment_no = D.Amendment_no and C.UNIT_CODE = D.UNIT_CODE  and  D.UNIT_CODE = A.UNIT_CODE and  D.Item_Code = A.Item_Code and d.active_flag='A' left outer join itembal_mst B on  A.Item_Code = B.Item_Code and A.UNIT_CODE = B.UNIT_CODE  and B.Location_Code = '" & Me.txtLocationCode.Text & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsobject.RowCount > 0 Then
            rsobject.MoveFirst() : Intcounter = 1
            With Me.MainGrid
                'If mode is add then set grid at 0
                If Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    .MaxRows = 0
                End If
                While Not rsobject.EOFRecord
                    If Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                        For intC = 1 To .MaxRows
                            .Row = intC
                            .Col = EnumGrid.ItemCode
                            If StrComp(Trim(.Text), rsobject.GetValue("Item_Code"), CompareMethod.Text) = 0 Then
                                rsobject.MoveNext() : Intcounter = Intcounter + 1
                                GoTo loopend
                            End If
                        Next
                    End If
                    AddBlankRow()
                    .Row = .MaxRows
                    .Col = EnumGrid.ItemCode
                    .Text = rsobject.GetValue("Item_Code")
                    .Col = EnumGrid.ItemDescription
                    .Text = rsobject.GetValue("Description")
                    .Col = EnumGrid.CUSTDRGNO
                    .Text = rsobject.GetValue("Cust_Item_Code")
                    AcceptQty = rsobject.GetValue("Accepted_Quantity")
                    .Col = EnumGrid.ReceiveQty : .Text = CStr(AcceptQty)
                    DisQty = rsobject.GetValue("DisQty")
                    .Col = EnumGrid.DispatchQty
                    .Text = rsobject.GetValue("DisQty")
                    .Col = EnumGrid.StockQty
                    .Text = rsobject.GetValue("Cur_Bal")
                    .Col = EnumGrid.BalanceQty : .Text = CStr(AcceptQty - DisQty)
                    .Col = EnumGrid.InvoiceQty = 0
                    rsobject.MoveNext()
                    Intcounter = Intcounter + 1
loopend:
                End While
            End With
        Else
            MsgBox("No data available to display.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            Exit Sub
        End If
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function ValidDataBeforeSave() As Boolean
        On Error GoTo ErrHandler
        Dim blnFlag As Boolean
        Dim strError As String
        Dim lstControl As System.Windows.Forms.Control
        Dim Intcounter As Short
        Dim intCounter1 As Short
        Dim rsStock As New ClsResultSetDB
        blnFlag = True
        Intcounter = 0
        ValidDataBeforeSave = True
        Dim blnCheck As Boolean
        Dim StrItemCode As String
        Dim dblInvQty As Double
        Dim dblKanbanQty As Double
        If Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            'Check Grin No
            If Me.txtGrinNo.Text = "" Then
                Intcounter = Intcounter + 1
                strError = strError & vbCrLf & Intcounter & ". Grin No"
                If lstControl Is Nothing Then
                    lstControl = Me.txtGrinNo
                End If
                blnFlag = False
            End If
            'Check Grin No
            If Me.txtLocationCode.Text = "" Then
                Intcounter = Intcounter + 1
                strError = strError & vbCrLf & Intcounter & ". Location Code"
                If lstControl Is Nothing Then
                    lstControl = Me.txtLocationCode
                End If
                blnFlag = False
            End If
            'Date can't be less then current date or greater then tomorrow date
            If DateDiff(Microsoft.VisualBasic.DateInterval.Day, GetServerDate(), Me.dtDocDate.Value) > 0 Or DateDiff(Microsoft.VisualBasic.DateInterval.Day, GetServerDate(), Me.dtDocDate.Value) < -1 And Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                Intcounter = Intcounter + 1
                strError = strError & vbCrLf & Intcounter & ". Date can't less the current date or greater then tomorrow date"
                If lstControl Is Nothing Then
                    lstControl = Me.dtDocDate
                End If
                blnFlag = False
            End If
            'Validation for GRID
            blnCheck = False
            With Me.MainGrid
                For intCounter1 = 1 To .MaxRows
                    .Row = intCounter1
                    .Col = EnumGrid.CheckBox
                    If System.Math.Abs(Val(.Value)) = 1 Then
                        blnCheck = True
                    End If
                Next
            End With
            If blnCheck = False Then
                Intcounter = Intcounter + 1
                strError = strError & vbCrLf & Intcounter & ". Check atleast one row"
                If lstControl Is Nothing Then
                    lstControl = Me.MainGrid
                End If
                blnFlag = False
            End If
            With Me.MainGrid
                For intCounter1 = 1 To .MaxRows
                    .Row = intCounter1
                    .Col = EnumGrid.CheckBox
                    If System.Math.Abs(Val(.Value)) = 1 Then
                        .Col = EnumGrid.InvoiceQty
                        If Val(.Text) <= 0 Then
                            Intcounter = Intcounter + 1
                            .Col = EnumGrid.ItemCode
                            strError = strError & vbCrLf & Intcounter & ". 0 is not a valid invoice quantity for item " & .Text
                            .Col = EnumGrid.InvoiceQty
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                            blnFlag = False
                        Else
                            .Col = EnumGrid.ItemCode
                            If Val(SelectDataFromTable("cur_bal", "itembal_mst", " item_Code = '" & .Text & "' and location_Code = '" & Me.txtLocationCode.Text & "' and Unit_Code = '" & gstrUNITID & "'")) <= 0 Then
                                Intcounter = Intcounter + 1
                                strError = strError & vbCrLf & Intcounter & ". Stock in hand is less the invoice quantity for item " & .Text
                                .Col = EnumGrid.InvoiceQty
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                blnFlag = False
                            End If
                        End If
                    End If
                Next
            End With
            If Me.optNagare.Checked = True Then
                With Me.MainGrid
                    For intCounter1 = 1 To .MaxRows
                        .Row = intCounter1
                        .Col = EnumGrid.CheckBox
                        If System.Math.Abs(Val(.Value)) = 1 Then
                            .Col = EnumGrid.ItemCode : StrItemCode = .Text
                            .Col = EnumGrid.InvoiceQty : dblInvQty = Val(.Text)
                            rsKanBan.Filter = "ItemCode = '" & StrItemCode & "'"
                            If rsKanBan.EOF Then
                                Intcounter = Intcounter + 1
                                strError = strError & vbCrLf & Intcounter & ". Select KANBAN detail against item " & StrItemCode
                                .Col = EnumGrid.KanbanNo
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                blnFlag = False
                            Else
                                rsKanBan.MoveFirst()
                                dblKanbanQty = 0
                                While Not rsKanBan.EOF
                                    dblKanbanQty = dblKanbanQty + rsKanBan.Fields("kanbanqty").Value
                                    rsKanBan.MoveNext()
                                End While
                                If dblInvQty <> dblKanbanQty Then
                                    Intcounter = Intcounter + 1
                                    strError = strError & vbCrLf & Intcounter & ". Kanban Qty is not equal to Invoice Qty against Item " & StrItemCode
                                    .Col = EnumGrid.KanbanNo
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                    blnFlag = False
                                End If
                            End If
                            rsKanBan.Filter = ADODB.FilterGroupEnum.adFilterNone
                        End If
                    Next
                End With
            End If
        Else
        End If
        If blnFlag = False Then
            MsgBox("Following are invalid or can not be entered" & vbCrLf & strError, MsgBoxStyle.Information, ResolveResString(100))
            'lstControl.SetFocus
            ValidDataBeforeSave = False
        End If
        Exit Function 'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub ShowDataInViewMode()
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim strDocNo As String
        Dim rsobject As New ClsResultSetDB
        Dim Intcounter As Short
        strDocNo = Me.txtDocumentNo.Text
        'Fetch header Data
        strsql = " Select distinct grin_no,location_Code,against_nagare,vendor_Code,vendor_Name= " & " Case doc_category " & " When 'U' Then (Select cust_name from customer_mst where customer_Code = grn_hdr.vendor_Code and Unit_Code = '" & gstrUNITID & "') " & " When 'J' Then (Select vendor_name from vendor_mst where vendor_Code = grn_hdr.vendor_Code and Unit_Code = '" & gstrUNITID & "') " & " End " & " ,Doc_Category,invoice_No,invoice_Date,mkt_57f4challan_hdr.doc_Date,cancel_flag,invoice_lock,isnull(NatureOfProc,'') NatureOfProc,isnull(Trans_Name,'') Trans_Name,isnull(Truck_No,'') Truck_No  " & " From mkt_57f4challan_hdr inner join grn_hdr on grn_hdr.doc_no = mkt_57f4challan_hdr.grin_no and grn_hdr.Unit_Code = mkt_57f4challan_hdr.Unit_Code where mkt_57f4challan_hdr.doc_no = '" & Me.txtDocumentNo.Text & "' and mkt_57f4challan_hdr.Unit_Code = '" & gstrUNITID & "'"
        rsobject.GetResult(strsql)
        If rsobject.RowCount > 0 Then
            Me.txtGrinNo.Text = rsobject.GetValue("Grin_no")
            Me.Lbl57F4No.Text = rsobject.GetValue("Invoice_no")
            Me.lblCustName.Text = rsobject.GetValue("vendor_Name")
            Me.lblCustCode.Text = rsobject.GetValue("Vendor_Code")
            Me.txtLocationCode.Text = rsobject.GetValue("location_Code")
            Me.dt57F4Date.Value = rsobject.GetValue("invoice_Date")
            Me.dtDocDate.Value = rsobject.GetValue("Doc_Date")
            If rsobject.GetValue("against_nagare") = 0 Then
                Me.optwithOutNagare.Checked = True
            Else
                Me.optNagare.Checked = True
                GenerateDisRecordSet()
                SetGrid()
                SetKanbanGrid()
            End If
            Me.lblYesNo.Text = IIf(rsobject.GetValue("invoice_lock") = 1, "Yes", "No")
            Me.txtNatureOfProc.Text = rsobject.GetValue("NatureOfProc")
            Me.txtTransporter.Text = rsobject.GetValue("Trans_Name")
            Me.txtTruckNo.Text = rsobject.GetValue("Truck_No")
        Else
            MsgBox("Selected document number does not exist ", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            Exit Sub
        End If
        Me.CmdButtons_57F4.Enabled(2) = True
        Me.CmdButtons_57F4.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
        Me.CmdButtons_57F4.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
        'Fetch Detail Data
        strsql = " Select Doc_No,A.Item_Code,description,Cust_Drgno,Balance_Qty,Invoice_Qty from mkt_57f4challan_dtl A inner join item_mst B on " & " a.item_Code = b.item_Code and a.Unit_code = b.Unit_code where doc_no = '" & strDocNo & "' and a.Unit_code = '" & gstrUNITID & "'"
        rsobject.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim rskanban1 As New ClsResultSetDB
        If rsobject.RowCount > 0 Then
            rsobject.MoveFirst() : Intcounter = 1
            With Me.MainGrid
                .MaxRows = 0
                While Not rsobject.EOFRecord
                    AddBlankRow()
                    .Row = Intcounter
                    .Col = EnumGrid.CheckBox
                    .Value = True
                    .Col = EnumGrid.ItemCode
                    .Text = rsobject.GetValue("Item_Code")
                    .Col = EnumGrid.ItemDescription
                    .Text = rsobject.GetValue("description")
                    .Col = EnumGrid.CUSTDRGNO
                    .Text = rsobject.GetValue("Cust_Drgno")
                    .Col = EnumGrid.BalanceQty
                    .Text = rsobject.GetValue("Balance_Qty")
                    .Col = EnumGrid.InvoiceQty
                    .Text = rsobject.GetValue("Invoice_Qty")
                    If Me.optNagare.Checked = True Then
                        Call rskanban1.GetResult(" Select KanBan_No,Mkt_57F4ChallanKanBan_Dtl.Quantity,sch_Date,DiffQty,UNLOC,USLOC,Sch_Time From Mkt_57F4ChallanKanBan_Dtl left outer join Mkt_57F4KANBANNO('" & rsobject.GetValue("Item_Code") & "','" & Me.lblCustCode.Text & "', '" & gstrUNITID & "') on Mkt_57F4ChallanKanBan_Dtl.kanban_no = kanbanno and Mkt_57F4ChallanKanBan_Dtl.unit_code = unit_code where doc_no = '" & Me.txtDocumentNo.Text & "' and Mkt_57F4ChallanKanBan_Dtl.item_Code = '" & rsobject.GetValue("Item_Code") & "' and Mkt_57F4ChallanKanBan_Dtl.unit_code = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rskanban1.RowCount > 0 Then
                            rskanban1.MoveFirst()
                            While Not rskanban1.EOFRecord
                                rsKanBan.AddNew()
                                rsKanBan.Fields("itemcode").Value = rsobject.GetValue("Item_Code")
                                rsKanBan.Fields("kanbanno").Value = rskanban1.GetValue("KanBan_No")
                                rsKanBan.Fields("kanbanqty").Value = rskanban1.GetValue("Quantity")
                                If Len(rskanban1.GetValue("sch_Date")) = 0 Then
                                    rsKanBan.Fields("SchDate").Value = getDateForDB(GetServerDate())
                                Else
                                    rsKanBan.Fields("SchDate").Value = getDateForDB(rskanban1.GetValue("sch_Date"))
                                End If
                                rsKanBan.Fields("UnLoc").Value = rskanban1.GetValue("UnLoc")
                                rsKanBan.Fields("UsLoc").Value = rskanban1.GetValue("UsLoc")
                                rsKanBan.Fields("BalanceQty").Value = Val(rskanban1.GetValue("DiffQty"))
                                rsKanBan.Update()
                                rskanban1.MoveNext()
                            End While
                        End If
                        rskanban1 = Nothing
                    Else
                        .Col = EnumGrid.KanbanNo
                        .ColHidden = True
                    End If
                    rsobject.MoveNext()
                    Intcounter = Intcounter + 1
                End While
            End With
        Else
            MsgBox("Details of document number does not exist ", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            Exit Sub
        End If
        rsobject = Nothing
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Sub SetKanbanGrid()
        On Error GoTo ErrHandler
        With Me.KANBANGrid
            .MaxRows = 0
            .Row = 0
            .MaxCols = KANBANGrid1.SelectedQty
            .Col = KANBANGrid1.KanbanNo
            .Text = "Kanban No."
            .set_ColWidth(KANBANGrid1.KanbanNo, 15)
            .Col = KANBANGrid1.SchDate
            .Text = "Schedule Date"
            .set_ColWidth(KANBANGrid1.SchDate, 10)
            .Col = KANBANGrid1.SChTime
            .Text = "Schedule Time"
            .set_ColWidth(KANBANGrid1.SChTime, 7)
            .Col = KANBANGrid1.UNLoc
            .Text = "UNLOC"
            .set_ColWidth(KANBANGrid1.UNLoc, 7)
            .Col = KANBANGrid1.USLoc
            .Text = "USLOC"
            .set_ColWidth(KANBANGrid1.USLoc, 7)
            .Col = KANBANGrid1.BalanceQty
            .Text = "Balance Qty"
            .set_ColWidth(KANBANGrid1.BalanceQty, 10)
            .Col = KANBANGrid1.SelectedQty
            .Text = "KanBan Qty"
            .set_ColWidth(KANBANGrid1.SelectedQty, 10)
            .set_RowHeight(0, 15)
            .ColsFrozen = 2
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Sub AddBlankKanBanRow()
        On Error GoTo ErrHandler
        With Me.KANBANGrid
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = KANBANGrid1.KanbanNo
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = KANBANGrid1.SchDate
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeDate
            .TypeDateCentury = True
            .TypeDateSeparator = Asc("/")
            .Col = KANBANGrid1.SChTime
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeTime
            .TypeTimeSeparator = Asc(":")
            .TypeTimeSeconds = False
            .Col = KANBANGrid1.UNLoc
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = KANBANGrid1.USLoc
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = KANBANGrid1.BalanceQty
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatMax = "99999999.99"
            .TypeFloatMin = "0.00"
            .Col = KANBANGrid1.SelectedQty
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
            .TypeFloatMax = "99999999.99"
            .TypeFloatMin = "0.00"
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub GenerateDisRecordSet()
        If Not rsKanBan Is Nothing Then
            rsKanBan = Nothing
        End If
        rsKanBan = New ADODB.Recordset
        rsKanBan.Fields.Append("ItemCode", ADODB.DataTypeEnum.adVarChar, 30)
        rsKanBan.Fields.Append("KanbanNo", ADODB.DataTypeEnum.adVarChar, 30)
        rsKanBan.Fields.Append("SchDate", ADODB.DataTypeEnum.adDate)
        rsKanBan.Fields.Append("SchTime", ADODB.DataTypeEnum.adDBTime)
        rsKanBan.Fields.Append("UnLoc", ADODB.DataTypeEnum.adVarChar, 10)
        rsKanBan.Fields.Append("UsLoc", ADODB.DataTypeEnum.adVarChar, 10)
        rsKanBan.Fields.Append("BalanceQty", ADODB.DataTypeEnum.adDouble)
        rsKanBan.Fields.Append("KanBanQty", ADODB.DataTypeEnum.adDouble)
        rsKanBan.Open()
    End Sub
    Private Function LockInvoice() As Boolean
        Dim strDocNo As String
        Dim strLocation_Code As String
        Dim StrItemCode As String
        Dim dblStockinHand As Double
        Dim dblInvoiceQty As Double
        Dim strKanBanNumber As String
        Dim dblKanbanQty As Double
        Dim rsobject As New ClsResultSetDB
        Dim rsKanBan As New ClsResultSetDB
        Dim rsCheck As New ClsResultSetDB
        Dim strString As String
        Dim StrCustomerCode As String
        On Error GoTo Errorhandler
        LockInvoice = False
        Dim blnCheck As Boolean
        'Saving of document number and location code into variable
        strDocNo = Trim(Me.txtDocumentNo.Text)
        strLocation_Code = Trim(Me.txtLocationCode.Text)
        StrCustomerCode = SelectDataFromTable("vendor_Code", "grn_hdr", " doc_no = '" & Me.txtGrinNo.Text & "' and Unit_Code = '" & gstrUNITID & "'")
        'Code for update lock flag
        blnCheck = False
        'Code for stock updation
        mP_Connection.BeginTrans()
        strString = "Select Item_Code,Cust_Drgno,Invoice_Qty From mkt_57f4challan_dtl where doc_no = '" & strDocNo & "' and doc_type = 109 and Unit_Code = '" & gstrUNITID & "'"
        rsobject.GetResult(strString, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsobject.RowCount <= 0 Then
            MsgBox("Records not available against document number " & strDocNo, MsgBoxStyle.Critical, ResolveResString(100))
            mP_Connection.RollbackTrans()
            LockInvoice = False
            Exit Function
        Else
            rsobject.MoveFirst()
            While Not rsobject.EOFRecord
                StrItemCode = rsobject.GetValue("Item_Code")
                dblInvoiceQty = rsobject.GetValue("invoice_Qty")
                If Val(SelectDataFromTable("cur_bal", "itembal_mst", " item_Code = '" & StrItemCode & "' and location_Code = '" & strLocation_Code & "' and Unit_Code = '" & gstrUNITID & "'")) < Val(CStr(dblInvoiceQty)) Then
                    MsgBox("Stock in hand for item " & StrItemCode & " is less than invoice Quantity", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, ResolveResString(100))
                    mP_Connection.RollbackTrans()
                    LockInvoice = False
                    Exit Function
                Else
                    rsKanBan.GetResult("Select Item_Code,KanBan_No,Quantity From mkt_57f4challankanban_dtl where doc_no = '" & strDocNo & "' and doc_type = 109 and item_Code = '" & StrItemCode & "' and Unit_Code = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If rsKanBan.RowCount > 0 Then
                        While Not rsKanBan.EOFRecord
                            mP_Connection.Execute("Update Dailymktschedule Set Despatch_qty = Despatch_qty + " & rsKanBan.GetValue("Quantity") & " where item_Code = '" & StrItemCode & "' and account_Code = '" & StrCustomerCode & "' and status = 1 and Unit_Code = '" & gstrUNITID & "' and trans_Date in (Select convert(varchar(11),sch_Date,106) from mkt_enagaredtl where item_Code = '" & StrItemCode & "' and kanbanno = '" & rsKanBan.GetValue("kanban_no") & "' and Unit_Code = '" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            rsKanBan.MoveNext()
                        End While
                    End If
                End If
                mP_Connection.Execute(" Update itembal_mst set cur_bal = cur_bal - " & dblInvoiceQty & " where location_Code = '" & strLocation_Code & "' and item_Code = '" & StrItemCode & "' and Unit_Code = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                rsobject.MoveNext()
            End While
            strString = "Update mkt_57f4challan_hdr Set Invoice_Lock = 1 where doc_no = '" & strDocNo & "' and Location_Code = '" & strLocation_Code & "' and doc_type = 109 and Unit_Code = '" & gstrUNITID & "'"
            mP_Connection.Execute(strString, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        End If
        blnCheck = True
        rsobject.ResultSetClose()
        rsKanBan.ResultSetClose()
        mP_Connection.CommitTrans()
        rsobject = Nothing
        rsKanBan = Nothing
        LockInvoice = True
        Exit Function
Errorhandler:  'The Error Handling Code Starts here
        mP_Connection.RollbackTrans()
        LockInvoice = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Sub PrintingInvoice()
        On Error GoTo ErrHandler
        objInvoicePrint.ConnectionString = gstrCONNECTIONSTRING
        objInvoicePrint.Connection()
        objInvoicePrint.FileName = gstrLocalCDrive & "EmproInv\InvoicePrint.txt"
        objInvoicePrint.BCFileName = gstrLocalCDrive & "EmproInv\BarCode.txt"
        objInvoicePrint.CompanyName = gstrCOMPANY
        objInvoicePrint.Address1 = gstr_RGN_ADDRESS1
        objInvoicePrint.Address2 = gstr_RGN_ADDRESS2
        'On Error Resume Next
        If Me.chkRemoval.CheckState = System.Windows.Forms.CheckState.Checked Then
            objInvoicePrint.Print_57F4Challan_SUNVAC(gstrUNITID, True, (Me.txtDocumentNo.Text), dtpRemoval.Text & " " & VB6.Format(dtpRemovalTime.Value.Hour, "00") & ":" & VB6.Format(dtpRemovalTime.Value.Minute, "00"))
        Else
            objInvoicePrint.Print_57F4Challan_SUNVAC(gstrUNITID, True, (Me.txtDocumentNo.Text))
        End If
        rtbInvoicePreview.LoadFile(objInvoicePrint.FileName)
        rtbInvoicePreview.BackColor = System.Drawing.Color.White
        cmdPrint.Image = My.Resources.ico231.ToBitmap
        cmdClose.Image = My.Resources.ico217.ToBitmap
        cmdPrint.Enabled = True : cmdClose.Enabled = True
        rtbInvoicePreview.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        ReplaceJunkCharacters()
        rtbInvoicePreview.Enabled = True
        rtbInvoicePreview.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        'Function called, if error occurred
    End Sub
    Private Sub ReplaceJunkCharacters()
        '----------------------------------------------------------------------------
        'Author         :   Arshad Ali
        'Argument       :   Non
        'Return Value   :   Non
        'Function       :   Removes all special characters used for formating from text file
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo Errorhandler
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(15), "") 'Remove Uncompress Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(18), "") 'Remove Decompress Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "G", "") 'Remove Bold Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "H", "") 'Remove DeBold Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(12), "") 'Remove DeUnderline Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "-1", "") 'Remove Underline Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "-0", "") 'Remove DeUnderline Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "W1", "") 'Remove DoubleWidth Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "W0", "") 'Remove DeDoubleWidth Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "M", "") 'Remove Middle Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "P", "") 'Remove DeMiddle Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "E", "") 'Remove Elite Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "F", "") 'Remove DeElite Character
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function SelectDataFromTable(ByRef mstrFieldName As String, ByRef mstrTableName As String, ByRef mstrCondition As String) As String
        '------------------------------------------------------------------------------
        'Created By     -   Sourabh Khatri
        'Description    -   Get Data from BackEnd
        '------------------------------------------------------------------------------
        Dim StrSQLQuery As String
        Dim GetDataFromTable As ClsResultSetDB
        On Error GoTo ErrHandler
        StrSQLQuery = "Select " & mstrFieldName & " From " & mstrTableName & " Where " & mstrCondition
        GetDataFromTable = New ClsResultSetDB
        If GetDataFromTable.GetResult(StrSQLQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic) Then
            If GetDataFromTable.GetNoRows > 0 Then
                SelectDataFromTable = GetDataFromTable.GetValue(mstrFieldName)
            Else
                SelectDataFromTable = ""
            End If
        Else
            SelectDataFromTable = ""
        End If
        GetDataFromTable.ResultSetClose()
        GetDataFromTable = Nothing
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Sub printBarCode(ByVal pstrFileName As String)
        'Author         :   Arshad Ali
        'Argument       :
        'Return Value   :
        'Function       :
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        Dim varTemp As Object
        Dim strString As String
        On Error GoTo ErrHandler
        'strString = App.Path + "\pdf-dot.bat BarCode.txt 4 2 2 1"
        strString = gstrLocalCDrive & "EmproInv\pdf-dot.bat " & gstrLocalCDrive & "EmproInv\BarCode.txt 4 2 2 1"
        strString = gstrLocalCDrive & "EmproInv\pdf-dot.bat " & pstrFileName & " 4 2 2 1"
        varTemp = Shell("cmd.exe /c " & strString)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtNatureOfProc_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtNatureOfProc.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If Not (KeyAscii = 32) Then
            Call KeyPressValidation(KeyAscii, 2)
        End If
        If KeyAscii = 13 Then
            Me.txtTransporter.Focus()
        End If
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtTransporter_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTransporter.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If Not (KeyAscii = 32) Then
            Call KeyPressValidation(KeyAscii, 2)
        End If
        If KeyAscii = 13 Then
            Me.txtTruckNo.Focus()
        End If
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtTruckNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTruckNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If Not (KeyAscii = 32) Then
            Call KeyPressValidation(KeyAscii, 2)
        End If
        If KeyAscii = 13 Then
            Me.optNagare.Focus()
        End If
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub KANBANGrid_EditChange(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles KANBANGrid.EditChange
        On Error GoTo ErrHandler
        Dim intCounter As Integer
        Dim dblTot As Double
        If e.col = KANBANGrid1.SelectedQty Then
            With Me.KANBANGrid
                'Validation for Balance Quantity Vs Selected Quantity
                .Row = .ActiveRow : .Col = KANBANGrid1.BalanceQty
                dblTot = Val(.Text)
                .Row = .ActiveRow : .Col = KANBANGrid1.SelectedQty
                If Val(.Text) > dblTot Then
                    .Text = dblTot
                End If
                dblTot = 0
                'Validation for Total Selected Quantity Vs Invoice Quantity
                For intCounter = 1 To .MaxRows
                    .Row = intCounter : .Col = KANBANGrid1.SelectedQty
                    dblTot = dblTot + Val(.Text)
                Next
                If dblTot > Val(Me.lblBalQty) Then
                    Call MsgBox("Total Kanban quantity can not exceed than invoice quantity", vbCritical + vbOKOnly, ResolveResString(100))
                    .Row = e.row : .Col = e.col : .Text = 0
                    Exit Sub
                Else
                    Me.lblKANBANTotal.Text = dblTot
                End If
            End With
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub KANBANGrid_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles KANBANGrid.KeyPressEvent
        On Error GoTo ErrHandler
        If Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            e.keyAscii = 0
        End If
        Dim dblbalQty As Double
        With Me.KANBANGrid
            If .ActiveCol = KANBANGrid1.SelectedQty Then
                .Row = .ActiveRow : .Col = KANBANGrid1.BalanceQty
                dblbalQty = Val(.Text)
            End If
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub MainGrid_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles MainGrid.ButtonClicked
        On Error GoTo ErrHandler
        Dim strItemCode As String
        Dim strCustDrgNo As String
        Dim dblbalQty As Double
        Dim strCustCode As String
        Dim rsobject As New ClsResultSetDB
        If e.col = EnumGrid.KanbanNo Then
            With Me.MainGrid
                .Row = e.row : .Col = EnumGrid.CheckBox
                If Val(.Value) = 0 And Me.CmdButtons_57F4.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    MsgBox("Please check row before selecting kanban no.", vbCritical, ResolveResString(100))
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                    Exit Sub
                End If
                .Row = e.row : .Col = EnumGrid.InvoiceQty : dblbalQty = Val(.Text)
                'Validate balance quantity
                If dblbalQty <= 0 Then
                    MsgBox("Invoice Quantity can not be 0", vbCritical, ResolveResString(100))
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                    Exit Sub
                End If
                'Show header date
                .Col = EnumGrid.ItemCode : strItemCode = Trim$(.Text)
                .Col = EnumGrid.CUSTDRGNO : strCustDrgNo = Trim$(.Text)
                Me.lblItem.Text = strItemCode : Me.lblCustPart.Text = strCustDrgNo
                Me.lblBalQty.Text = dblbalQty
                'Show Detail date
                rsKanBan.Filter = "itemCode = '" & strItemCode & "'"
                If Not rsKanBan.EOF Then
                    With Me.KANBANGrid
                        .MaxRows = 0 : Me.lblKANBANTotal.Text = 0
                        While Not rsKanBan.EOF
                            Call AddBlankKanBanRow()
                            .Row = .MaxRows
                            .Col = KANBANGrid1.KanbanNo : .Text = rsKanBan.Fields("kanbanno")
                            .Col = KANBANGrid1.SchDate : .Text = Format(rsKanBan.Fields("schDate"), gstrDateFormat)
                            .Col = KANBANGrid1.SChTime : .Text = Format(rsKanBan.Fields("SchTime"), "hh:mm")
                            .Col = KANBANGrid1.UNLoc : .Text = rsKanBan.Fields("unLoc")
                            .Col = KANBANGrid1.USLoc : .Text = rsKanBan.Fields("usLoc")
                            .Col = KANBANGrid1.BalanceQty : .Text = rsKanBan.Fields("balanceqty")
                            .Col = KANBANGrid1.SelectedQty : .Text = rsKanBan.Fields("KanbanQty")
                            Me.lblKANBANTotal.Text = Val(Me.lblKANBANTotal.Text) + Val(rsKanBan.Fields("KanbanQty"))
                            rsKanBan.MoveNext()
                        End While
                        Me.frmKANBAN.Visible = True
                        .Row = 1 : .Col = KANBANGrid1.SelectedQty : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                    End With
                Else
                    rsobject.GetResult("Select isnull(vendor_Code,'') vendor_Code from grn_hdr where doc_no = '" & Trim$(Me.txtGrinNo.Text) & "' and Unit_Code = '" & gstrUNITID & "'")
                    strCustCode = rsobject.GetValue("vendor_Code")
                    Dim cmdObject As New ADODB.Command
                    Dim Rs As New ADODB.Recordset
                    With cmdObject
                        .ActiveConnection = mP_Connection
                        .CommandTimeout = 0
                        .CommandText = "Select Kanbanno,DiffQty,UNLOC,USLOC,Sch_Date,Sch_Time from Mkt_57F4KANBANNO('" & strItemCode & "','" & strCustCode & "','" & gstrUNITID & "') where DiffQty > 0 and Unit_Code = '" & gstrUNITID & "'"
                        Rs = cmdObject.Execute
                    End With
                    If Not Rs.EOF Then
                        Rs.MoveFirst()
                        With Me.KANBANGrid
                            .MaxRows = 0 : Me.lblKANBANTotal.Text = 0
                            While Not Rs.EOF
                                Call AddBlankKanBanRow()
                                .Row = .MaxRows
                                .Col = KANBANGrid1.KanbanNo : .Text = Rs.Fields("kanbanno")
                                .Col = KANBANGrid1.SchDate : .Text = Rs.Fields("sch_Date") ', "dd/mm/yy")
                                .Col = KANBANGrid1.SChTime : .Text = Rs.Fields("Sch_Time")
                                .Col = KANBANGrid1.UNLoc : .Text = Rs.Fields("unLoc")
                                .Col = KANBANGrid1.USLoc : .Text = Rs.Fields("usLoc")
                                .Col = KANBANGrid1.BalanceQty : .Text = Val(Rs.Fields("DiffQty"))
                                .Col = KANBANGrid1.SelectedQty : .Text = "0.00"
                                Rs.MoveNext()
                            End While
                            Me.frmKANBAN.Visible = True
                            .Row = 1 : .Col = KANBANGrid1.SelectedQty : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                        End With
                    Else
                        Call MsgBox("No Pending Kanban exists for item " & strItemCode, vbInformation + vbOKOnly, ResolveResString(100))
                    End If
                    Rs.Close()
                    Rs = Nothing
                    cmdObject = Nothing
                End If
                rsKanBan.Filter = ADODB.FilterGroupEnum.adFilterNone
            End With
        End If
        rsobject = Nothing
        Exit Sub     'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub MainGrid_EditChange(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles MainGrid.EditChange
        Dim dblStockQty As Double
        Dim dblbalanceqty As Double
        Dim dblInvoiceQty As Double
        On Error GoTo ErrHandler
        If e.col = EnumGrid.InvoiceQty Then
            With Me.MainGrid
                .Row = e.row : .Col = EnumGrid.BalanceQty : dblbalanceqty = Val(.Text)
                .Row = e.row : .Col = EnumGrid.StockQty : dblStockQty = Val(.Text)
                .Row = e.row : .Col = EnumGrid.InvoiceQty : dblInvoiceQty = Val(.Text)
                .Row = e.row : .Col = EnumGrid.InvoiceQty
                If dblInvoiceQty > dblbalanceqty Then
                    .Text = dblbalanceqty
                    If dblStockQty < Val(.Text) Then
                        .Text = dblStockQty
                    End If
                    Exit Sub
                End If
                If Me.CmdButtons_57F4.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    If dblInvoiceQty > dblStockQty Then
                        .Text = dblStockQty
                    End If
                End If
            End With
        End If
        Exit Sub     'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub MainGrid_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles MainGrid.KeyPressEvent
        On Error GoTo ErrHandler
        If Me.MainGrid.ActiveCol = EnumGrid.InvoiceQty Or Me.MainGrid.ActiveCol = EnumGrid.CheckBox Or Me.MainGrid.ActiveCol = EnumGrid.KanbanNo Then
            e.keyAscii = e.keyAscii
        Else
            e.keyAscii = 0
        End If
        Exit Sub     'To prevent the execution of errhandler
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlFormHeader_57F4_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader_57F4.Click
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Set Property CheckFormName as per Rule book
        ' Datetime      : 16 april 2005
        '--------------------------------------------------------------------
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        Exit Sub    'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
End Class