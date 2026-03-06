Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0044
	Inherits System.Windows.Forms.Form
	'-----------------------------------------------------------------------
	' Copyright (c)     :MIND Ltd.
	' Form Name         :frmMKTTRN0044
	' Function Name     :Dispatch Slip
	' Created By        :Davinder singh
	' Created On        :10-Nov-2005
	' Modify Date       :NIL
    ' Revision History  :-
    '---------------------------------------------------------------------------------------
    'Modified by    :   Virendra Gupta
    'Modified ON    :   18/05/2011
    'Modified to support MultiUnit functionality
    '-----------------------------------------------------------------------
	Dim mintIndex As Short
	Dim mblnValidQty As Boolean
	Dim mstrDocNo As String
    Dim mblnStatus As Boolean
    Private Enum EnumGrid
        Item_Code = 1
        Item_CmdHlp = 2
        Item_Description = 3
        Cust_DrgNo = 4
        Cust_DrgnoHlp = 5
        Cust_Drgno_Description = 6
        item_qty = 7
    End Enum

    Private Sub CmdCustCodeHlp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdCustCodeHlp.Click
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To show the customer code help
        ' Created On    :10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo errHandler

        Dim strCustCode() As String
        strCustCode = Me.ctlEMPHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select customer_code,cust_name from customer_mst where Unit_Code = '" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))", "Customer Code Help")
        If Not (UBound(strCustCode) <= 0) Then
            If (Len(strCustCode(0)) >= 1) And strCustCode(0) = "0" Then
                MsgBox("No. Record Exist To Display", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            Else
                Me.TxtCustCode.Text = strCustCode(0)
                Me.LblCustName.Text = strCustCode(1)
                Me.TxtCustCode.Focus()
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub

    Private Sub cmdDocNoHlp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDocNoHlp.Click
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To show the customer code help
        ' Created On    :10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo errHandler

        Dim strDocNo() As String
        strDocNo = Me.ctlEMPHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select A.Doc_no,A.customer_code," & DateColumnNameInShowList("A.dispatch_dt") & " as dispatch_dt,B.Cust_name from dispatchslip_hdr A,customer_mst B where  A.customer_code=B.customer_code and A.Unit_code=B.Unit_code and A.Unit_Code = '" & gstrUNITID & "' and ((isnull(b.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= b.deactive_date))", "Document No. Help")
        If Not (UBound(strDocNo) <= 0) Then
            If (Len(strDocNo(0)) >= 1) And strDocNo(0) = "0" Then
                MsgBox("No. Record Exist To Display", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            Else
                Me.TxtDocNo.Text = strDocNo(0)
                Me.TxtCustCode.Text = strDocNo(1)
                Me.DtpCurDate.Value = ConvertToDate(strDocNo(2))
                Me.LblCustName.Text = strDocNo(3)
                Call FillGrid()
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub CmdGrpDispatchSlip_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles CmdGrpDispatchSlip.ButtonClick
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Perform the operation according to the button clicked
        ' Created On    :10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo errHandler
        Dim StrSql As String
        Dim Btn As MsgBoxResult
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                Call EnableControls(True, Me, True)
                Me.DtpCurDate.Focus()
              
                Me.ToolTip1.SetToolTip(Me.FsDispatchSlip, "Double Click the first column to delete current row")
                Call EnableDocNoCtrls(False)
                Me.TxtDocNo.Enabled = True
                Me.cmdDocNoHlp.Enabled = True
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                If Not ValidateBeforeSave() Then
                    With mP_Connection
                        .BeginTrans()
                        StrSql = MakeInsertQry()
                        .Execute(StrSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        .CommitTrans()
                    End With
                    Call MakeTextFile()
                    Me.TxtDocNo.Text = mstrDocNo
                    MsgBox("Transaction completed successfully with Docoment No.: " & mstrDocNo, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Me.CmdGrpDispatchSlip.Revert()
                    Call EnableControls(False, Me, False)
                    Call EnableDocNoCtrls(True)
                    With Me.FsDispatchSlip
                        .Enabled = True
                        .BlockMode = True
                        .Row = 1
                        .Row2 = .MaxRows
                        .Col = 1
                        .Col2 = .MaxCols
                        .Lock = True
                        .BlockMode = False
                        Me.ToolTip1.SetToolTip(Me.FsDispatchSlip, "")
                    End With
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                Btn = MsgBox("Cancel will undo all changes, Proceed ?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, ResolveResString(100))
                If Btn = MsgBoxResult.Yes Then
                    Me.DtpCurDate.Value = GetServerDate()
                    Me.TxtCustCode.Text = ""
                    With Me.FsDispatchSlip
                        .MaxRows = 0
                        Me.ToolTip1.SetToolTip(Me.FsDispatchSlip, "")
                    End With
                    Me.CmdGrpDispatchSlip.Revert()
                    Call EnableControls(False, Me, True)
                    Call EnableDocNoCtrls(True)
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                If Me.CmdGrpDispatchSlip.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    Btn = MsgBox("Do you want to save the changes ?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, ResolveResString(100))
                    If Btn = MsgBoxResult.Yes Then
                        CmdGrpDispatchSlip_ButtonClick(CmdGrpDispatchSlip, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE))
                    End If
                End If
                Me.Close()
        End Select
        Exit Sub
errHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0044_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '-----------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Set Property CheckFormName as per Rule book
        ' Created On    :10-Nov-2005
        '---------------------------------------------------------------------------------------------
        On Error GoTo errHandler
        mdifrmMain.CheckFormName = mintIndex
        Exit Sub 'To prevent the execution of errhandler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0044_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '---------------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Set Property NODEFONTBOLD as per Rule book
        ' Created On    :10-Nov-2005
        '---------------------------------------------------------------------------------------------------
        On Error GoTo errHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub 'To prevent the execution of errhandler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0044_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '------------------------------------------------------------------------------------------------------'
        ' Author        : Davinder singh
        ' Arguments     : Keycode and shift
        ' Return Value  : Nil
        ' Function      : To invoke the onlinehelp associated with form
        ' Created On    :10-Nov-2005
        '------------------------------------------------------------------------------------------------------'
        On Error GoTo errHandler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlFormHeader_Click(ctlFormHeader, New System.EventArgs())
        End If
        Exit Sub 'To prevent the execution of errhandler
errHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0044_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '---------------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Keycode of key pressed
        ' Return Value  : Nil
        ' Function      : Perform function as CLOSE button does
        ' Created On    :10-Nov-2005
        '-----------------------------------------------------------------------------------------------
        On Error GoTo errHandler ' Error Handler
        If KeyCode = System.Windows.Forms.Keys.Escape And Shift = 0 Then CmdGrpDispatchSlip_ButtonClick(CmdGrpDispatchSlip, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE))
        Exit Sub 'To prevent the execution of errhandler
errHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0044_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '---------------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Set Property AddFormNameToWindowList &
        ' Function FitToClient as per Rule book
        ' Intialise the Controls on form load
        ' Created On    :10-Nov-2005
        '-----------------------------------------------------------------------------------------------
        On Error GoTo errHandler ' Error Handler
        mintIndex = mdifrmMain.AddFormNameToWindowList(Me.Tag)
        FitToClient(Me, FraMain, ctlFormHeader, CmdGrpDispatchSlip, 500)
        Call InitializeCtrls()
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0044_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '-------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Cancel as Integer
        ' Return Value  : NIL
        ' Function      : RemoveFormNameFromWindowList
        ' Release form Object Memory from Database.
        ' Created On    :10-Nov-2005
        '--------------------------------------------------------------------------------------------------
        On Error GoTo errHandler
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Me.Dispose()
        Exit Sub 'To prevent the execution of errhandler
errHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub InitializeCtrls()
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To initialize the controls on the form
        ' Created On    :10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo errHandler
        Me.CmdCustCodeHlp.Image = My.Resources.ico111.ToBitmap
        Me.cmdDocNoHlp.Image = My.Resources.ico111.ToBitmap
        Me.TxtCustCode.MaxLength = 8
        Me.TxtDocNo.MaxLength = 10
        Me.DtpCurDate.Format = DateTimePickerFormat.Custom
        Me.DtpCurDate.CustomFormat = gstrDateFormat
        Me.DtpCurDate.Value = GetServerDate()
        With Me.FsDispatchSlip
            .MaxRows = 0
            .MaxCols = 7
            .set_RowHeight(0, 300)
            .Row = 0
            .set_ColWidth(0, 300)
            .set_ColWidth(EnumGrid.Item_Code, 1500)
            .set_ColWidth(EnumGrid.Item_CmdHlp, 300)
            .set_ColWidth(EnumGrid.Item_Description, 2000)
            .set_ColWidth(EnumGrid.Cust_DrgNo, 1500)
            .set_ColWidth(EnumGrid.Cust_DrgnoHlp, 300)
            .set_ColWidth(EnumGrid.Cust_Drgno_Description, 2000)
            .set_ColWidth(EnumGrid.item_qty, 1000)
            Call .SetText(EnumGrid.Item_Code, 0, "Item Code")
            Call .SetText(EnumGrid.Item_CmdHlp, 0, " ")
            Call .SetText(EnumGrid.Item_Description, 0, "Description")
            Call .SetText(EnumGrid.Cust_DrgNo, 0, "Cust Part Code")
            Call .SetText(EnumGrid.Cust_DrgnoHlp, 0, " ")
            Call .SetText(EnumGrid.Cust_Drgno_Description, 0, "Description")
            Call .SetText(EnumGrid.item_qty, 0, "Qty.")
            Call CmdGrpDispatchSlip.ShowButtons(True, False, False, False)
            Call CmdGrpDispatchSlip.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) / 3.5), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) / 1.23), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
            Call EnableControls(False, Me, True)
            Call EnableDocNoCtrls(True)
        End With
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ADDRow()
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Add the new row in the grid
        ' Created On    :10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo errHandler
        With Me.FsDispatchSlip
            .MaxRows = .MaxRows + 1
            .set_RowHeight(.MaxRows, 300)
            .Row = .MaxRows
            .Col = EnumGrid.Item_Code
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = EnumGrid.Item_CmdHlp
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
            .TypeButtonPicture = My.Resources.resEmpower.ico111.ToBitmap
            .Col = EnumGrid.Item_Description
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = EnumGrid.Cust_DrgNo
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = EnumGrid.Cust_DrgnoHlp
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
            .TypeButtonPicture = My.Resources.resEmpower.ico111.ToBitmap
            .Col = EnumGrid.Cust_Drgno_Description
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = EnumGrid.item_qty
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Row = .MaxRows
            .Col = EnumGrid.Item_Code
            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
            If Me.CmdGrpDispatchSlip.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                .set_ColWidth(EnumGrid.Item_CmdHlp, 0)
                .set_ColWidth(EnumGrid.Cust_DrgnoHlp, 0)
            Else
                .set_ColWidth(EnumGrid.Item_CmdHlp, 300)
                .set_ColWidth(EnumGrid.Cust_DrgnoHlp, 300)
            End If
            .Focus()
        End With
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtCustCode.TextChanged
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To clear the controls on changing the customer code
        ' Created On    :10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo errHandler
        Me.LblCustName.Text = ""
        Me.FsDispatchSlip.MaxRows = 0
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TxtCustCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To call the customer code help on F1 key press
        ' Created On    :10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo errHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then CmdCustCodeHlp.PerformClick()
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtCustCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To add the row in the grid on hiting the enter key
        ' Created On    : 10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo errHandler
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            Call txtCustCode_Validating(TxtCustCode, New System.ComponentModel.CancelEventArgs(False))
            If mblnStatus Then
                mblnStatus = False
            Else
                Call ADDRow()
            End If
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Function MakeInsertQry() As String
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To make the Query to save data into tables
        ' Created On    : 10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo errHandler
        Dim StrSql As String
        Dim intLoopCtr As Short
        mstrDocNo = GenerateDocNO(105, Me.DtpCurDate.Text, eMPowerFunctions.DocTypeEnum.Doc_ECN, False, True)
        StrSql = "insert into dispatchslip_hdr values('" & mstrDocNo & "','" & Trim(Me.TxtCustCode.Text) & "','" & getDateForDB(Me.DtpCurDate.Value) & "',getdate(),'" & mP_User & "',getdate(),'" & mP_User & "','" & gstrUNITID & "')"
        With Me.FsDispatchSlip
            For intLoopCtr = 1 To .MaxRows
                .Row = intLoopCtr
                StrSql = StrSql & vbCrLf
                StrSql = StrSql & "insert into dispatchslip_dtl values('" & mstrDocNo & "','"
                .Col = EnumGrid.Item_Code
                StrSql = StrSql & Trim(.Text) & "','"
                .Col = EnumGrid.Cust_DrgNo
                StrSql = StrSql & Trim(.Text) & "',"
                .Col = EnumGrid.item_qty
                StrSql = StrSql & Trim(.Text)
                StrSql = StrSql & ",getDate(),'" & mP_User & "',getDate(),'" & mP_User & "','" & gstrUNITID & "')"
            Next
        End With
        MakeInsertQry = StrSql
        Exit Function 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub MakeTextFile()
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To make the text file with name as the document No. generated
        ' Created On    : 10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo errHandler
        Dim FSO As Scripting.FileSystemObject
        Dim TxtStream As Scripting.TextStream
        Dim intLoopCtr As Short
        Dim StrLine As String
        FSO = New Scripting.FileSystemObject
        If Not FSO.FolderExists(gstrLocalCDrive & "\Bar Code Files") Then
            FSO.CreateFolder(gstrLocalCDrive & "\Bar Code Files")
        End If
        If Not FSO.FolderExists(gstrLocalCDrive & "\Bar Code Files\Dispatch Slip") Then
            FSO.CreateFolder(gstrLocalCDrive & "\Bar Code Files\Dispatch Slip")
        End If
        TxtStream = FSO.CreateTextFile(gstrLocalCDrive & "\Bar Code Files\Dispatch Slip\" & mstrDocNo & ".txt")
        With Me.FsDispatchSlip
            For intLoopCtr = 1 To .MaxRows
                .Row = intLoopCtr
                StrLine = ""
                StrLine = mstrDocNo & "*"
                .Col = EnumGrid.Cust_DrgNo
                StrLine = StrLine & Trim(.Text) & "*"
                .Col = EnumGrid.item_qty
                StrLine = StrLine & Trim(.Text)
                TxtStream.WriteLine((StrLine))
            Next
        End With
        TxtStream.Close()
        FSO = Nothing
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function ValidateBeforeSave() As Boolean
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To validate the data before saving
        ' Created On    : 10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo errHandler
        Dim intLoopCtr As Short
        ValidateBeforeSave = False
        Call DeleteBlankRows()
        With Me.FsDispatchSlip
            If .MaxRows = 0 Then
                MsgBox("No data is available to save", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                ValidateBeforeSave = True
                Exit Function
            End If
            For intLoopCtr = 1 To .MaxRows
                'item code
                .Col = EnumGrid.Item_Code
                If Trim(.Text) = "" Then
                    MsgBox("Item code can't be blank", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    mblnStatus = True
                End If
                If mblnStatus Then
                    ValidateBeforeSave = True
                    mblnStatus = False
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    .Focus()
                    Exit Function
                End If
                'Cust Drgno
                .Col = EnumGrid.Cust_DrgNo
                If Trim(.Text) = "" Then
                    MsgBox("Cust Drgno can't be blank", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    mblnStatus = True
                End If
                If mblnStatus Then
                    ValidateBeforeSave = True
                    mblnStatus = False
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    .Focus()
                    Exit Function
                End If
                'Qty.
                .Col = EnumGrid.item_qty
                If Trim(.Text) = "" Or Trim(.Text) = "0" Then
                    MsgBox("Quantity Can't be Zero/Blank", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    ValidateBeforeSave = True
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    .Focus()
                    Exit Function
                End If
            Next
        End With
        Exit Function 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function DeleteBlankRows() As Object
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Check and delete the blank rows before saving
        ' Created On    : 10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo errHandler
        Dim StrA As Object
        Dim StrB As Object
        Dim intLoopCtr As Short
        Dim counter As Short
        With Me.FsDispatchSlip
            counter = .MaxRows
            For intLoopCtr = 1 To counter
                .Row = intLoopCtr
                StrA = Nothing
                Call .GetText(EnumGrid.Item_Code, intLoopCtr, StrA)
                StrB = Nothing
                Call .GetText(EnumGrid.Cust_DrgNo, intLoopCtr, StrB)
                If Trim(StrA) = "" And Trim(StrB) = "" Then
                    .Action = FPSpreadADO.ActionConstants.ActionDeleteRow
                    .MaxRows = .MaxRows - 1
                    intLoopCtr = intLoopCtr - 1
                End If
                If intLoopCtr >= .MaxRows Then Exit For
            Next
        End With
        Exit Function 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub txtCustCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles TxtCustCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To validate the entered customer code
        ' Created On    : 10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo errHandler
        Dim StrSql As String
        Dim rs As ClsResultSetDB
        If Trim(Me.TxtCustCode.Text) <> "" Then
            rs = New ClsResultSetDB
            StrSql = "select customer_code,cust_name from customer_mst where customer_code='" & Trim(Me.TxtCustCode.Text) & "' and Unit_code = '" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
            Call rs.GetResult(StrSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rs.GetNoRows <= 0 Then
                MsgBox("Entered customer code is Invalid", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                mblnStatus = True
                Me.TxtCustCode.Focus()
                GoTo EventExitSub
            Else
                Me.LblCustName.Text = rs.GetValue("cust_name")
            End If
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Function EnableDocNoCtrls(ByRef Bln As Boolean) As Object
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To validate the entered customer code
        ' Created On    : 10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo errHandler
        Me.TxtDocNo.Enabled = Bln
        Me.cmdDocNoHlp.Enabled = Bln
        If Bln Then
            TxtDocNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        Else
            TxtDocNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        End If
        Exit Function 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function FillGrid() As Object
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To Fill the grid in View mode
        ' Created On    : 10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo errHandler
        Dim rsGrid As ClsResultSetDB
        Dim intLoopCtr As Short
        Dim intCtr As Short
        rsGrid = New ClsResultSetDB
        Call rsGrid.GetResult("select A.doc_no,A.item_code,A.cust_drgno,A.dispatch_qty,B.description,C.drg_desc from dispatchslip_dtl A,item_mst B,custitem_mst C where A.item_code=B.item_code and A.cust_drgno=C.cust_drgno and A.Unit_Code=B.Unit_Code and A.Unit_Code=C.Unit_Code and A.doc_no='" & Me.TxtDocNo.Text & "' and A.Unit_code = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        intCtr = rsGrid.GetNoRows
        If intCtr > 0 Then
            rsGrid.MoveFirst()
            With Me.FsDispatchSlip
                .Enabled = True
                .MaxRows = 0
                For intLoopCtr = 1 To intCtr
                    Call ADDRow()
                    Call .SetText(EnumGrid.Item_Code, intCtr, rsGrid.GetValue("item_code"))
                    Call .SetText(EnumGrid.Item_Description, intCtr, rsGrid.GetValue("description"))
                    Call .SetText(EnumGrid.Cust_DrgNo, intCtr, rsGrid.GetValue("cust_drgno"))
                    Call .SetText(EnumGrid.Cust_Drgno_Description, intCtr, rsGrid.GetValue("drg_desc"))
                    Call .SetText(EnumGrid.item_qty, intCtr, rsGrid.GetValue("dispatch_qty"))
                    rsGrid.MoveNext()
                Next
                .BlockMode = True
                .Row = 1
                .Row2 = .MaxRows
                .Col = 1
                .Col2 = .MaxCols
                .Lock = True
                .BlockMode = False
            End With
        Else
            MsgBox("No data is available to display", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
        End If
        rsGrid.ResultSetClose()
        Exit Function 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub txtDocNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtDocNo.TextChanged
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To clear the controls when textvbox is cleared
        ' Created On    : 10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo errHandler
        If Trim(Me.TxtDocNo.Text) = "" Then
            Call EnableControls(False, Me, True)
            Call EnableDocNoCtrls(True)
            Me.DtpCurDate.Value = GetServerDate()
            Me.TxtDocNo.Focus()
        End If
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtDocNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TxtDocNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To call the Doc. No. help on F1 key press
        ' Created On    :10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo errHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then Me.cmdDocNoHlp.PerformClick()
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtDocNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtDocNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To call the validation of Doc No on hitting the enter key
        ' Created On    : 10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo errHandler
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            Call txtDocNo_Validating(TxtDocNo, New System.ComponentModel.CancelEventArgs(False))
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtDocNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles TxtDocNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To validate the entered Doc No.
        ' Created On    : 10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo errHandler
        Dim StrSql As String
        Dim rs As ClsResultSetDB
        If Trim(Me.TxtDocNo.Text) <> "" Then
            rs = New ClsResultSetDB
            StrSql = "select customer_code,dispatch_dt from dispatchslip_hdr where Unit_Code ='" & gstrUNITID & "' and doc_no=" & Me.TxtDocNo.Text
            Call rs.GetResult(StrSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rs.GetNoRows > 0 Then
                Me.TxtCustCode.Text = rs.GetValue("customer_code")
                Me.DtpCurDate.Value = rs.GetValue("dispatch_dt")
                Call FillGrid()
            Else
                MsgBox("Entered Document No. is Invalid", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                Me.TxtDocNo.SelectionStart = 0
                Me.TxtDocNo.SelectionLength = Len(Me.TxtDocNo.Text)
                Me.TxtDocNo.Focus()
                GoTo EventExitSub
            End If
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub ctlFormHeader_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader.Click
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Display the Help
        ' Created On    :10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo errHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        Exit Sub 'This is to avoid the execution of the error handler
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub DtpCurDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DtpCurDate.KeyPress
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Set the focus to customer code text box
        ' Created On    :10-Nov-2005
        '--------------------------------------------------------------------
        Dim KeyAscii As Short = Asc(e.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = Keys.Enter Then Me.TxtCustCode.Focus()
        Exit Sub    'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub DtpCurDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtpCurDate.ValueChanged
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Reset the controls on changing the date
        ' Created On    :10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo ErrHandler
        'Call EnableControls(True, Me, True)
        Me.TxtDocNo.Enabled = True
        Me.FsDispatchSlip.MaxRows = 0
        Call EnableDocNoCtrls(False)
        Me.TxtDocNo.Enabled = True
        TxtDocNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        Me.cmdDocNoHlp.Enabled = True
        Exit Sub    'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub FsDispatchSlip_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles FsDispatchSlip.ButtonClicked
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To show the Item code and Customer Drgno. helps according to button clicked
        ' Created On    : 10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim strItemCode() As String
        Dim intLoopCtr As Integer
        Dim strItems As String
        If e.row <> 0 Then
            Select Case e.col
                Case EnumGrid.Item_CmdHlp
                    With Me.FsDispatchSlip
                        If .MaxRows > 1 Then
                            strItems = "("
                            For intLoopCtr = 1 To .MaxRows
                                .Row = intLoopCtr
                                .Col = EnumGrid.Item_Code
                                strItems = strItems & "'" & Trim(.Text) & "',"
                            Next
                            strItems = Mid(strItems, 1, Len(strItems) - 1)
                            strItems = strItems & ")"
                        Else
                            strItems = "('')"
                        End If
                    End With
                    strsql = "set dateformat 'dmy'" & vbCrLf
                    strsql = strsql & "select distinct A.Item_code,B.Description from "
                    strsql = strsql & "(select item_code from DailyMktSchedule "
                    strsql = strsql & "Where Unit_code = '" & gstrUNITID & "' and trans_date Between '01/' +right(('0'+cast(month('" & Format(Me.DtpCurDate.Value, "dd/MMM/yyyy") & "') as varchar(2))),2)+'/'+ cast(year('" & Format(Me.DtpCurDate.Value, "dd/MMM/yyyy") & "')as varchar(4)) and '" & Format(Me.DtpCurDate.Value, "dd/MMM/yyyy") & "' and account_code='" & Trim(Me.TxtCustCode.Text) & "' and status=1"
                    strsql = strsql & vbCrLf & "group by item_code" & vbCrLf
                    strsql = strsql & "Having (Sum(schedule_quantity) - Sum(despatch_qty)) > 0 "
                    strsql = strsql & vbCrLf & "Union" & vbCrLf
                    strsql = strsql & "select item_code from  MonthlyMktSchedule where Unit_code = '" & gstrUNITID & "' and year_month='" & Year(Me.DtpCurDate.Value) & Format(Month(Me.DtpCurDate.Value), "0#") & "'  and account_code='" & Trim(Me.TxtCustCode.Text) & "' and status=1" & vbCrLf
                    strsql = strsql & "group by item_code" & vbCrLf
                    strsql = strsql & "Having (Sum(schedule_qty) - Sum(despatch_qty)) > 0) A,item_mst B"
                    strsql = strsql & " Where A.item_code=B.item_code and B.status='A' and B.Unit_Code = '" & gstrUNITID & "' and A.item_code not in " & strItems
                    strItemCode = Me.ctlEMPHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strsql, "Item Code Help")
                    If Not (UBound(strItemCode) <= 0) Then
                        If (Len(strItemCode(0)) >= 1) And strItemCode(0) = "0" Then
                            MsgBox("No Record Available To Display", vbInformation + vbOKOnly, ResolveResString(100))
                            If Me.FsDispatchSlip.MaxRows = 1 Then
                                Me.FsDispatchSlip.MaxRows = 0
                                Me.DtpCurDate.Focus()
                            Else
                                Me.FsDispatchSlip.MaxRows = Me.FsDispatchSlip.MaxRows - 1
                            End If
                        Else
                            With Me.FsDispatchSlip
                                Call .SetText(EnumGrid.Item_Code, e.row, strItemCode(0))
                                Call .SetText(EnumGrid.Item_Description, e.row, strItemCode(1))
                                .Row = .Row
                                .Col = EnumGrid.Cust_DrgNo
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            End With
                        End If
                    End If
                Case EnumGrid.Cust_DrgnoHlp
                    With Me.FsDispatchSlip
                        .Row = e.row
                        .Col = EnumGrid.Item_Code
                        If Trim(.Text) <> "" Then
                            strsql = "select distinct cust_drgno,drg_desc from custitem_mst where item_code='" & Trim(.Text) & "' and Unit_code = '" & gstrUNITID & "'"
                            strItemCode = Me.ctlEMPHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strsql, "Item Code Help")
                            If Not (UBound(strItemCode) <= 0) Then
                                If (Len(strItemCode(0)) >= 1) And strItemCode(0) = "0" Then
                                    MsgBox("No. Record Available To Display", vbInformation + vbOKOnly, ResolveResString(100))
                                Else
                                    Call .SetText(EnumGrid.Cust_DrgNo, e.row, strItemCode(0))
                                    Call .SetText(EnumGrid.Cust_Drgno_Description, e.row, strItemCode(1))
                                    .Row = e.row
                                    .Col = EnumGrid.item_qty
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                End If
                            End If
                        Else
                            MsgBox("Please First Select the Item")
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            Exit Sub
                        End If
                    End With
            End Select
        End If
        Exit Sub    'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub FsDispatchSlip_DblClick(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_DblClickEvent) Handles FsDispatchSlip.DblClick
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Delete the active row if mode is not VIEW
        ' Created On    : 10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo ErrHandler
        If Me.CmdGrpDispatchSlip.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If e.col = 0 And e.row <> 0 Then
                With Me.FsDispatchSlip
                    .Row = e.row
                    .Action = FPSpreadADO.ActionConstants.ActionDeleteRow
                    .MaxRows = .MaxRows - 1
                End With
            End If
        End If
        Exit Sub    'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub FsDispatchSlip_EditChange(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles FsDispatchSlip.EditChange
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Clear the next cells if the value of previous cell changed
        ' Created On    : 10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo ErrHandler
        With Me.FsDispatchSlip
            Select Case e.col
                Case EnumGrid.Item_Code
                    Call .SetText(EnumGrid.Item_Description, e.row, "")
                    Call .SetText(EnumGrid.Cust_DrgNo, e.row, "")
                    Call .SetText(EnumGrid.Cust_Drgno_Description, e.row, "")
                    Call .SetText(EnumGrid.item_qty, e.row, "")
                Case EnumGrid.Cust_DrgNo
                    Call .SetText(EnumGrid.Cust_Drgno_Description, e.row, "")
                    'Call .SetText(EnumGrid.Item_Qty, Row, "")
                Case EnumGrid.item_qty
                    .Row = e.row
                    .Col = EnumGrid.Cust_DrgNo
                    If e.col = EnumGrid.item_qty Then
                        If Trim(.Text) = "" Then
                            MsgBox("Customer Drgno. can not be left blank", vbInformation + vbOKOnly, ResolveResString(100))
                            .Col = EnumGrid.Cust_DrgNo
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            Exit Sub
                        End If
                    End If
            End Select
        End With
        Exit Sub    'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub FsDispatchSlip_KeyDownEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles FsDispatchSlip.KeyDownEvent
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To call the item code and Customer part code helps
        ' Created On    :10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim e1 As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent
        If e.keyCode = Keys.F1 And e.shift = 0 Then
            With Me.FsDispatchSlip
                If .ActiveRow > 0 Then
                    Select Case .ActiveCol
                        Case EnumGrid.Item_Code
                            Call FsDispatchSlip_ButtonClicked(FsDispatchSlip, New AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent(FsDispatchSlip.ActiveCol + 1, FsDispatchSlip.ActiveRow + 1, 0))
                        Case EnumGrid.Cust_DrgNo
                            Call FsDispatchSlip_ButtonClicked(FsDispatchSlip, New AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent(FsDispatchSlip.ActiveCol + 1, FsDispatchSlip.ActiveRow, 0))
                    End Select
                End If
            End With
        End If
        Exit Sub    'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub FsDispatchSlip_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles FsDispatchSlip.KeyPressEvent
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Call the validations and set the focus to next controls on hitting the enter key
        ' Created On    : 10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim E1 As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent
        mblnStatus = False
        With Me.FsDispatchSlip
            If e.keyAscii = Keys.Enter Then
                Select Case .ActiveCol
                    Case EnumGrid.Item_Code
                        Call FsDispatchSlip_LeaveCell(FsDispatchSlip, New AxFPSpreadADO._DSpreadEvents_LeaveCellEvent(FsDispatchSlip.ActiveCol + 1, FsDispatchSlip.ActiveRow + 1, FsDispatchSlip.ActiveCol + 1, FsDispatchSlip.ActiveRow + 1, False))
                        If Not mblnStatus Then
                            .Row = .ActiveRow
                            .Col = EnumGrid.Cust_DrgNo
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        Else
                            mblnStatus = False
                        End If
                    Case EnumGrid.Cust_DrgNo
                        Call FsDispatchSlip_LeaveCell(FsDispatchSlip, New AxFPSpreadADO._DSpreadEvents_LeaveCellEvent(FsDispatchSlip.ActiveCol + 1, FsDispatchSlip.ActiveRow + 1, FsDispatchSlip.ActiveCol + 1, FsDispatchSlip.ActiveRow + 1, False))
                        If Not mblnStatus Then
                            .Row = .ActiveRow
                            .Col = EnumGrid.item_qty
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        Else
                            mblnStatus = False
                        End If
                    Case EnumGrid.item_qty
                        .Row = .ActiveRow
                        .Col = EnumGrid.item_qty
                        If .Text = "" Or .Text = "0" Then
                            MsgBox("Quantity Can't be Zero/Blank ", vbInformation + vbOKOnly, ResolveResString(100))
                            .Text = 0
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                        Else
                            If .ActiveRow = .MaxRows Then
                                Call ADDRow()
                                Me.FsDispatchSlip.Focus()
                            Else
                                Me.CmdGrpDispatchSlip.Focus()
                            End If
                        End If
                End Select
            End If
        End With
        Exit Sub    'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub FsDispatchSlip_LeaveCell(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles FsDispatchSlip.LeaveCell
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Validate the entered value in the cell
        ' Created On    :10-Nov-2005
        '--------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim rsData As ClsResultSetDB
        With Me.FsDispatchSlip
            If Not e.newCol <= 0 Then
                rsData = New ClsResultSetDB
                Select Case e.col
                    Case EnumGrid.Item_Code
                        .Row = e.row
                        .Col = EnumGrid.Item_Code
                        If Trim(.Text) <> "" Then
                            If Not (e.newCol = EnumGrid.Item_CmdHlp And e.newRow = e.row) Then
                                strsql = ""
                                strsql = strsql & "select A.Item_code,B.Description from "
                                strsql = strsql & "(select item_code from DailyMktSchedule "
                                strsql = strsql & "Where Unit_Code = '" & gstrUNITID & "' and trans_date Between '01/' +right(('0'+cast(month('" & Format(Me.DtpCurDate.Value, "dd/MMM/yyyy") & "') as varchar(2))),2)+'/'+ cast(year('" & Format(Me.DtpCurDate.Value, "dd/MMM/yyyy") & "')as varchar(4)) and '" & Format(Me.DtpCurDate.Value, "dd/MMM/yyyy") & "' and account_code='" & Trim(Me.TxtCustCode.Text) & "' and status=1 "
                                strsql = strsql & vbCrLf & "group by item_code" & vbCrLf
                                strsql = strsql & "Having (Sum(schedule_quantity) - Sum(despatch_qty)) > 0 "
                                strsql = strsql & vbCrLf & "Union" & vbCrLf
                                strsql = strsql & "select item_code from  MonthlyMktSchedule where Unit_Code = '" & gstrUNITID & "' and year_month='" & Year(Me.DtpCurDate.Value) & Format(Month(Me.DtpCurDate.Value), "0#") & "' and account_code='" & Trim(Me.TxtCustCode.Text) & "' and status=1 " & vbCrLf
                                strsql = strsql & "group by item_code" & vbCrLf
                                strsql = strsql & "Having (Sum(schedule_qty) - Sum(despatch_qty)) > 0) A,item_mst B"
                                strsql = strsql & " Where A.item_code=B.item_code and B.status='A' and B.Unit_Code = '" & gstrUNITID & "' and A.item_code in ('" & Trim(.Text) & "')"
                                Call rsData.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                If rsData.GetNoRows <= 0 Then
                                    MsgBox("Entered Item Code is not valid", vbInformation + vbOKOnly, ResolveResString(100))
                                    rsData.ResultSetClose()
                                    rsData = Nothing
                                    mblnStatus = True
                                    .Focus()
                                    e.cancel = True
                                    Exit Sub
                                Else
                                    Call .SetText(EnumGrid.Item_Description, e.row, rsData.GetValue("description"))
                                End If
                            End If
                        End If
                    Case EnumGrid.Cust_DrgNo
                        .Row = e.row
                        .Col = EnumGrid.Cust_DrgNo
                        If Trim(.Text) <> "" Then
                            If Not (e.newCol = EnumGrid.Cust_DrgnoHlp And e.newRow = e.row) Then
                                strsql = "select cust_drgno,drg_desc from custitem_mst where Unit_Code = '" & gstrUNITID & "' and cust_drgno='" & Trim(.Text) & "' and item_code='"
                                .Col = EnumGrid.Item_Code
                                strsql = strsql & Trim(.Text) & "'"
                                .Col = EnumGrid.Cust_DrgNo
                                Call rsData.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                If rsData.GetNoRows <= 0 Then
                                    MsgBox("Entered Customer Drgno is not valid", vbInformation + vbOKOnly, ResolveResString(100))
                                    rsData.ResultSetClose()
                                    rsData = Nothing
                                    mblnStatus = True
                                    .Focus()
                                    e.cancel = True
                                    Exit Sub
                                Else
                                    Call .SetText(EnumGrid.Cust_Drgno_Description, e.row, rsData.GetValue("drg_desc"))
                                End If
                            End If
                        End If
                End Select
            End If
        End With
        Exit Sub    'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
End Class