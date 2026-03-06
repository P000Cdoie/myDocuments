Option Strict Off
Option Explicit On
Friend Class frmMKTMST0015
	Inherits System.Windows.Forms.Form
	'(C) 2001 MIND, All rights reserved
	'Form Name      - frmMKTMST0015.frm
	'Description    - Updations in System Parameters
	'                   -This form will update the values of System Parameter table
	'Created by     - Shubhra Verma
	'Creation Date  - 05/06/2006 (dd/mm/yyyy Format)
	'Modified Date  -   21/06/2006 (dd/mm/yyyy Format)
	'
	'Rev. History   -
	'Reason and
	'Logic for      -
    '-------------------------------------------------------------------------------
    'Revised by     :   Nitin Mehta
    'Revision date  :   21/04/2011
    'Reason         :   Multi Unit changes
    '-------------------------------------------------------------------------------
    Dim mintIndex As Short
    Dim RsSalesParam As New ADODB.Recordset
    Dim rsSalesParam2 As New ADODB.Recordset
    Dim mStrSQL As Object
    Dim mstrsql2 As String
    Dim mintCount As Short
    Dim mintRow As Short
    Dim mPos As Short
    Dim mValue As Object
    Dim mVbStr As String
    Dim mCharChk As String
    Dim strString As String
    Public Enum enumSysParamGrid
        Col_name = 1
        Description = 2
        Values = 3
    End Enum

    Private Sub cmd0015Edit_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtnEditGrp.ButtonClickEventArgs) Handles cmd0015Edit.ButtonClick
        On Error GoTo Errorhandler
        With spdSalesParam
            Select Case e.Button
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                    'it allows user to edit data of column "value" in the grid
                    With spdSalesParam
                        .Col = enumSysParamGrid.Values
                        .Col2 = enumSysParamGrid.Values
                        .Row = 1
                        .Row2 = .MaxRows
                        .BlockMode = True
                        .Lock = False
                        .BlockMode = False
                    End With
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                    Call frmMKTMST0015_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                    If Me.cmd0015Edit.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or Me.cmd0015Edit.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                        mVbStr = CStr(MsgBox("Do you want to save the changes", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "eMPro"))
                        If mVbStr = CStr(MsgBoxResult.Yes) Then
                            Call cmd0015Edit_ButtonClick(cmd0015Edit, New UCActXCtl.UCbtnEditGrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE))
                        End If
                    End If
                    Me.Close()
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                    'It saves data in sales_parameter table
                    mStrSQL = "select * from sales_parameter where unit_code='" & gstrUNITID & "'"
                    RsSalesParam.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                    RsSalesParam.Open(mStrSQL, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                    With Me.spdSalesParam
                        For mintRow = 1 To .MaxRows
                            mValue = Nothing
                            Call .GetText(enumSysParamGrid.Col_name, mintRow, mValue)
                            For mintCount = 0 To .MaxRows - 1
                                If Trim(UCase(RsSalesParam.Fields(mintCount).Name)) = Trim(UCase(mValue)) Then
                                    .Col = enumSysParamGrid.Values
                                    mValue = Nothing
                                    Call .GetText(enumSysParamGrid.Values, mintRow, mValue)
                                    If RsSalesParam.Fields(mintCount).Type = ADODB.DataTypeEnum.adChar Or ADODB.DataTypeEnum.adVarChar Then
                                        strString = mValue
                                        mValue = QuoteString(strString)
                                        mValue = strString
                                        'QuoteString(mValue)
                                    End If
                                    .Col = enumSysParamGrid.Values
                                    mstrsql2 = "Update sales_parameter set " & RsSalesParam.Fields(mintCount).Name & "  = '" & mValue & "' where unit_code='" & gstrUNITID & "' "
                                    rsSalesParam2.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                                    rsSalesParam2.Open(mstrsql2, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
                                    GoTo Label1
                                End If
                            Next
Label1:
                        Next
                    End With
                    RsSalesParam.Close()
                    MsgBox("Record Updated Successfully", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Me.cmd0015Edit.Revert()
                    .Col = enumSysParamGrid.Values
                    .Col2 = enumSysParamGrid.Values
                    .Row = 1
                    .Row2 = .MaxRows
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
            End Select
        End With
        Exit Sub
        rsSalesParam2.Close()
Errorhandler:
        ''Debug.Print("")
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0015_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        ' On pressing F4 , help gets dispayed
        If (KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0) Then Call ctlMkt0015Hdr_Click(ctlMkt0015Hdr, New System.EventArgs())
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0015_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                'If user press the ESC Key ,the Form will be unloaded
                If (Me.cmd0015Edit.Mode) <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call Me.cmd0015Edit.Revert()
                        gblnCancelUnload = False : gblnFormAddEdit = False
                        spdSalesParam.Refresh()
                        Call POPULATEGRID()
                        spdSalesParam.Col = enumSysParamGrid.Values
                        spdSalesParam.Col2 = enumSysParamGrid.Values
                        spdSalesParam.Row = 1
                        spdSalesParam.Row2 = spdSalesParam.MaxRows
                        spdSalesParam.BlockMode = True
                        spdSalesParam.Lock = True
                        spdSalesParam.BlockMode = False
                        With Me
                        End With
                        cmd0015Edit.Focus()
                        GoTo EventExitSub
                    Else
                        Me.ActiveControl.Focus()
                    End If
                End If
            Case System.Windows.Forms.Keys.Return
                'Action on Return Key
                With Me.spdSalesParam
                    If .ActiveRow <> .MaxRows Then
                        If .ActiveCol = .MaxCols Then
                            .Col = 1
                            .Row = .ActiveRow + 1
                        Else
                            .Col = .ActiveCol + 1
                        End If
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    Else
                        .Row = 1
                        .Col = 1
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    End If
                End With
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
    Private Sub frmMKTMST0015_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex ''Holds the Form Name
        frmModules.NodeFontBold(Me.Tag) = True
        Exit Sub
ErrHandler:  ''Error Handler runs in case of Unhandled error
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
    End Sub
    Private Sub frmMKTMST0015_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0015_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        'Load the caption
        Call FillLabelFromResFile(Me)
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.AppStarting)
        'Size the form to client workspace
        Call FitToClient(Me, (Me.Framain), (Me.ctlMkt0015Hdr), (Me.cmd0015Edit))
        'get the index of form in the window list
        mintIndex = mdifrmMain.AddFormNameToWindowList(Me.ctlMkt0015Hdr.Tag)
        'Call InitializeFormSettings 'Initial Form Settings
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call POPULATEGRID()
        Call InitializeFormSettings() 'Initial Form Settings
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0015_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        frmModules.NodeFontBold(Me.Tag) = False
        Me.Dispose()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub InitializeFormSettings()
        'Author-                Shubhra Verma
        'Arguments              NIL
        'Return Value           NIL
        'Function Comments      this function is used to initialize teh form settings
        'Creation Date          21/jun/2006
        On Error GoTo ErrHandler
        With spdSalesParam
            .Col = enumSysParamGrid.Col_name
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = enumSysParamGrid.Description
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = 1
            .Col2 = .MaxCols
            .Row = 1
            .Row2 = .MaxRows
            .BlockMode = True
            .Lock = True
            .BlockMode = False
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub

    Private Function POPULATEGRID() As Object
        'Author-                Shubhra Verma
        'Arguments              NIL
        'Return Value           NIL
        'Function Comments      this function is used to fill the values in a grid
        'Creation Date          21/jun/2006
        On Error GoTo ErrHandler
        mStrSQL = "select * from sales_parameter where unit_code='" & gstrUNITID & "'"
        RsSalesParam = New ADODB.Recordset
        RsSalesParam.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        RsSalesParam.Open(mStrSQL, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        If RsSalesParam.RecordCount >= 1 Then
            Me.spdSalesParam.MaxRows = 0
            RsSalesParam.MoveFirst()
            For mintCount = 0 To RsSalesParam.Fields.Count - 1
                If RsSalesParam.Fields(mintCount).Name.ToLower = "unit_code" Then
                    Continue For
                End If
                With Me.spdSalesParam
                    .MaxRows = .MaxRows + 1
                    'for the Value of column 'column name'
                    .Col = enumSysParamGrid.Col_name : .Row = mintCount + 1
                    Call .SetText(enumSysParamGrid.Col_name, mintCount + 1, RsSalesParam.Fields(mintCount).Name)
                    'for the Value of column 'Value'
                    Select Case RsSalesParam(mintCount).Type
                        Case ADODB.DataTypeEnum.adSmallInt, ADODB.DataTypeEnum.adBigInt, ADODB.DataTypeEnum.adDecimal, ADODB.DataTypeEnum.adDouble, ADODB.DataTypeEnum.adInteger, ADODB.DataTypeEnum.adTinyInt, ADODB.DataTypeEnum.adUnsignedBigInt, ADODB.DataTypeEnum.adUnsignedInt, ADODB.DataTypeEnum.adVarNumeric, ADODB.DataTypeEnum.adUnsignedSmallInt
                            .Col = enumSysParamGrid.Values : .Row = mintCount + 1
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                            .TypeMaxEditLen = RsSalesParam(mintCount).DefinedSize
                            If IsDBNull(RsSalesParam(mintCount).Value) Then
                                Call .SetText(enumSysParamGrid.Values, mintCount + 1, 0)
                            Else
                                Call .SetText(enumSysParamGrid.Values, mintCount + 1, RsSalesParam(mintCount).Value)
                            End If
                        Case ADODB.DataTypeEnum.adNumeric
                            .Col = enumSysParamGrid.Values : .Row = mintCount + 1
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                            .TypeMaxEditLen = RsSalesParam(mintCount).DefinedSize
                            .TypeFloatDecimalPlaces = 4
                            If Not IsDBNull(RsSalesParam(mintCount).Value) Then
                                Call .SetText(enumSysParamGrid.Values, mintCount + 1, Format(RsSalesParam(mintCount).Value, "0.0000"))
                            Else
                                Call .SetText(enumSysParamGrid.Values, mintCount + 1, Format(0, "0.0000"))
                            End If
                        Case ADODB.DataTypeEnum.adBoolean
                            .Col = enumSysParamGrid.Values : .Row = mintCount + 1
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                            If IsDBNull(RsSalesParam(mintCount).Value) Then
                                Call .SetText(enumSysParamGrid.Values, mintCount + 1, RsSalesParam.Fields(mintCount).Value)
                            ElseIf RsSalesParam(mintCount).Value Then
                                Call .SetText(enumSysParamGrid.Values, mintCount + 1, 1)
                            Else
                                Call .SetText(enumSysParamGrid.Values, mintCount + 1, RsSalesParam.Fields(mintCount).Value)
                            End If
                        Case ADODB.DataTypeEnum.adLongVarChar
                            .Col = enumSysParamGrid.Values : .Row = mintCount + 1
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                            .TypeMaxEditLen = 8000
                            Call .SetText(enumSysParamGrid.Values, mintCount + 1, RsSalesParam(mintCount).Value)
                        Case Else
                            .Col = enumSysParamGrid.Values : .Row = mintCount + 1
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                            .TypeMaxEditLen = RsSalesParam(mintCount).DefinedSize
                            Call .SetText(enumSysParamGrid.Values, mintCount + 1, RsSalesParam(mintCount).Value)
                    End Select
                End With
            Next
        End If
        RsSalesParam.Close()
        'for the Value of column 'description'
        mStrSQL = "select distinct s1.column_name,s1.description from sales_parameter_desc s1, syscolumns s2, sysobjects s3 where s1.unit_code='" & gstrUNITID & "' and s3.name='sales_parameter' and s1.column_name=s2.name"
        RsSalesParam.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        RsSalesParam.Open(mStrSQL, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        If RsSalesParam.RecordCount >= 1 Then
            RsSalesParam.MoveFirst()
            With Me.spdSalesParam
                For mintCount = 1 To RsSalesParam.RecordCount
                    For mintRow = 1 To .MaxRows
                        mValue = Nothing
                        Call .GetText(enumSysParamGrid.Col_name, mintRow, mValue)
                        If Trim(UCase(RsSalesParam("column_name").Value)) = Trim(UCase(mValue)) Then
                            .Col = enumSysParamGrid.Description : .Row = mintRow
                            Call .SetText(enumSysParamGrid.Description, mintRow, RsSalesParam("description").Value)
                            GoTo label
                        End If
                    Next
label:
                    RsSalesParam.MoveNext()
                Next
            End With
        End If
        RsSalesParam.Close()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    '    Private Function POPULATEGRID() As Object commented by Abhijit 15 FEB 2023
    '        'Author-                Shubhra Verma
    '        'Arguments              NIL
    '        'Return Value           NIL
    '        'Function Comments      this function is used to fill the values in a grid
    '        'Creation Date          21/jun/2006
    '        On Error GoTo ErrHandler
    '        mStrSQL = "select * from sales_parameter where unit_code='" & gstrUNITID & "'"
    '        RsSalesParam = New ADODB.Recordset
    '        RsSalesParam.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    '        RsSalesParam.Open(mStrSQL, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
    '        If RsSalesParam.RecordCount >= 1 Then
    '            Me.spdSalesParam.MaxRows = 0
    '            RsSalesParam.MoveFirst()
    '            For mintCount = 0 To RsSalesParam.Fields.Count - 1
    '                If RsSalesParam.Fields(mintCount).Name.ToLower = "unit_code" Then
    '                    Continue For
    '                End If
    '                With Me.spdSalesParam
    '                    .MaxRows = .MaxRows + 1
    '                    'for the Value of column 'column name'
    '                    .Col = enumSysParamGrid.Col_name : .Row = mintCount + 1
    '                    Call .SetText(enumSysParamGrid.Col_name, mintCount + 1, RsSalesParam.Fields(mintCount).Name)
    '                    'for the Value of column 'Value'
    '                    Select Case RsSalesParam(mintCount).Type
    '                        Case ADODB.DataTypeEnum.adSmallInt, ADODB.DataTypeEnum.adBigInt, ADODB.DataTypeEnum.adDecimal, ADODB.DataTypeEnum.adDouble, ADODB.DataTypeEnum.adInteger, ADODB.DataTypeEnum.adTinyInt, ADODB.DataTypeEnum.adUnsignedBigInt, ADODB.DataTypeEnum.adUnsignedInt, ADODB.DataTypeEnum.adVarNumeric, ADODB.DataTypeEnum.adUnsignedSmallInt
    '                            .Col = enumSysParamGrid.Values : .Row = mintCount + 1
    '                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
    '                            .TypeMaxEditLen = RsSalesParam(mintCount).DefinedSize
    '                            If IsDBNull(RsSalesParam(mintCount).Value) Then
    '                                Call .SetText(enumSysParamGrid.Values, mintCount + 1, 0)
    '                            Else
    '                                Call .SetText(enumSysParamGrid.Values, mintCount + 1, RsSalesParam(mintCount).Value)
    '                            End If
    '                        Case ADODB.DataTypeEnum.adNumeric
    '                            .Col = enumSysParamGrid.Values : .Row = mintCount + 1
    '                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                            .TypeMaxEditLen = RsSalesParam(mintCount).DefinedSize
    '                            .TypeFloatDecimalPlaces = 4
    '                            If Not IsDBNull(RsSalesParam(mintCount).Value) Then
    '                                Call .SetText(enumSysParamGrid.Values, mintCount + 1, Format(RsSalesParam(mintCount).Value, "0.0000"))
    '                            Else
    '                                Call .SetText(enumSysParamGrid.Values, mintCount + 1, Format(0, "0.0000"))
    '                            End If
    '                        Case ADODB.DataTypeEnum.adBoolean
    '                            .Col = enumSysParamGrid.Values : .Row = mintCount + 1
    '                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
    '                            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
    '                            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
    '                            If IsDBNull(RsSalesParam(mintCount).Value) Then
    '                                Call .SetText(enumSysParamGrid.Values, mintCount + 1, RsSalesParam.Fields(mintCount).Value)
    '                            ElseIf RsSalesParam(mintCount).Value Then
    '                                Call .SetText(enumSysParamGrid.Values, mintCount + 1, 1)
    '                            Else
    '                                Call .SetText(enumSysParamGrid.Values, mintCount + 1, RsSalesParam.Fields(mintCount).Value)
    '                            End If
    '                        Case Else
    '                            .Col = enumSysParamGrid.Values : .Row = mintCount + 1
    '                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
    '                            .TypeMaxEditLen = RsSalesParam(mintCount).DefinedSize
    '                            Call .SetText(enumSysParamGrid.Values, mintCount + 1, RsSalesParam(mintCount).Value)
    '                    End Select
    '                End With
    '            Next
    '        End If
    '        RsSalesParam.Close()
    '        'for the Value of column 'description'
    '        mStrSQL = "select distinct s1.column_name,s1.description from sales_parameter_desc s1, syscolumns s2, sysobjects s3 where s1.unit_code='" & gstrUNITID & "' and s3.name='sales_parameter' and s1.column_name=s2.name"
    '        RsSalesParam.CursorLocation = ADODB.CursorLocationEnum.adUseClient
    '        RsSalesParam.Open(mStrSQL, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
    '        If RsSalesParam.RecordCount >= 1 Then
    '            RsSalesParam.MoveFirst()
    '            With Me.spdSalesParam
    '                For mintCount = 1 To RsSalesParam.RecordCount
    '                    For mintRow = 1 To .MaxRows
    '                        mValue = Nothing
    '                        Call .GetText(enumSysParamGrid.Col_name, mintRow, mValue)
    '                        If Trim(UCase(RsSalesParam("column_name").Value)) = Trim(UCase(mValue)) Then
    '                            .Col = enumSysParamGrid.Description : .Row = mintRow
    '                            Call .SetText(enumSysParamGrid.Description, mintRow, RsSalesParam("description").Value)
    '                            GoTo label
    '                        End If
    '                    Next
    'label:
    '                    RsSalesParam.MoveNext()
    '                Next
    '            End With
    '        End If
    '        RsSalesParam.Close()
    '        Exit Function
    'ErrHandler:
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '    End Function
    Private Sub ctlMkt0015Hdr_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlMkt0015Hdr.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
End Class