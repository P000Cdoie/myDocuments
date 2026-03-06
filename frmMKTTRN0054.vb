Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMKTTRN0054
	Inherits System.Windows.Forms.Form
    '********************************************************************************************************
    'Copyright (c)  -   MIND
    'Name of module -   FRMMFGTRN0054.frm
    'Created By     -   Shubhra Verma
    'Created On     -   02-Aug-2007
    'description    -   Authorization of Pending Call Offs
    '               -   frmMKTTRN0054
    'Issue ID       -   20756
    'Purpose        -   Whenever Uploading the Release File,It should check
    '                   Transmission number for the SenderID/Customer Code in the
    '                   Release File.If the last number is not equal to Current
    '                   (in new Release)-1 then Alert for the Same alongwith mail.
    '                   Also Uploading should be allowed after Authorisation for
    '                   uploading of the same.
    '---------------------------------------------------------------------------------------
    'Modified by    :   Virendra Gupta
    'Modified ON    :   20/05/2011
    'Modified to support MultiUnit functionality
    '-----------------------------------------------------------------------
    '***********************************************************************************
    '********************************************************************************************************
    Dim mintIndex As Short
    Dim AUTHDATE As String
    Public bool_frm54 As Boolean = False
    Public cus_code As String = ""
    Public File_name As String = ""
    Private Enum ENUM_Grid
        check = 1
        calloffno
        CallOffDate
        MissingCallOff
        CustomerCode
        CustomerName
        ConsigneeCode
        ConsigneeName
        AUTHORIZATIONREMARKS
    End Enum
    Private Enum ENUM_DTL
        status = 1
        calloffno
        CallOffDate
        CustomerCode
        CustomerName
        ConsigneeCode
        ConsigneeName
    End Enum
    Public upld As Boolean
    Private Sub cmdGrpAuthorise_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As UCActXCtl.cmdGrpAuthorise.ButtonClickEventArgs) Handles cmdGrpAuthorise.ButtonClick
        On Error GoTo ErrHandler
        Dim enmValue As Short
        Dim YesNoCancel As String
        Select Case eventArgs.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_AUTHORIZE
                'Added by rajni if no record exist than disable the authorize button
                If Me.spdCAllOffs.MaxRows = 0 Then
                    MsgBox("No Record(s) Exist.", MsgBoxStyle.OkOnly)
                    Exit Sub
                End If
                Call FN_Save()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                If optAuth.Checked = True Then
                    Me.Close()
                Else
                    YesNoCancel = CStr(MsgBox("Do You Want To Authorize Any Call Off", MsgBoxStyle.YesNoCancel, ResolveResString(100)))
                    If YesNoCancel = CStr(MsgBoxResult.Yes) Then
                        'Save the data before closing
                        Call cmdGrpAuthorise_ButtonClick(cmdGrpAuthorise, New UCActXCtl.cmdGrpAuthorise.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_AUTHORIZE))
                    End If
                    If YesNoCancel = CStr(MsgBoxResult.No) Then
                        gblnCancelUnload = False
                        gblnFormAddEdit = False
                        Me.Close()
                    End If
                    If YesNoCancel = CStr(MsgBoxResult.Cancel) Then 'Set the global variable
                        gblnCancelUnload = True
                        gblnFormAddEdit = True
                        cmdGrpAuthorise.Focus()
                    End If
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_REFRESH
                txtAuthDate.Text = ""
                txtCustomer.Text = ""
                lblCustDesc.Text = ""
                txtCallOffNo.Text = ""
                spdCAllOffs.MaxRows = 0
                spdCallOffDtl.MaxRows = 0
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmdhelpAuthDate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdhelpAuthDate.Click
        On Error GoTo ErrHandler
        Dim strHelp() As String
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        Me.txtAuthDate.Text = ""
        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select distinct convert(varchar,authorizationdate,106) as AuthorizationDate, authorizedby from   AuthCallOffs_Hdr WHERE status = 'A' and Unit_Code = '" & gstrUNITID & "'", "Document Numbers")
        If UBound(strHelp) <> -1 Then
            If strHelp(0) <> "0" Then
                If strHelp(0) <> "" Then
                    txtAuthDate.Text = VB6.Format(strHelp(0), "DD MMM YYYY")
                End If
            Else
                MsgBox(" No CallOff Authorized", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub CmdHelpCallOffNo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdHelpCallOffNo.Click
        On Error GoTo ErrHandler
        Dim strHelp() As String
        Dim sql As String
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        If txtCustomer.Text = "" Then
            MsgBox("Please Select Customer Code.", MsgBoxStyle.OkOnly, ResolveResString(100))
            txtCallOffNo.Text = ""
            txtCustomer.Focus()
            Exit Sub
        End If
        sql = "SELECT DISTINCT CONVERT(VARCHAR(24),AUTHORIZATIONDATE,113) AS AuthorizationDate,  CallOffNo,CONVERT(VARCHAR,CALLOFFDATE,106) AS CallOffDate FROM AUTHCALLOFFS_HDR  WHERE CUSTOMER_CODE = '" & txtCustomer.Text & "' AND STATUS = 'A' and Unit_code = '" & gstrUNITID & "'"
        If txtAuthDate.Text <> "" Then
            sql = sql & "  AND CONVERT(VARCHAR(24),AUTHORIZATIONDATE,113) = '" & txtAuthDate.Text & "'"
        End If
        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, sql, "CallOff No ", 1)
        Dim str_Renamed As String
        If UBound(strHelp) <> -1 Then
            If strHelp(0) <> "0" Then
                If strHelp(0) <> "" Then
                    txtCallOffNo.Text = strHelp(1)
                End If
                AUTHDATE = ""
                AUTHDATE = VB.Left(strHelp(0), 24)
            Else
                MsgBox(" No record available", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
        End If
        If txtCallOffNo.Text <> "" Then
            Call txtCallOffNo_Validating(txtCallOffNo, New System.ComponentModel.CancelEventArgs(False))
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdhelpCust_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelpCust.Click
        On Error GoTo ErrHandler
        Dim strHelp() As String
        Dim sql As String
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        sql = "SELECT DISTINCT CUSTOMER_CODE ,CUSTOMER_NAME FROM AUTHCALLOFFS_HDR WHERE   ISNULL(STATUS,'') = 'A' and Unit_Code = '" & gstrUNITID & "'"
        If txtAuthDate.Text <> "" Then
            sql = sql & " AND convert(varchar,AuthorizationDate,106) = '" & txtAuthDate.Text & "' "
        End If
        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, sql, "Customer Code ")
        If UBound(strHelp) <> -1 Then
            If strHelp(0) <> "0" Then
                If strHelp(0) <> "" Then
                    Me.txtCustomer.Text = strHelp(0)
                    lblCustDesc.Text = strHelp(1)
                End If
            Else
                MsgBox(" No record available", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAuthDate_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAuthDate.TextChanged
        On Error GoTo ErrHandler
        If optAuth.Checked = True Then
            If txtAuthDate.Text > VB6.Format(GetServerDate, "DD MMM YYYY") Then
                txtAuthDate.Text = VB6.Format(GetServerDate, "DD MMM YYYY")
            End If
        End If
        txtCustomer.Text = ""
        lblCustDesc.Text = ""
        txtCallOffNo.Text = ""
        spdCallOffDtl.MaxRows = 0
        spdCAllOffs.MaxRows = 0
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0054_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex ''Holds the Form Name
        frmModules.NodeFontBold(Me.Tag) = True
        Exit Sub
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Call InitializeFormSettings()
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0054_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        'Load the caption
        Call FillLabelFromResFile(Me)
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.AppStarting)
        'Size the form to client workspace
        Call FitToClient(Me, (Me.FraMain), (Me.ctlFormHeader1), (Me.cmdGrpAuthorise), 500)
        'get the index of form in the window list
        mintIndex = mdifrmMain.AddFormNameToWindowList(Me.ctlFormHeader1.Tag)
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Me.cmdGrpAuthorise.Enabled(0) = True
        Call InitializeFormSettings() 'Initial Form Settings
        Me.optUnAuth.Checked = True
        If Me.spdCAllOffs.MaxRows = 0 Then
            Me.cmdGrpAuthorise.Enabled(0) = False
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0054_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0054_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        frmModules.NodeFontBold(Me.Tag) = False
        Me.Dispose()
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub InitializeFormSettings()
        On Error GoTo ErrHandler
        With spdCAllOffs
            .MaxCols = 9
            .MaxRows = 0
            .Row = 0
            .Col = ENUM_Grid.check : .Text = "Select"
            .Col = ENUM_Grid.calloffno : .Text = "CallOff No"
            .Col = ENUM_Grid.CallOffDate : .Text = "CallOff Date "
            .Col = ENUM_Grid.MissingCallOff : .Text = "Missing CallOff"
            .Col = ENUM_Grid.CustomerCode : .Text = "Customer Code"
            .Col = ENUM_Grid.CustomerName : .Text = "Customer Name"
            .Col = ENUM_Grid.ConsigneeCode : .Text = "Consignee Code"
            .Col = ENUM_Grid.ConsigneeName : .Text = "Consignee Name"
            .Col = ENUM_Grid.AUTHORIZATIONREMARKS : .Text = "Remarks"
            .Row = -1
            .Col = ENUM_Grid.check
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
            .Col = ENUM_Grid.calloffno
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.CallOffDate
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.MissingCallOff
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.CustomerCode
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.CustomerName
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.ConsigneeCode
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.ConsigneeName
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_Grid.AUTHORIZATIONREMARKS
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .set_ColWidth(ENUM_Grid.check, 5)
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .set_ColWidth(ENUM_Grid.calloffno, 10)
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .set_ColWidth(ENUM_Grid.CallOffDate, 10)
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
            .set_ColWidth(ENUM_Grid.MissingCallOff, 10)
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .set_ColWidth(ENUM_Grid.CustomerCode, 10)
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
            .set_ColWidth(ENUM_Grid.CustomerName, 20)
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
            .set_ColWidth(ENUM_Grid.ConsigneeCode, 10)
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
            .set_ColWidth(ENUM_Grid.ConsigneeName, 20)
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
            .set_ColWidth(ENUM_Grid.AUTHORIZATIONREMARKS, 50)
        End With
        With spdCallOffDtl
            .MaxCols = 7
            .MaxRows = 0
            .Row = 0
            .Col = ENUM_DTL.calloffno : .Text = "CallOff No"
            .Col = ENUM_DTL.CallOffDate : .Text = "CallOff Date "
            .Col = ENUM_DTL.CustomerCode : .Text = "Customer Code"
            .Col = ENUM_DTL.CustomerName : .Text = "Customer Name"
            .Col = ENUM_DTL.ConsigneeCode : .Text = "Consignee Code"
            .Col = ENUM_DTL.ConsigneeName : .Text = "Consignee Name"
            .Col = ENUM_DTL.status : .Text = "Status"
            .Row = -1
            .Col = ENUM_DTL.calloffno
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_DTL.CallOffDate
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_DTL.CustomerCode
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_DTL.CustomerName
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_DTL.ConsigneeCode
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_DTL.ConsigneeName
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .Col = ENUM_DTL.status
            .CellType = FPSpreadADO.CellTypeConstants.CellTypePicture
            .TypePictCenter = True
            .TypePictStretch = False
            .TypePictMaintainScale = True
            .set_ColWidth(ENUM_DTL.calloffno, 10)
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .set_ColWidth(ENUM_DTL.CallOffDate, 10)
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
            .set_ColWidth(ENUM_DTL.CustomerCode, 10)
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
            .set_ColWidth(ENUM_DTL.CustomerName, 20)
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
            .set_ColWidth(ENUM_DTL.ConsigneeCode, 10)
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
            .set_ColWidth(ENUM_DTL.ConsigneeName, 20)
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
            .set_ColWidth(ENUM_DTL.status, 10)
        End With
        optUnAuth.Checked = True
        txtCallOffNo.Enabled = False
        txtCustomer.Enabled = False
        cmdhelpAuthDate.Enabled = False
        txtAuthDate.Enabled = False
        txtAuthDate.Text = VB6.Format(GetServerDate, "DD MMM YYYY")
        CmdHelpCallOffNo.Enabled = False
        cmdHelpCust.Enabled = False
        spdCAllOffs.MaxRows = 1
        Call fn_UnAuthCallOffs()
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub optAuth_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAuth.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            With spdCAllOffs
                .MaxRows = 0
                .MaxRows = 0
                txtCallOffNo.Enabled = True
                txtCustomer.Enabled = True
                cmdhelpAuthDate.Enabled = True
                txtAuthDate.Enabled = True
                txtAuthDate.Text = ""
                CmdHelpCallOffNo.Enabled = True
                cmdHelpCust.Enabled = True
                .set_ColWidth(ENUM_Grid.check, 0)
                spdCallOffDtl.set_ColWidth(ENUM_DTL.status, 10)
                .Col = ENUM_Grid.AUTHORIZATIONREMARKS
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                cmdGrpAuthorise.Enabled(0) = False
                cmdGrpAuthorise.Enabled(1) = True
                If .MaxRows > 0 Then
                    .Col = ENUM_Grid.calloffno
                    .Row = 1
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    .Focus()
                End If
            End With
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub optUnAuth_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optUnAuth.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            With spdCAllOffs
                .MaxRows = 0
                spdCallOffDtl.MaxRows = 0
                txtCallOffNo.Enabled = False
                txtCustomer.Enabled = False
                cmdhelpAuthDate.Enabled = False
                txtAuthDate.Enabled = False
                txtAuthDate.Text = VB6.Format(GetServerDate, "DD MMM YYYY")
                CmdHelpCallOffNo.Enabled = False
                cmdHelpCust.Enabled = False
                txtCustomer.Text = ""
                lblCustDesc.Text = ""
                txtCallOffNo.Text = ""
                .set_ColWidth(ENUM_Grid.check, 10)
                spdCallOffDtl.set_ColWidth(ENUM_Grid.check, 10)
                .Col = ENUM_Grid.AUTHORIZATIONREMARKS
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                cmdGrpAuthorise.Enabled(0) = True
                cmdGrpAuthorise.Enabled(1) = False
                Call fn_UnAuthCallOffs()
                If .MaxRows > 0 Then
                    .Col = ENUM_Grid.calloffno
                    .Row = 1
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    .Focus()
                End If
            End With
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Function fn_UnAuthCallOffs() As Object
        On Error GoTo ErrHandler
        Dim rsUnAuth As New ClsResultSetDB
        Dim LastCallOff As String
        spdCAllOffs.MaxRows = 0
        spdCallOffDtl.MaxRows = 0
        Dim strDelete As String = ""
        strDelete = "delete from AUTHCALLOFFS_HDR where Rtrim(LASTCALLOFF)=Rtrim('Unknown') and Unit_Code = '" & gstrUNITID & "'"
        mP_Connection.Execute(strDelete, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        strDelete = ""
        strDelete = "delete from AUTHCALLOFFS_DTL where Rtrim(CallOffNo)=Rtrim('Unknown') and Unit_Code = '" & gstrUNITID & "'"
        mP_Connection.Execute(strDelete, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        rsUnAuth.GetResult("Select LASTCALLOFF from AUTHCALLOFFS_HDR where Unit_Code = '" & gstrUNITID & "'")
        LastCallOff = rsUnAuth.GetValue("LASTCALLOFF")
        If LastCallOff = "Unknown" Then
            rsUnAuth.ResultSetClose()
            rsUnAuth = Nothing
            Exit Function
        End If
        rsUnAuth.ResultSetClose()
        rsUnAuth = Nothing
        rsUnAuth = New ClsResultSetDB
        rsUnAuth.GetResult("SELECT DISTINCT H.CALLOFFNO  ,H.CALLOFFDATE, H.LASTCALLOFF+1 as MissingCallOff,H.CUSTOMER_CODE,H.CUSTOMER_NAME, D.Consignee_CODE,D.CONSIGNEE_NAME,H.AUTHORIZATIONREMARKS,H.status FROM AUTHCALLOFFS_HDR H, AUTHCALLOFFS_DTL  D WHERE ISNULL(STATUS,'')='' AND D.CALLOFFNO = H.CALLOFFNO AND D.UNIT_CODE = H.UNIT_CODE AND H.SENDERID = D.SENDERID and H.Unit_Code = '" & gstrUNITID & "' AND H.ENT_DT IN (SELECT MAX(ENT_DT) FROM AUTHCALLOFFS_HDR WHERE D.CALLOFFNO = H.CALLOFFNO AND H.SENDERID = D.SENDERID AND  Unit_Code = '" & gstrUNITID & "'  GROUP BY CUSTOMER_CODE, SENDERID)")
        If rsUnAuth.GetNoRows > 0 Then
            rsUnAuth.MoveFirst()
            spdCAllOffs.MaxRows = 1
            spdCAllOffs.Row = spdCAllOffs.MaxRows
            While Not rsUnAuth.EOFRecord
                spdCAllOffs.SetText(ENUM_Grid.calloffno, spdCAllOffs.MaxRows, rsUnAuth.GetValue("calloffno"))
                spdCAllOffs.SetText(ENUM_Grid.CallOffDate, spdCAllOffs.MaxRows, VB6.Format(rsUnAuth.GetValue("calloffdate"), "DD MMM YYYY"))
                spdCAllOffs.SetText(ENUM_Grid.MissingCallOff, spdCAllOffs.MaxRows, rsUnAuth.GetValue("MissingCallOff"))
                spdCAllOffs.SetText(ENUM_Grid.CustomerCode, spdCAllOffs.MaxRows, rsUnAuth.GetValue("Customer_CODE"))
                spdCAllOffs.SetText(ENUM_Grid.CustomerName, spdCAllOffs.MaxRows, rsUnAuth.GetValue("Customer_NAME"))
                spdCAllOffs.SetText(ENUM_Grid.ConsigneeCode, spdCAllOffs.MaxRows, rsUnAuth.GetValue("Consignee_CODE"))
                spdCAllOffs.SetText(ENUM_Grid.ConsigneeName, spdCAllOffs.MaxRows, rsUnAuth.GetValue("CONSIGNEE_NAME"))
                spdCAllOffs.SetText(ENUM_Grid.AUTHORIZATIONREMARKS, spdCAllOffs.MaxRows, rsUnAuth.GetValue("AUTHORIZATIONREMARKS"))
                rsUnAuth.MoveNext()
                spdCAllOffs.MaxRows = spdCAllOffs.MaxRows + 1
                spdCAllOffs.Row = spdCAllOffs.MaxRows
            End While
            spdCAllOffs.MaxRows = spdCAllOffs.MaxRows - 1
        End If
        rsUnAuth.ResultSetClose()
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function FN_Save() As Object
        On Error GoTo ErrHandler
        Dim rsFileName As New ClsResultSetDB
        Dim varcalloffNO As Object
        Dim varCallOffDate As Object
        Dim varMissingCallOff As Object
        Dim varCUSTOMER As Object
        Dim varconsignee As Object
        Dim varRemarks As Object
        Dim varSelect As Object
        upld = True
        With spdCAllOffs
            varSelect = Nothing
            .GetText(ENUM_Grid.check, .ActiveRow, varSelect)
            If varSelect = "" Or varSelect = "0" Then
                MsgBox("Please Select Any Row", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                Exit Function
            End If
            varcalloffNO = Nothing
            varCallOffDate = Nothing
            varMissingCallOff = Nothing
            varCUSTOMER = Nothing
            varconsignee = Nothing
            varRemarks = Nothing
            .GetText(ENUM_Grid.calloffno, .ActiveRow, varcalloffNO)
            .GetText(ENUM_Grid.CallOffDate, .ActiveRow, varCallOffDate)
            .GetText(ENUM_Grid.MissingCallOff, .ActiveRow, varMissingCallOff)
            .GetText(ENUM_Grid.CustomerCode, .ActiveRow, varCUSTOMER)
            .GetText(ENUM_Grid.ConsigneeCode, .ActiveRow, varconsignee)
            .GetText(ENUM_Grid.AUTHORIZATIONREMARKS, .ActiveRow, varRemarks)
        End With
        If varRemarks = "" Then
            MsgBox("Authorization Remarks Are Mandatory", MsgBoxStyle.OkOnly, ResolveResString(100))
            Exit Function
        End If
        mP_Connection.Execute("UPDATE AUTHCALLOFFS_HDR SET AuthorizedBy = '" & mP_User & "',AuthorizationDate = GETDATE(),  AuthorizationRemarks = '" & Replace(varRemarks, "'", "''") & "',STATUS = 'A' FROM AUTHCALLOFFS_HDR H, AUTHCALLOFFS_DTL D  WHERE H.LastCallOff = CAST('" & varMissingCallOff & "' AS INT) -1 AND  H.CUSTOMER_CODE = '" & varCUSTOMER & "' AND H.SENDERID = D.SENDERID AND H.CALLOFFNO = D.CALLOFFNO AND ISNULL(H.STATUS,'')='' AND H.ENT_DT = (SELECT MAX(H.ENT_DT) FROM AUTHCALLOFFS_HDR H, AUTHCALLOFFS_DTL D WHERE H.LastCallOff = CAST('" & varMissingCallOff & "' AS INT) -1 AND H.CUSTOMER_CODE = '" & varCUSTOMER & "' AND H.SENDERID = D.SENDERID AND H.CALLOFFNO = D.CALLOFFNO AND ISNULL(H.STATUS,'')='' and H.Unit_code = D.Unit_code and H.Unit_code = '" & gstrUNITID & "') and H.Unit_code = D.Unit_code and H.Unit_code = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.Execute(" UPDATE AUTHCALLOFFS_HDR SET STATUS = 'P' FROM AUTHCALLOFFS_HDR H, AUTHCALLOFFS_DTL D WHERE H.LastCallOff = CAST(" & varMissingCallOff & " AS INT) -1 AND H.CUSTOMER_CODE = '" & varCUSTOMER & "' AND H.SENDERID = D.SENDERID AND H.CALLOFFNO = D.CALLOFFNO AND ISNULL(H.STATUS,'')='' and H.Unit_code = D.Unit_code and H.Unit_code = '" & gstrUNITID & "' AND H.ENT_DT < (SELECT MAX(H.ENT_DT) FROM AUTHCALLOFFS_HDR H, AUTHCALLOFFS_DTL D WHERE H.LastCallOff = CAST(" & varMissingCallOff & " AS INT) -1 AND H.CUSTOMER_CODE = '" & varCUSTOMER & "' AND H.SENDERID = D.SENDERID AND H.CALLOFFNO = D.CALLOFFNO and H.Unit_code = D.Unit_code and H.Unit_code = '" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        spdCAllOffs.Col = 1
        spdCAllOffs.Row = 1
        spdCAllOffs.Action = FPSpreadADO.ActionConstants.ActionActiveCell
        spdCAllOffs.Focus()
        rsFileName.GetResult("SELECT FILENAME FROM AUTHCALLOFFS_HDR WHERE  CUSTOMER_CODE = '" & varCUSTOMER & "' AND CALLOFFNO = '" & varcalloffNO & "' and Unit_code = '" & gstrUNITID & "'")
        bool_frm54 = True
        cus_code = varCUSTOMER.ToString
        Call fn_UnAuthCallOffs()
        Dim objFrmChild As Form
        Dim bool_checkfrm28open As Boolean = False
        For Each objFrmChild In My.Forms.mdifrmMain.MdiChildren
            If (objFrmChild.Name.ToUpper.Equals("frmMKTTRN0028".ToUpper)) Then
                FRMMKTTRN0028.BringToFront()
                bool_checkfrm28open = True
                Call FRMMKTTRN0028.CmdUploadCSV_Click(Nothing, New System.EventArgs())
            End If
        Next
        If bool_checkfrm28open = False Then
            FRMMKTTRN0028.cust_code = varCUSTOMER.ToString
            FRMMKTTRN0028.file_name = rsFileName.GetValue("FILENAME").ToString
            File_name = rsFileName.GetValue("FILENAME").ToString
            FRMMKTTRN0028.Show()
        End If
        rsFileName.ResultSetClose()
        mP_Connection.Execute("EXEC AUTOMAILER '" & gstrUNITID & "','" & varcalloffNO & "', '" & varCallOffDate & "','CAuth'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub txtAuthDate_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAuthDate.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 39 Then
            eventArgs.Handled = True
        End If
        If KeyCode = 112 And Shift = 0 Then
            Call cmdhelpAuthDate_Click(cmdhelpAuthDate, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAuthDate_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtAuthDate.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim rsAuthDate As New ClsResultSetDB
        If txtAuthDate.Text <> "" Then
            rsAuthDate.GetResult("Select distinct convert(varchar,authorizationdate,106)  as AuthorizationDate, authorizedby from AuthCallOffs_Hdr WHERE status = 'A'  and convert(varchar,'" & txtAuthDate.Text & "',106) = convert(varchar,authorizationdate,106)  and Unit_Code = '" & gstrUNITID & "'")
            If rsAuthDate.GetNoRows = 0 Then
                MsgBox("Invalid Date.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                rsAuthDate.ResultSetClose()
                GoTo EventExitSub
            End If
        End If
        rsAuthDate.ResultSetClose()
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtCallOffNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCallOffNo.TextChanged
        On Error GoTo ErrHandler
        spdCallOffDtl.MaxRows = 0
        spdCAllOffs.MaxRows = 0
        AUTHDATE = ""
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtCallOffNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCallOffNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 39 Then
            eventArgs.Handled = True
        End If
        If KeyCode = 112 And Shift = 0 Then
            Call CmdHelpCallOffNo_Click(CmdHelpCallOffNo, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCallOffNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCallOffNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then
            Call FN_AUTHCALLOFFS()
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
    Private Sub txtCallOffNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCallOffNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim rsAuthDate As New ClsResultSetDB
        Dim sql As String
        If txtCallOffNo.Text <> "" Then
            If txtCustomer.Text = "" Then
                MsgBox("Please Select Customer Code.", MsgBoxStyle.OkOnly, ResolveResString(100))
                txtCallOffNo.Text = ""
                txtCustomer.Focus()
                GoTo EventExitSub
            End If
            sql = "SELECT DISTINCT CALLOFFNO,CONVERT(VARCHAR,CALLOFFDATE,106) AS CALLOFFDATE FROM AUTHCALLOFFS_HDR  WHERE CUSTOMER_CODE = '" & txtCustomer.Text & "' AND STATUS = 'A' and calloffno = '" & txtCallOffNo.Text & "'  and Unit_Code = '" & gstrUNITID & "'"
            If txtAuthDate.Text <> "" Then
                sql = sql & "  AND CONVERT(VARCHAR,AUTHORIZATIONDATE,106) = '" & txtAuthDate.Text & "' "
            End If
            rsAuthDate.GetResult(sql)
            If rsAuthDate.GetNoRows = 0 Then
                MsgBox("Invalid CallOff.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                txtCallOffNo.Focus()
                rsAuthDate.ResultSetClose()
                GoTo EventExitSub
            Else
                Call FN_AUTHCALLOFFS()
            End If
        End If
        rsAuthDate.ResultSetClose()
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtcustomer_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustomer.TextChanged
        On Error GoTo ErrHandler
        txtCallOffNo.Text = ""
        spdCallOffDtl.MaxRows = 0
        spdCAllOffs.MaxRows = 0
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Function FN_AUTHCALLOFFS() As Object
        On Error GoTo ErrHandler
        Dim RSAUTH As New ClsResultSetDB
        Dim rsMISSING As New ClsResultSetDB
        Dim INTMiss As Integer
        Dim Range As Integer
        Dim varMissingCallOff As Object
        Dim varCUSTOMER As Object
        Dim sql As String
        Dim varcalloffNO As Object
        sql = ""
        If Len(LTrim(RTrim(AUTHDATE))) > 0 Then
            sql = "AND AUTHORIZATIONDATE = '" & AUTHDATE & "'"
        End If
        RSAUTH.GetResult("SELECT DISTINCT H.CALLOFFNO  ,H.CALLOFFDATE,  H.LASTCALLOFF+1 as MissingCallOff,H.CUSTOMER_CODE,H.CUSTOMER_NAME,  D.CONSIGNEE_CODE , D.CONSIGNEE_NAME,H.AUTHORIZATIONREMARKS, H.status,  H.Ent_dt FROM AUTHCALLOFFS_HDR H, AUTHCALLOFFS_DTL D  WHERE H.CUSTOMER_CODE = '" & txtCustomer.Text & "'  AND H.SENDERID = D.SENDERID AND H.UNIT_CODE = D.UNIT_CODE AND H.CALLOFFNO = D.CALLOFFNO AND  H.STATUS = 'A' AND H.CALLOFFNO = '" & txtCallOffNo.Text & "'  and H.Unit_Code = '" & gstrUNITID & "'" & sql)
        If RSAUTH.GetNoRows > 0 Then
            RSAUTH.MoveFirst()
            spdCAllOffs.MaxRows = 1
            spdCAllOffs.Row = spdCAllOffs.MaxRows
            While Not RSAUTH.EOFRecord
                spdCAllOffs.SetText(ENUM_Grid.calloffno, spdCAllOffs.MaxRows, RSAUTH.GetValue("calloffno"))
                spdCAllOffs.SetText(ENUM_Grid.CallOffDate, spdCAllOffs.MaxRows, VB6.Format(RSAUTH.GetValue("calloffdate"), "DD MMM YYYY"))
                spdCAllOffs.SetText(ENUM_Grid.MissingCallOff, spdCAllOffs.MaxRows, RSAUTH.GetValue("MissingCallOff"))
                spdCAllOffs.SetText(ENUM_Grid.CustomerCode, spdCAllOffs.MaxRows, RSAUTH.GetValue("Customer_CODE"))
                spdCAllOffs.SetText(ENUM_Grid.CustomerName, spdCAllOffs.MaxRows, RSAUTH.GetValue("Customer_NAME"))
                spdCAllOffs.SetText(ENUM_Grid.ConsigneeCode, spdCAllOffs.MaxRows, RSAUTH.GetValue("Consignee_CODE"))
                spdCAllOffs.SetText(ENUM_Grid.ConsigneeName, spdCAllOffs.MaxRows, RSAUTH.GetValue("CONSIGNEE_NAME"))
                spdCAllOffs.SetText(ENUM_Grid.AUTHORIZATIONREMARKS, spdCAllOffs.MaxRows, RSAUTH.GetValue("AUTHORIZATIONREMARKS"))
                RSAUTH.MoveNext()
                spdCAllOffs.MaxRows = spdCAllOffs.MaxRows + 1
                spdCAllOffs.Row = spdCAllOffs.MaxRows
            End While
            spdCAllOffs.MaxRows = spdCAllOffs.MaxRows - 1
        End If
        If spdCAllOffs.MaxRows > 0 Then
            varCUSTOMER = Nothing
            varcalloffNO = Nothing
            varMissingCallOff = Nothing
            spdCAllOffs.GetText(ENUM_Grid.CustomerCode, spdCAllOffs.MaxRows, varCUSTOMER)
            spdCAllOffs.GetText(ENUM_Grid.calloffno, spdCAllOffs.MaxRows, varcalloffNO)
            spdCAllOffs.GetText(ENUM_Grid.MissingCallOff, spdCAllOffs.MaxRows, varMissingCallOff)
            rsMISSING.GetResult("SELECT MIN(CALLOFFNO) as CALLOFFNO FROM AUTHCALLOFFS_HDR  Where LASTCALLOFF = CAST('" & varMissingCallOff & "' AS INT) -1  AND CALLOFFNO > LASTCALLOFF AND STATUS = 'P' and Unit_Code = '" & gstrUNITID & "'")
            sql = "SELECT DISTINCT H.CALLOFFNO  ,H.CALLOFFDATE,  H.LASTCALLOFF+1 as MissingCallOff,H.CUSTOMER_CODE,H.CUSTOMER_NAME,  D.CONSIGNEE_CODE , D.CONSIGNEE_NAME, H.AUTHORIZATIONREMARKS, H.status,  H.Ent_dt  FROM AUTHCALLOFFS_HDR H, AUTHCALLOFFS_DTL D  WHERE H.LastCallOff = CAST('" & varMissingCallOff & "' AS INT)-1 AND  H.CUSTOMER_CODE = '" & varCUSTOMER & "' AND H.SENDERID = D.SENDERID  AND H.CALLOFFNO = D.CALLOFFNO  AND H.STATUS ='P' and H.Unit_code = D.Unit_code and H.Unit_code = '" & gstrUNITID & "'  AND H.ENT_DT NOT IN (SELECT TOP 1 H.ENT_DT  FROM AUTHCALLOFFS_HDR H, AUTHCALLOFFS_DTL D  WHERE H.LastCallOff = CAST('" & varMissingCallOff & "'  AS INT) -1 AND  H.CUSTOMER_CODE = '" & varCUSTOMER & "' AND H.SENDERID = D.SENDERID and H.Unit_code = D.Unit_code and H.Unit_code = '" & gstrUNITID & "'  AND H.CALLOFFNO = D.CALLOFFNO    ORDER BY H.ENT_DT DESC)   ORDER BY H.ENT_DT"
            spdCallOffDtl.MaxRows = 1
            RSAUTH.ResultSetClose()
            RSAUTH = New ClsResultSetDB
            RSAUTH.GetResult(sql)
            If RSAUTH.GetNoRows > 0 Then
                RSAUTH.MoveFirst()
                If (IIf(IsDBNull(rsMISSING.GetValue("calloffno")), "", rsMISSING.GetValue("calloffno")) <> "") And (RSAUTH.GetNoRows > 0) Then
                    If rsMISSING.GetValue("calloffno") > RSAUTH.GetValue("missingcalloff") Then
                        INTMiss = RSAUTH.GetValue("missingcalloff")
                    Else
                        INTMiss = rsMISSING.GetValue("calloffno")
                    End If
                Else
                    INTMiss = RSAUTH.GetValue("missingcalloff")
                End If
                RSAUTH.MoveFirst()
                Range = RSAUTH.GetValue("calloffno")
            Else
                If Val(varMissingCallOff) < Val(varcalloffNO) Then
                    INTMiss = IIf(varMissingCallOff = "", 0, varMissingCallOff)
                    Range = IIf(varcalloffNO = "", 0, varcalloffNO)
                End If
            End If
            If rsMISSING.GetNoRows > 0 Then
                While INTMiss < Range
                    spdCallOffDtl.SetText(ENUM_DTL.calloffno, spdCallOffDtl.MaxRows, INTMiss)
                    spdCallOffDtl.Col = ENUM_DTL.calloffno
                    spdCallOffDtl.BorderStyle = FPSpreadADO.BorderStyleConstants.BorderStyleFixedSingle
                    spdCallOffDtl.SetText(ENUM_DTL.CallOffDate, spdCallOffDtl.MaxRows, "")
                    spdCallOffDtl.SetText(ENUM_DTL.CustomerCode, spdCallOffDtl.MaxRows, "")
                    spdCallOffDtl.SetText(ENUM_DTL.CustomerName, spdCallOffDtl.MaxRows, "")
                    spdCallOffDtl.SetText(ENUM_DTL.ConsigneeCode, spdCallOffDtl.MaxRows, "")
                    spdCallOffDtl.SetText(ENUM_DTL.ConsigneeName, spdCallOffDtl.MaxRows, "")
                    spdCallOffDtl.Row = spdCallOffDtl.MaxRows
                    spdCallOffDtl.Col = ENUM_DTL.status
                    spdCallOffDtl.TypePictPicture = ImgRed.Image
                    INTMiss = INTMiss + 1
                    spdCallOffDtl.MaxRows = spdCallOffDtl.MaxRows + 1
                End While
            End If
            rsMISSING.ResultSetClose()
            If RSAUTH.GetNoRows > 0 Then
                RSAUTH.MoveFirst()
                While Not RSAUTH.EOFRecord
                    spdCallOffDtl.SetText(ENUM_DTL.calloffno, spdCallOffDtl.MaxRows, RSAUTH.GetValue("calloffno"))
                    spdCallOffDtl.Col = ENUM_DTL.calloffno
                    spdCallOffDtl.BorderStyle = FPSpreadADO.BorderStyleConstants.BorderStyleFixedSingle
                    spdCallOffDtl.SetText(ENUM_DTL.CallOffDate, spdCallOffDtl.MaxRows, VB6.Format(RSAUTH.GetValue("calloffdate"), "DD MMM YYYY"))
                    spdCallOffDtl.SetText(ENUM_DTL.CustomerCode, spdCallOffDtl.MaxRows, RSAUTH.GetValue("Customer_CODE"))
                    spdCallOffDtl.SetText(ENUM_DTL.CustomerName, spdCallOffDtl.MaxRows, RSAUTH.GetValue("Customer_NAME"))
                    spdCallOffDtl.SetText(ENUM_DTL.ConsigneeCode, spdCallOffDtl.MaxRows, RSAUTH.GetValue("Consignee_CODE"))
                    spdCallOffDtl.SetText(ENUM_DTL.ConsigneeName, spdCallOffDtl.MaxRows, RSAUTH.GetValue("CONSIGNEE_NAME"))
                    spdCallOffDtl.Row = spdCallOffDtl.MaxRows
                    spdCallOffDtl.Col = ENUM_DTL.status
                    spdCallOffDtl.TypePictPicture = ImgYellow.Image
                    RSAUTH.MoveNext()
                    spdCallOffDtl.MaxRows = spdCallOffDtl.MaxRows + 1
                End While
                RSAUTH.ResultSetClose()
            End If
            spdCallOffDtl.MaxRows = spdCallOffDtl.MaxRows - 1
        End If
        Timer1.Start()
        Timer1.Enabled = True
        cmdGrpAuthorise.Focus()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub txtCustomer_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustomer.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 39 Then
            eventArgs.Handled = True
        End If
        If KeyCode = 112 And Shift = 0 Then
            Call cmdhelpCust_Click(cmdHelpCust, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtcustomer_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustomer.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim sql As String
        Dim rsCust As New ClsResultSetDB
        If txtCustomer.Text <> "" Then
            sql = "SELECT DISTINCT CUSTOMER_CODE ,CUSTOMER_NAME FROM AUTHCALLOFFS_HDR WHERE   ISNULL(STATUS,'') = 'A' AND CUSTOMER_CODE = '" & txtCustomer.Text & "' and Unit_Code = '" & gstrUNITID & "'"
            If txtAuthDate.Text <> "" Then
                sql = sql & " AND convert(varchar,AuthorizationDate,106) = '" & txtAuthDate.Text & "' "
            End If
            rsCust.GetResult(sql)
            If rsCust.GetNoRows = 0 Then
                MsgBox("Invalid Customer Code", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                txtCustomer.Text = ""
                txtCustomer.Focus()
            End If
        End If
        rsCust.ResultSetClose()
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub spdCAllOffs_ClickEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles spdCAllOffs.ClickEvent
        On Error GoTo ErrHandler
        Dim sql As String, sql1 As String
        Dim varcalloffNO As Object
        Dim varMissingCallOff As Object
        Dim varCUSTOMER As Object
        Dim rsUnAuth As New ClsResultSetDB
        Dim rsMISSING As New ClsResultSetDB
        Dim rsRange As New ClsResultSetDB
        'Dim INTMiss As Long
        'Dim Range As Long
        Dim INTMiss As String
        Dim Range As String
        Dim intcol As Integer
        Dim intRow As Integer
        Dim val As Object
        intRow = e.row
        intcol = e.col
        ImgRed.Height = 270
        ImgRed.Width = 250
        ImgRed1.Height = 270
        ImgRed1.Width = 250
        ImgYellow.Height = 270
        ImgYellow.Width = 250
        ImgYellow1.Height = 270
        ImgYellow1.Width = 250
        If e.col = 1 Then
            With spdCAllOffs
                .Row = 1
                .Row2 = .ActiveRow - 1
                .Col = 1
                .Col2 = 1
                .BlockMode = True
                .Value = 0
                .Row = .ActiveRow + 1
                .Row2 = .MaxRows
                .Col = 1
                .Col2 = 1
                .BlockMode = True
                .Value = 0
            End With
            varCUSTOMER = Nothing
            varcalloffNO = Nothing
            varMissingCallOff = Nothing
            spdCAllOffs.GetText(ENUM_Grid.CustomerCode, spdCAllOffs.ActiveRow, varCUSTOMER)
            spdCAllOffs.GetText(ENUM_Grid.calloffno, spdCAllOffs.ActiveRow, varcalloffNO)
            spdCAllOffs.GetText(ENUM_Grid.MissingCallOff, spdCAllOffs.ActiveRow, varMissingCallOff)
            rsMISSING.GetResult("SELECT MIN(CALLOFFNO) as CALLOFFNO FROM AUTHCALLOFFS_HDR" & _
               " Where LASTCALLOFF = CAST('" & varMissingCallOff & "' AS INT) -1" & _
               " AND CALLOFFNO > LASTCALLOFF AND ISNULL(STATUS,'') = '' and Unit_Code = '" & gstrUNITID & "' ")
            sql = "SELECT DISTINCT H.CALLOFFNO  ,H.CALLOFFDATE," & _
               " H.LASTCALLOFF+1 as MissingCallOff,H.CUSTOMER_CODE,H.CUSTOMER_NAME," & _
               " D.CONSIGNEE_CODE , D.CONSIGNEE_NAME, H.AUTHORIZATIONREMARKS, H.status," & _
               " H.Ent_dt" & _
               " FROM AUTHCALLOFFS_HDR H, AUTHCALLOFFS_DTL D" & _
               " WHERE H.LastCallOff = CAST('" & varMissingCallOff & "' AS INT)-1 AND" & _
               " H.CUSTOMER_CODE = '" & varCUSTOMER & "' AND H.SENDERID = D.SENDERID AND H.UNIT_CODE = D.UNIT_CODE" & _
               " AND H.CALLOFFNO = D.CALLOFFNO" & _
               " AND ISNULL(H.STATUS,'')='' AND  H.Unit_Code = '" & gstrUNITID & "'" & _
               " AND H.ENT_DT NOT IN (SELECT TOP 1 H.ENT_DT" & _
               " FROM AUTHCALLOFFS_HDR H, AUTHCALLOFFS_DTL D" & _
               " WHERE H.LastCallOff = CAST('" & varMissingCallOff & "'  AS INT) -1 AND" & _
               " H.CUSTOMER_CODE = '" & varCUSTOMER & "' AND H.SENDERID = D.SENDERID" & _
               " AND H.CALLOFFNO = D.CALLOFFNO" & _
               " AND ISNULL(H.STATUS,'')='' and H.Unit_code = D.Unit_code and H.Unit_code = '" & gstrUNITID & "'" & _
               " ORDER BY H.ENT_DT DESC) " & _
               " ORDER BY H.ENT_DT"
            spdCallOffDtl.MaxRows = 1
            rsUnAuth.GetResult(sql)
            If rsUnAuth.GetNoRows > 0 Then
                rsUnAuth.MoveFirst()
                If (IIf(IsDBNull(rsMISSING.GetValue("calloffno")), "", rsMISSING.GetValue("calloffno")) <> "") And (rsUnAuth.GetNoRows > 0) Then
                    If rsMISSING.GetValue("calloffno") > rsUnAuth.GetValue("missingcalloff") Then
                        INTMiss = rsUnAuth.GetValue("missingcalloff")
                    Else
                        INTMiss = rsMISSING.GetValue("calloffno")
                    End If
                Else
                    INTMiss = rsUnAuth.GetValue("missingcalloff")
                End If
                rsUnAuth.MoveFirst()
                Range = rsUnAuth.GetValue("calloffno")
            Else
                If varMissingCallOff < varcalloffNO Then
                    INTMiss = IIf(varMissingCallOff = "", 0, varMissingCallOff)
                    Range = IIf(varcalloffNO = "", 0, varcalloffNO)
                End If
            End If
            If rsMISSING.GetNoRows > 0 Then
                While INTMiss < Range
                    spdCallOffDtl.SetText(ENUM_DTL.calloffno, spdCallOffDtl.MaxRows, INTMiss.ToString)
                    spdCallOffDtl.Col = ENUM_DTL.calloffno
                    spdCallOffDtl.BorderStyle = FPSpreadADO.BorderStyleConstants.BorderStyleFixedSingle
                    spdCallOffDtl.SetText(ENUM_DTL.CallOffDate, spdCallOffDtl.MaxRows, "")
                    spdCallOffDtl.SetText(ENUM_DTL.CustomerCode, spdCallOffDtl.MaxRows, "")
                    spdCallOffDtl.SetText(ENUM_DTL.CustomerName, spdCallOffDtl.MaxRows, "")
                    spdCallOffDtl.SetText(ENUM_DTL.ConsigneeCode, spdCallOffDtl.MaxRows, "")
                    spdCallOffDtl.SetText(ENUM_DTL.ConsigneeName, spdCallOffDtl.MaxRows, "")
                    spdCallOffDtl.Row = spdCallOffDtl.MaxRows
                    spdCallOffDtl.Col = ENUM_DTL.status
                    spdCallOffDtl.TypePictPicture = ImgRed.Image
                    INTMiss = INTMiss + 1
                    spdCallOffDtl.MaxRows = spdCallOffDtl.MaxRows + 1
                End While
            End If
            rsMISSING.ResultSetClose()
            If rsUnAuth.GetNoRows > 0 Then
                rsUnAuth.MoveFirst()
                While Not rsUnAuth.EOFRecord
                    spdCallOffDtl.SetText(ENUM_DTL.calloffno, spdCallOffDtl.MaxRows, rsUnAuth.GetValue("calloffno"))
                    spdCallOffDtl.Col = ENUM_DTL.calloffno
                    spdCallOffDtl.BorderStyle = FPSpreadADO.BorderStyleConstants.BorderStyleFixedSingle
                    spdCallOffDtl.SetText(ENUM_DTL.CallOffDate, spdCallOffDtl.MaxRows, VB6.Format(rsUnAuth.GetValue("calloffdate"), "DD MMM YYYY"))
                    spdCallOffDtl.SetText(ENUM_DTL.CustomerCode, spdCallOffDtl.MaxRows, rsUnAuth.GetValue("Customer_CODE"))
                    spdCallOffDtl.SetText(ENUM_DTL.CustomerName, spdCallOffDtl.MaxRows, rsUnAuth.GetValue("Customer_NAME"))
                    spdCallOffDtl.SetText(ENUM_DTL.ConsigneeCode, spdCallOffDtl.MaxRows, rsUnAuth.GetValue("Consignee_CODE"))
                    spdCallOffDtl.SetText(ENUM_DTL.ConsigneeName, spdCallOffDtl.MaxRows, rsUnAuth.GetValue("CONSIGNEE_NAME"))
                    spdCallOffDtl.Row = spdCallOffDtl.MaxRows
                    spdCallOffDtl.Col = ENUM_DTL.status
                    spdCallOffDtl.TypePictPicture = ImgYellow.Image
                    rsUnAuth.MoveNext()
                    spdCallOffDtl.MaxRows = spdCallOffDtl.MaxRows + 1
                End While
            End If
            spdCallOffDtl.MaxRows = spdCallOffDtl.MaxRows - 1
            rsUnAuth.ResultSetClose()
            cmdGrpAuthorise.Focus()
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub spdCAllOffs_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles spdCAllOffs.KeyPressEvent
        On Error GoTo ErrHandler
        Dim sql As String, sql1 As String
        Dim varcalloffNO As Object
        Dim varMissingCallOff As Object
        Dim varCUSTOMER As Object
        Dim rsUnAuth As New ClsResultSetDB
        Dim rsMISSING As New ClsResultSetDB
        Dim rsRange As New ClsResultSetDB
        Dim INTMiss As Long
        Dim Range As Long
        Dim intcol As Integer
        Dim intRow As Integer
        If spdCAllOffs.Row = spdCAllOffs.ActiveRow And spdCAllOffs.Col = 1 Then
            If e.keyAscii = Keys.Space Then ' vbKeySpace Then
                ImgRed.Height = 270
                ImgRed.Width = 250
                ImgRed1.Height = 270
                ImgRed1.Width = 250
                ImgYellow.Height = 270
                ImgYellow.Width = 250
                ImgYellow1.Height = 270
                ImgYellow1.Width = 250
                If spdCAllOffs.ActiveCol = 1 Then
                    With spdCAllOffs
                        .Row = 1
                        .Row2 = .ActiveRow - 1
                        .Col = 1
                        .Col2 = 1
                        .BlockMode = True
                        .Value = 0
                        .Row = .ActiveRow + 1
                        .Row2 = .MaxRows
                        .Col = 1
                        .Col2 = 1
                        .BlockMode = True
                        .Value = 0
                    End With
                    varCUSTOMER = Nothing
                    varcalloffNO = Nothing
                    varMissingCallOff = Nothing
                    spdCAllOffs.GetText(ENUM_Grid.CustomerCode, spdCAllOffs.ActiveRow, varCUSTOMER)
                    spdCAllOffs.GetText(ENUM_Grid.calloffno, spdCAllOffs.ActiveRow, varcalloffNO)
                    spdCAllOffs.GetText(ENUM_Grid.MissingCallOff, spdCAllOffs.ActiveRow, varMissingCallOff)
                    rsMISSING.GetResult("select max(convert(numeric,calloffno)) + 1 as calloffno from authcalloffs_hdr where customer_code = '" & varCUSTOMER & "' and status = 'A' and Unit_Code = '" & gstrUNITID & "'")
                    sql = "SELECT DISTINCT H.CALLOFFNO,H.CALLOFFDATE," & _
                       " H.LASTCALLOFF + 1 AS MISSINGCALLOFF,H.CUSTOMER_CODE,H.CUSTOMER_NAME,D.consignee_CODE," & _
                       " D.CONSIGNEE_NAME,ISNULL(H.AUTHORIZATIONREMARKS,'') as AUTHORIZATIONREMARKS,H.status" & _
                       " FROM AUTHCALLOFFS_HDR H, AUTHCALLOFFS_DTL  D" & _
                       " Where D.CALLOFFNO = H.CALLOFFNO AND  isnull(h.status,'') = '' " & _
                       " AND CUSTOMER_CODE  = '" & varCUSTOMER & "' AND H.CALLOFFNO < '" & varcalloffNO & "' AND H.CALLOFFNO >= '" & varMissingCallOff & "' and H.Unit_code = D.Unit_code and H.Unit_code = '" & gstrUNITID & "' order by h.calloffno"
                    spdCallOffDtl.MaxRows = 1
                    rsUnAuth.GetResult(sql)
                    If rsUnAuth.GetNoRows > 0 Then
                        rsUnAuth.MoveFirst()
                        If rsMISSING.GetNoRows > 0 And rsUnAuth.GetNoRows > 0 Then
                            If rsMISSING.GetValue("calloffno") > rsUnAuth.GetValue("missingcalloff") Then
                                INTMiss = rsMISSING.GetValue("calloffno")
                            Else
                                INTMiss = rsUnAuth.GetValue("missingcalloff")
                            End If
                        Else
                            INTMiss = rsUnAuth.GetValue("missingcalloff")
                        End If
                        rsUnAuth.MoveFirst()
                        Range = rsUnAuth.GetValue("calloffno")
                    ElseIf rsMISSING.GetNoRows > 0 Then
                        INTMiss = rsMISSING.GetValue("calloffno")
                        Range = IIf(varcalloffNO = "", 0, varcalloffNO)
                        rsRange.GetResult("select min(convert(numeric,calloffno)) as calloffno from authcalloffs_hdr where calloffno < " & varcalloffNO & " and calloffno > " & INTMiss & " and customer_code = '" & varCUSTOMER & "' and Unit_Code = '" & gstrUNITID & "'")
                        If rsRange.GetNoRows > 0 And IIf(IsDBNull(rsRange.GetValue("calloffno")), "", rsRange.GetValue("calloffno")) <> "" Then
                            Range = rsRange.GetValue("calloffno")
                        End If
                    Else
                        INTMiss = IIf(varMissingCallOff = "", 0, varMissingCallOff)
                        Range = IIf(varcalloffNO = "", 0, varcalloffNO)
                    End If
                    rsMISSING.GetResult("SELECT DISTINCT H.CALLOFFNO,H.CALLOFFDATE, " & _
                       " H.LASTCALLOFF + 1 AS MISSINGCALLOFF,h.CUSTOMER_CODE ," & _
                       " h.CUSTOMER_NAME,h.status FROM AUTHCALLOFFS_HDR H, AUTHCALLOFFS_DTL  D" & _
                       " Where D.CALLOFFNO = H.CALLOFFNO AND  isnull(h.status,'') = '' " & _
                       " AND CUSTOMER_CODE  = '" & varCUSTOMER & "' AND H.CALLOFFNO = '" & varcalloffNO & "' and H.Unit_code = D.Unit_code and H.Unit_code = '" & gstrUNITID & "'")
                    If rsMISSING.GetNoRows > 0 Then
                        While INTMiss < Range
                            spdCallOffDtl.SetText(ENUM_DTL.calloffno, spdCallOffDtl.MaxRows, INTMiss)
                            spdCallOffDtl.Col = ENUM_DTL.calloffno
                            spdCallOffDtl.BorderStyle = FPSpreadADO.BorderStyleConstants.BorderStyleFixedSingle
                            spdCallOffDtl.SetText(ENUM_DTL.CallOffDate, spdCallOffDtl.MaxRows, "")
                            spdCallOffDtl.SetText(ENUM_DTL.CustomerCode, spdCallOffDtl.MaxRows, rsMISSING.GetValue("Customer_CODE"))
                            spdCallOffDtl.SetText(ENUM_DTL.CustomerName, spdCallOffDtl.MaxRows, rsMISSING.GetValue("Customer_NAME"))
                            spdCallOffDtl.SetText(ENUM_DTL.ConsigneeCode, spdCallOffDtl.MaxRows, "")
                            spdCallOffDtl.SetText(ENUM_DTL.ConsigneeName, spdCallOffDtl.MaxRows, "")
                            spdCallOffDtl.Row = spdCallOffDtl.MaxRows
                            spdCallOffDtl.Col = ENUM_DTL.status
                            spdCallOffDtl.TypePictPicture = ImgRed.Image
                            INTMiss = INTMiss + 1
                            spdCallOffDtl.MaxRows = spdCallOffDtl.MaxRows + 1
                        End While
                    End If
                    rsMISSING.ResultSetClose()
                    rsUnAuth.GetResult("SELECT DISTINCT H.CALLOFFNO,H.CALLOFFDATE, " & _
                       " H.LASTCALLOFF + 1 AS MISSINGCALLOFF,H.CUSTOMER_CODE,H.CUSTOMER_NAME," & _
                       " D.consignee_CODE, D.CONSIGNEE_NAME,ISNULL(H.AUTHORIZATIONREMARKS,'') " & _
                       " as AUTHORIZATIONREMARKS,H.status FROM AUTHCALLOFFS_HDR H," & _
                       " AUTHCALLOFFS_DTL  D Where D.CALLOFFNO = H.CALLOFFNO AND" & _
                       " isnull(h.status,'') = '' AND CUSTOMER_CODE  = '" & varCUSTOMER & "'" & _
                       " AND H.CALLOFFNO < '" & varcalloffNO & "' and H.Unit_code = D.Unit_code and H.Unit_code = '" & gstrUNITID & "' AND " & _
                       " H.CALLOFFNO >(select max(convert(numeric,calloffno)) from authcalloffs_hdr" & _
                       " where status = 'A' and customer_code = '" & varCUSTOMER & "' and Unit_code = '" & gstrUNITID & "')")
                    If rsUnAuth.GetNoRows > 0 Then
                        rsUnAuth.MoveFirst()
                        While Not rsUnAuth.EOFRecord
                            spdCallOffDtl.SetText(ENUM_DTL.calloffno, spdCallOffDtl.MaxRows, rsUnAuth.GetValue("calloffno"))
                            spdCallOffDtl.Col = ENUM_DTL.calloffno
                            spdCallOffDtl.BorderStyle = FPSpreadADO.BorderStyleConstants.BorderStyleFixedSingle
                            spdCallOffDtl.SetText(ENUM_DTL.CallOffDate, spdCallOffDtl.MaxRows, Format(rsUnAuth.GetValue("calloffdate"), "DD MMM YYYY"))
                            spdCallOffDtl.SetText(ENUM_DTL.CustomerCode, spdCallOffDtl.MaxRows, rsUnAuth.GetValue("Customer_CODE"))
                            spdCallOffDtl.SetText(ENUM_DTL.CustomerName, spdCallOffDtl.MaxRows, rsUnAuth.GetValue("Customer_NAME"))
                            spdCallOffDtl.SetText(ENUM_DTL.ConsigneeCode, spdCallOffDtl.MaxRows, rsUnAuth.GetValue("Consignee_CODE"))
                            spdCallOffDtl.SetText(ENUM_DTL.ConsigneeName, spdCallOffDtl.MaxRows, rsUnAuth.GetValue("CONSIGNEE_NAME"))
                            spdCallOffDtl.Row = spdCallOffDtl.MaxRows
                            spdCallOffDtl.Col = ENUM_DTL.status
                            spdCallOffDtl.TypePictPicture = ImgYellow.Image
                            rsUnAuth.MoveNext()
                            spdCallOffDtl.MaxRows = spdCallOffDtl.MaxRows + 1
                        End While
                    End If
                    spdCallOffDtl.MaxRows = spdCallOffDtl.MaxRows - 1
                    rsUnAuth.ResultSetClose()
                    cmdGrpAuthorise.Focus()
                End If
            End If
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
End Class