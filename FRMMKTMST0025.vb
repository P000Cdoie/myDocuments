'--------------------------------------------
'CREATED BY : SHABBIR HUSSAIN
'CREATED ON : 06 JAN 2011
'FORM DESC  : FORD Trigger mapping
'--------------------------------------------
'Modified by Roshan Singh on 31 JAN 2012 for multiunit functionality
Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Public Class FrmMKTMST0025
#Region "Form level variable Declarations"
    Dim mintFormtag As Short
    Dim mstrIPAddress As String
    Dim mctlError As System.Windows.Forms.Control
    Dim blnIsPopulating As Boolean
    Private Enum enumsspr
        ModelCode = 1
        CategoryCode
        BodyColor
        LinkItems
        MappingStatus
        Active
    End Enum
#End Region

#Region "Form Events"
    Private Sub FrmMKTMST0025_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        Try
            mdifrmMain.CheckFormName = mintFormtag
            frmModules.NodeFontBold(Me.Tag) = True
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub FrmMKTMST0025_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate

        Try
            frmModules.NodeFontBold(Me.Tag) = False
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub FrmMKTMST0025_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Try
            mstrIPAddress = gstrIpaddressWinSck
            Call FitToClient(Me, fraMain, ctlFormHeader1, GrpButtons, 550)
            Call EnableControls(False, Me, True)
            MdiParent = prjMPower.mdifrmMain
            gblnCancelUnload = False : gblnFormAddEdit = False
            mintFormtag = mdifrmMain.AddFormNameToWindowList(Me.ctlFormHeader1.Tag)
            fraMain.Enabled = True
            sspr.Enabled = True
            SetSpreadProperty()
            GrpButtons.Revert()
            CmdPopulateMapping.Enabled = True
            CmdHlpModel.Enabled = True
            CmdHlpCatCode.Enabled = True
            ctlFormHeader1.HeaderString = "FORD Trigger Mapping"
            Me.Text = "MKTMST0025 - FORD Trigger Mapping"
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub FrmMKTMST0025_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000

        Try
            If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
                Call ctlFormHeader1_Click(ctlFormHeader1, New System.EventArgs())
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub FrmMKTMST0025_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Try
            Select Case KeyAscii
                Case System.Windows.Forms.Keys.Escape
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        GrpButtons.Revert()
                        SetSpreadProperty()
                        fraMain.Enabled = True
                        CmdPopulateMapping.Enabled = True
                        CmdHlpModel.Enabled = True
                        CmdHlpCatCode.Enabled = True
                    Else
                        Me.ActiveControl.Focus()
                    End If
            End Select
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            e.KeyChar = Chr(KeyAscii)
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        End Try

    End Sub

    Private Sub FrmMKTMST0025_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Try
            Me.Dispose()
            mdifrmMain.RemoveFormNameFromWindowList = mintFormtag
            frmModules.NodeFontBold(Me.Tag) = False
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub
#End Region

#Region "Routines"
    Private Sub PopulateMapping()

        Dim oSqlCmd As New SqlCommand
        Dim oDr As SqlDataReader = Nothing
        Dim strSQL As String = ""
        Dim strCondition As String = ""

        Try
            blnIsPopulating = True
            SetSpreadProperty()

            With oSqlCmd
                .Connection = SqlConnectionclass.GetConnection
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .CommandText = "USP_POPULATEFORDTRIGGER_MAPPING"
                .Parameters.Add("@Unit_Code", SqlDbType.VarChar, 10).Value = gstrUNITID
                .Parameters.Add("@MODEL_CODE", SqlDbType.VarChar, 50).Value = lblModel.Text.Trim
                .Parameters.Add("@CATEGORY_CODE", SqlDbType.VarChar, 15).Value = lblCatCode.Text.Trim
                .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 50).Value = gstrIpaddressWinSck
                oDr = .ExecuteReader()
            End With

            If oDr.HasRows = True Then
                With sspr
                    While oDr.Read
                        addRowAtEnterKeyPress(1)
                        .Row = .MaxRows
                        .Col = enumsspr.ModelCode
                        .Value = oDr("MODEL_CODE").ToString()
                        .Col = enumsspr.CategoryCode
                        .Text = oDr("CATEGORY_CODE").ToString()
                        .Col = enumsspr.BodyColor
                        .Text = oDr("BODY_COLOR").ToString()
                        .Col = enumsspr.Active
                        .Value = Val(oDr("IsActive"))
                        If Val(oDr("IsActive")) = 0 Then
                            .Col = 1
                            .Col2 = .MaxCols
                            .Row = .MaxRows
                            .Row2 = .MaxRows
                            .BlockMode = True
                            .BackColor = Me.BackColor
                            .BlockMode = False
                        Else
                            If Val(oDr("STATUS")) = 0 Then
                                .Col = enumsspr.LinkItems
                                .Col2 = enumsspr.LinkItems
                                .Row = .MaxRows
                                .Row2 = .MaxRows
                                .BlockMode = True
                                .FontSize = 12
                                .TypeButtonColor = Color.Crimson
                                .TypeButtonTextColor = Color.White
                                .FontBold = True
                                .BlockMode = False
                            Else
                                .Col = enumsspr.LinkItems
                                .Col2 = enumsspr.LinkItems
                                .Row = .MaxRows
                                .Row2 = .MaxRows
                                .BlockMode = True
                                .FontSize = 12
                                .TypeButtonColor = Color.DarkGreen
                                .TypeButtonTextColor = Color.White
                                .FontBold = True
                                .BlockMode = False
                            End If

                            .Col = 1
                            .Col2 = .MaxCols
                            .Row = .MaxRows
                            .Row2 = .MaxRows
                            .BlockMode = True
                            .FontSize = 12
                            .BackColor = Color.White
                            .BlockMode = False
                        End If
                    End While
                End With
            End If
            If oDr.IsClosed = False Then oDr.Close()
            oSqlCmd.Connection.Close()
            oSqlCmd.Dispose()
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            blnIsPopulating = False
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Function SaveMapping() As Boolean

        Dim oSqlCn As SqlConnection = Nothing
        Dim oSqlTrans As SqlTransaction = Nothing
        Dim oSqlCmd As New SqlCommand()
        Dim strSql As String = ""
        Dim strModel As String = ""
        Dim strCatCode As String = ""
        Dim strBodyColor As String = ""
        Dim strStatus As String = ""
        Dim LngCount As Long = 0

        SaveMapping = False

        Try
            oSqlCn = New SqlConnection
            oSqlCn = SqlConnectionclass.GetConnection()
            oSqlTrans = oSqlCn.BeginTransaction

            With oSqlCmd
                .Connection = oSqlCn
                .Transaction = oSqlTrans
                .CommandText = "USP_SAVE_TRIGGER_ITEM_MAPPING"
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .Parameters.Add("@Unit_Code", SqlDbType.VarChar, 10).Value = gstrUNITID
                .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 50).Value = gstrIpaddressWinSck
                .Parameters.Add("@ERRMSG", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                .ExecuteNonQuery()

                If .Parameters(.Parameters.Count - 1).Value.ToString().Trim.Length <> 0 Then
                    MessageBox.Show("Error while saving Trigger Item Mapping details : " & vbCrLf & .Parameters(.Parameters.Count - 1).Value.ToString(), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
                    oSqlTrans.Rollback()
                    oSqlCmd = Nothing
                    oSqlCn.Close()
                    oSqlCn = Nothing
                    oSqlTrans = Nothing
                    Exit Function
                End If
            End With

            With sspr
                For LngCount = 1 To .MaxRows
                    .Row = LngCount
                    .Col = enumsspr.ModelCode
                    strModel = .Text.Trim.ToString.ToUpper

                    .Col = enumsspr.CategoryCode
                    strCatCode = .Text.Trim.ToString.ToUpper

                    .Col = enumsspr.BodyColor
                    strBodyColor = .Text.Trim.ToString.ToUpper

                    .Col = enumsspr.Active
                    strSql = "UPDATE TRIGGER_ITEM_MAPPING SET ISACTIVE=" & .Value & " "
                    strSql = strSql & "WHERE Unit_Code='" & gstrUNITID & "' And CATEGORY_CODE='" & strCatCode.Trim & "' AND MODEL_CODE='" & strModel.Trim & "' AND BODY_COLOR='" & strBodyColor.Trim & "'"
                    SqlConnectionclass.ExecuteNonQuery(oSqlTrans, CommandType.Text, strSql)
                Next
            End With

            oSqlTrans.Commit()
            oSqlCmd = Nothing
            If oSqlCn.State = ConnectionState.Open Then oSqlCn.Close()
            oSqlCn = Nothing
            oSqlTrans = Nothing
            SaveMapping = True
        Catch ex As Exception
            oSqlTrans.Rollback()
            oSqlCmd = Nothing
            If oSqlCn.State = ConnectionState.Open Then oSqlCn.Close()
            oSqlCn = Nothing
            oSqlTrans = Nothing
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Function
#End Region

#Region "Grid Events and Routines"
    Private Sub SetSpreadProperty()

        Try
            With sspr
                .MaxRows = 0
                .MaxCols = 0
                .MaxCols = enumsspr.Active
                .UserColAction = FPSpreadADO.UserColActionConstants.UserColActionSort
                .Row = 0
                .Col = enumsspr.ModelCode : .Text = "Model Code" : .set_ColWidth(enumsspr.ModelCode, 2400)
                .Col = enumsspr.CategoryCode : .Text = "Cat Codes" : .set_ColWidth(enumsspr.CategoryCode, 2400)
                .Col = enumsspr.BodyColor : .Text = "Body Color" : .set_ColWidth(enumsspr.BodyColor, 2400)
                .Col = enumsspr.LinkItems : .Text = "Link Items" : .set_ColWidth(enumsspr.LinkItems, 2000)
                .Col = enumsspr.MappingStatus : .Text = "Mapping Exists" : .set_ColWidth(enumsspr.MappingStatus, 400) : .ColHidden = True
                .Col = enumsspr.Active : .Text = "Active" : .set_ColWidth(enumsspr.Active, 1200)
                .set_RowHeight(0, 400)
            End With
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            sspr.Visible = True
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub SetSpreadColTypes(ByVal pintRowNo As Integer)

        Try
            With sspr
                .Row = pintRowNo
                .Col = enumsspr.ModelCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = True
                .Col = enumsspr.CategoryCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = True
                .Col = enumsspr.BodyColor : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = True
                .Col = enumsspr.LinkItems : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = "Link Items" : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = False
                .Col = enumsspr.MappingStatus : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Lock = True
                .Col = enumsspr.Active : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .TypeCheckCenter = True : .Lock = True
            End With
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub addRowAtEnterKeyPress(ByVal pintRows As Integer)

        Dim intRowHeight As Integer

        Try
            With sspr
                .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
                For intRowHeight = 1 To pintRows
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .set_RowHeight(.Row, 330)
                    Call SetSpreadColTypes(.Row)
                Next intRowHeight
            End With
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub sspr_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles sspr.ButtonClicked

        Dim frmLinkTriggerVsItem As frmLinkTriggerVsItem = Nothing
        Dim strModelCode As String, strCatCode As String, strBodyColor As String

        Try
            If blnIsPopulating = True Then Exit Sub

            If e.col = enumsspr.Active Then
                With sspr
                    .Col = e.col
                    .Row = e.row
                    If .Value = 0 Then
                        .Col = 1
                        .Col2 = .MaxCols
                        .Row = e.row
                        .Row2 = e.row
                        .BlockMode = True
                        .BackColor = Me.BackColor
                        .BlockMode = False
                    Else
                        .Col = 1
                        .Col2 = .MaxCols
                        .Row = e.row
                        .Row2 = e.row
                        .BlockMode = True
                        .BackColor = Color.White
                        .BlockMode = False
                    End If
                End With
            ElseIf e.col = enumsspr.LinkItems Then
                With sspr
                    .Row = e.row
                    .Col = enumsspr.ModelCode
                    strModelCode = .Text.Trim

                    .Col = enumsspr.CategoryCode
                    strCatCode = .Text.Trim

                    .Col = enumsspr.BodyColor
                    strBodyColor = .Text.Trim
                End With

                frmLinkTriggerVsItem = New frmLinkTriggerVsItem()
                With frmLinkTriggerVsItem
                    .mstrModelCode = strModelCode
                    .mstrCatCode = strCatCode
                    .mstrBodyColor = strBodyColor

                    If GrpButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                        .mstrMode = "NEW"
                    ElseIf GrpButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                        .mstrMode = "EDIT"
                    Else
                        .mstrMode = "VIEW"
                    End If

                    .StartPosition = FormStartPosition.CenterParent
                    .ssprItem.CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
                    .ssprItem.MaxRows = 0
                    .PopulateLinkDtlOfModelCatCodeBodyColor()
                    .ShowDialog()
                End With
                frmLinkTriggerVsItem = Nothing
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub
#End Region

#Region "Other Controls Events"
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        Try
            Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
#End Region

#Region "Button Controls"
    Private Sub GrpButtons_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtnEditGrp.ButtonClickEventArgs) Handles GrpButtons.ButtonClick

        Try
            Select Case e.Button
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                    If sspr.MaxRows = 0 Then
                        MessageBox.Show("No data to map !" & vbCr & "First populate the data", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        If CmdPopulateMapping.Enabled = True Then CmdPopulateMapping.Focus()
                        Exit Sub
                    End If

                    CmdPopulateMapping.Enabled = False
                    CmdHlpModel.Enabled = False
                    CmdHlpCatCode.Enabled = False

                    With sspr
                        .Col = enumsspr.Active
                        .Col2 = enumsspr.Active
                        .Row = 1
                        .Row2 = .MaxRows
                        .BlockMode = True
                        .Lock = False
                        .BlockMode = False
                    End With
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                    Select Case GrpButtons.Mode
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                            If sspr.MaxRows = 0 Then
                                MessageBox.Show("No record found to save !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                                gblnCancelUnload = True : gblnFormAddEdit = True
                                Exit Sub
                            End If

                            If SaveMapping() = False Then
                                Exit Sub
                            End If

                            gblnCancelUnload = False : gblnFormAddEdit = False
                            Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            Call EnableControls(False, Me)
                            GrpButtons.Revert()
                            SetSpreadProperty()
                            lblModel.Text = ""
                            lblCatCode.Text = ""
                            fraMain.Enabled = True
                            CmdPopulateMapping.Enabled = True
                            CmdHlpModel.Enabled = True
                            CmdHlpCatCode.Enabled = True
                    End Select
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                    Call FrmMKTMST0025_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                    Me.Close()
            End Select
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            gblnCancelUnload = True : gblnFormAddEdit = True
        End Try

    End Sub

    Private Sub CmdPopulateMapping_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdPopulateMapping.Click

        Try
            PopulateMapping()
            fraMain.Enabled = True
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub CmdHlpModel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdHlpModel.Click

        Dim strHlp() As String
        Dim strSQL As String = ""

        Try
            strSQL = "SELECT MODEL_CODE,MODEL_CODE FROM DAY_WISE_FORD_TRIGGER_DTL  "
            strSQL = strSQL + "WHERE unit_Code = '" & gstrUNITID & "' And TRGDATE>=GETDATE()-(SELECT TRGMAPPINGINTERVAL FROM PRODUCTIONCONF where unit_Code = '" & gstrUNITID & "') "
            If lblCatCode.Text.Trim.Length > 0 Then
                strSQL = strSQL + " AND CATEGORY_CODE='" & lblCatCode.Text.Trim & "' "
            End If
            strSQL = strSQL + "GROUP BY MODEL_CODE"

            strHlp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "FORD Trigger Models Help", 1)
            If Not (UBound(strHlp) <= 0) Then
                If Not (UBound(strHlp) = 0) Then
                    If (Len(strHlp(0)) >= 1) And strHlp(0) = "0" Then
                        MsgBox("No Trigger Details To Display.", MsgBoxStyle.Information, ResolveResString(100))
                        lblModel.Text = ""
                        Exit Sub
                    Else
                        lblModel.Text = strHlp(0)
                        CmdPopulateMapping.PerformClick()
                    End If
                End If
            Else
                lblModel.Text = ""
                CmdPopulateMapping.PerformClick()
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub CmdHlpCatCode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdHlpCatCode.Click

        Dim strHlp() As String
        Dim strSQL As String = ""

        Try
            strSQL = "SELECT CATEGORY_CODE,CATEGORY_CODE FROM DAY_WISE_FORD_TRIGGER_DTL  "
            strSQL = strSQL + "WHERE unit_Code = '" & gstrUNITID & "' And TRGDATE>=GETDATE()-(SELECT TRGMAPPINGINTERVAL FROM PRODUCTIONCONF where Unit_Code = '" & gstrUNITID & "') "
            If lblModel.Text.Trim.Length > 0 Then
                strSQL = strSQL + " AND MODEL_CODE='" & lblModel.Text.Trim & "' "
            End If
            strSQL = strSQL + "GROUP BY CATEGORY_CODE"

            strHlp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "FORD Trigger Cat Codes Help", 1)
            If Not (UBound(strHlp) <= 0) Then
                If Not (UBound(strHlp) = 0) Then
                    If (Len(strHlp(0)) >= 1) And strHlp(0) = "0" Then
                        MsgBox("No Trigger Details To Display.", MsgBoxStyle.Information, ResolveResString(100))
                        lblCatCode.Text = ""
                        Exit Sub
                    Else
                        lblCatCode.Text = strHlp(0)
                        CmdPopulateMapping.PerformClick()
                    End If
                End If
            Else
                lblCatCode.Text = ""
                CmdPopulateMapping.PerformClick()
            End If
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub
#End Region

End Class