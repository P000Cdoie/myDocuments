'--------------------------------------------------------------------------------------------------
'COPYRIGHT      :   MIND
'CREATED BY     :   MILIND MISHRA
'CREATED DATE   :   26 APR 2018
'SCREEN         :   Manual Picklist for FORD HMRS
'ISSUE ID       :   101505190 
'--------------------------------------------------------------------------------------------------
'MODIFIED BY    :   MILIND MISHRA
'MODIFIED DATE  :   16 MAY 2018
'ISSUE ID       :   101524167 
'--------------------------------------------------------------------------------------------------
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Linq
Imports System.Collections.Generic

Public Class frmMKTTRN0105

    Dim mintFormIndex As Integer


    Private Enum enmPickList
        ItemCode = 1
        ItemHlp
        Cust_drgNo
        Sch_date
        Sch_time
        Qty
        Dock_Code
        Container
        Plant_Code
    End Enum

#Region "Form level Events"
    Private Sub form_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub form_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Tag) = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub form_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            Me.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub

    Private Sub form_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        Dim keyascii As Short = Asc(e.KeyChar)
        Try
            If keyascii = 39 Then
                keyascii = 0
            End If
            e.KeyChar = Chr(keyascii)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Private Sub form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Call FitToClient(Me, GrpMain, ctlHeader, grpBtn, 500)
            Me.MdiParent = mdifrmMain
            mintFormIndex = mdifrmMain.AddFormNameToWindowList(Me.ctlHeader.Tag)
            SetGridsHeader()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub
#End Region

#Region "Methods"
    Private Sub SetGridsHeader()
        Try

            With sprPicklist
                .MaxRows = 0
                .MaxCols = [Enum].GetNames(GetType(enmPickList)).Count
                .Row = 0
                .set_RowHeight(0, 20)
                .Col = 0 : .set_ColWidth(0, 3)
                .Col = enmPickList.ItemCode : .Text = "Item Code" : .set_ColWidth(enmPickList.ItemCode, 16)
                .Col = enmPickList.ItemHlp : .Text = " " : .set_ColWidth(enmPickList.ItemHlp, 3)
                .Col = enmPickList.Cust_drgNo : .Text = "Drawing No." : .set_ColWidth(enmPickList.Cust_drgNo, 20)
                .Col = enmPickList.Sch_date : .Text = "Date" : .set_ColWidth(enmPickList.Sch_date, 10)
                .Col = enmPickList.Sch_time : .Text = "Time" : .set_ColWidth(enmPickList.Sch_time, 8)
                .Col = enmPickList.Qty : .Text = "Qty" : .set_ColWidth(enmPickList.Qty, 8)
                .Col = enmPickList.Dock_Code : .Text = "Dock Code" : .set_ColWidth(enmPickList.Dock_Code, 8)
                .Col = enmPickList.Container : .Text = "Container" : .set_ColWidth(enmPickList.Container, 8)
                .Col = enmPickList.Plant_Code : .Text = "Plant Code" : .set_ColWidth(enmPickList.Plant_Code, 10)
            End With


        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub AddBlankRow()
        Try
            With Me.sprPicklist
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .BackColorStyle = FPSpreadADO.BackColorStyleConstants.BackColorStyleUnderGrid
                .set_RowHeight(.Row, 15)
                .Col = enmPickList.ItemCode : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = enmPickList.ItemHlp : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonPicture = My.Resources.ico111.ToBitmap
                .Col = enmPickList.Cust_drgNo : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = enmPickList.Sch_date : .CellType = FPSpreadADO.CellTypeConstants.CellTypeDate
                .Col = enmPickList.Sch_time : .CellType = FPSpreadADO.CellTypeConstants.CellTypeTime
                .Col = enmPickList.Qty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                .Col = enmPickList.Dock_Code : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = enmPickList.Container : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = enmPickList.Plant_Code : .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
            End With

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function ValidateData() As Boolean
        Dim strItem As String
        Dim introw As Integer
        Dim dateToday As Date
        Dim dateTomorrow As Date
        Dim schDate As Date
        Try
            If Len(RTrim(LTrim(txtCustCode.Text))) = 0 Then
                MsgBox("Customer Code cannot be blank.", MsgBoxStyle.Exclamation, ResolveResString(100))
                Return False
            End If
            With sprPicklist
                .Row = .MaxRows
                If .Row = 0 Then
                    MsgBox("No row to generate picklist.", MsgBoxStyle.Information, ResolveResString(100))
                    Return False
                End If
                .Row = .MaxRows
                .Col = enmPickList.ItemCode : strItem = .Text
                If Len(RTrim(LTrim(strItem))) = 0 Then
                    MsgBox("Item Code cannot be blank.", MsgBoxStyle.Exclamation, ResolveResString(100))
                    .Col = enmPickList.ItemCode
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    .Focus()
                    Return False
                End If
                .Col = enmPickList.Cust_drgNo : strItem = .Text
                If Len(RTrim(LTrim(strItem))) = 0 Then
                    MsgBox("Drawing No. cannot be blank.", MsgBoxStyle.Exclamation, ResolveResString(100))
                    .Col = enmPickList.Cust_drgNo
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    .Focus()
                    Return False
                End If
                .Col = enmPickList.Sch_time : strItem = .Text
                If Len(RTrim(LTrim(strItem))) = 0 Then
                    MsgBox("Schedule Time cannot be blank.", MsgBoxStyle.Exclamation, ResolveResString(100))
                    .Col = enmPickList.Sch_time
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    .Focus()
                    Return False
                End If
                .Col = enmPickList.Qty : strItem = .Text
                If Len(RTrim(LTrim(strItem))) = 0 Then
                    MsgBox("Quantity cannot be blank.", MsgBoxStyle.Exclamation, ResolveResString(100))
                    .Col = enmPickList.Qty
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    .Focus()
                    Return False
                End If
                .Col = enmPickList.Plant_Code : strItem = .Text
                If Len(RTrim(LTrim(strItem))) = 0 Then
                    MsgBox("Warehouse Code cannot be blank.", MsgBoxStyle.Exclamation, ResolveResString(100))
                    .Col = enmPickList.Plant_Code
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    .Focus()
                    Return False
                End If
            End With
            dateToday = GetServerDate()
            dateTomorrow = DateAdd(DateInterval.Day, 1, GetServerDate)
            With sprPicklist
                For introw = 1 To .MaxRows
                    .Row = introw
                    .Col = enmPickList.Sch_date
                    schDate = .Text
                    If schDate > dateToday Then
                        If schDate > dateTomorrow Then
                            MsgBox("Schedule Date must be of Today and Tomorrow at row [" & introw & "].", MsgBoxStyle.Exclamation, ResolveResString(100))
                            Return False
                        End If
                    ElseIf schDate < dateToday Then
                        MsgBox("Schedule Date must be of Today and Tomorrow at row [" & introw & "].", MsgBoxStyle.Exclamation, ResolveResString(100))
                        Return False
                    End If
         
                Next
            End With
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
#End Region

    Private Sub cmdHelpCustCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelpCustCode.Click
        Dim strSql As String
        Dim strHelp() As String

        Try
            strSql = "SELECT * FROM GETFORDHMRSCUSTOMER WHERE UNIT_CODE='" & gStrUnitId & "'"
            strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql)
            If Not (UBound(strHelp) = -1) Then
                If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                    MsgBox("No Record Found !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                    Exit Sub
                Else
                    txtCustCode.Text = strHelp(0)
                    sprPicklist.MaxRows = 0
                    AddBlankRow()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub sprPicklist_ButtonClicked(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles sprPicklist.ButtonClicked
        Dim strSql As String
        Dim strHelp() As String
        Dim oRDR As SqlDataReader
        Dim strWHCode As String
        Dim strWHCode1 As String
        Try
            With sprPicklist
                If e.col = enmPickList.ItemHlp Then
                    strSql = "SELECT DISTINCT ITEM_CODE, ITEM_DESC, CUST_DRGNO ,DRG_DESC,DOCKCODE,CONTAINER   FROM CUSTITEM_MST "
                    strSql = strSql & " WHERE UNIT_CODE='" & gStrUnitId & "' AND ACCOUNT_CODE='" & txtCustCode.Text & "' AND DOCKCODE<>'' AND CONTAINER<>'' AND DOCKCODE IS NOT NULL "
                    strSql = strSql & " AND CONTAINER IS NOT NULL"
                    strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql)
                    If Not (UBound(strHelp) = -1) Then
                        If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                            MsgBox("No Record Found !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                            Exit Sub
                        Else
                            .Row = .ActiveRow
                            .Col = enmPickList.ItemCode : .Text = strHelp(0)
                            .Col = enmPickList.Cust_drgNo : .Text = strHelp(2)
                            .Col = enmPickList.Sch_date : .Text = GetServerDate()
                            .Col = enmPickList.Dock_Code : .Text = strHelp(4)
                            .Col = enmPickList.Container : .Text = strHelp(5)
                            '  strSql = "select WH_CODE from vw_getfordhmrs_whcode where unit_code='" & gStrUnitId & "'"
                            strSql = "SELECT PLANT_CODE FROM CUSTOMER_MST WHERE CUSTOMER_CODE ='" & txtCustCode.Text & "' AND UNIT_CODE='" & gStrUnitId & "'"
                            oRDR = SqlConnectionclass.ExecuteReader(strSql)
                            While (oRDR.Read)
                                strWHCode = oRDR("PLANT_CODE").ToString()
                                strWHCode1 = "" & strWHCode1 + " " + strWHCode & "" + Chr(9)
                                strWHCode = ""
                            End While
                            .Col = enmPickList.Plant_Code : .TypeComboBoxList = strWHCode1
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try
    End Sub

    Private Sub sprPicklist_DblClick(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_DblClickEvent) Handles sprPicklist.DblClick
        Dim msgReturnVal As MsgBoxResult
        Try
            If e.col = 0 And e.row > 0 Then
                With Me.sprPicklist
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = 1
                    .Row2 = .MaxRows
                    msgReturnVal = MsgBox("Do you want to delete the entire row?", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, ResolveResString(100))
                    If msgReturnVal = MsgBoxResult.Yes Then
                        .Row = .ActiveRow
                        .Action = FPSpreadADO.ActionConstants.ActionDeleteRow : .MaxRows = .MaxRows - 1
                    End If
                    If .MaxRows = 0 Then
                        txtCustCode.Text = ""
                    End If
                End With
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub sprPicklist_KeyUpEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyUpEvent) Handles sprPicklist.KeyUpEvent
        Try
            If e.keyCode = Keys.N And e.shift = 2 Then ' To add blank row Ctrl + N
                If txtCustCode.Text.Trim = "" Then Exit Sub
                With sprPicklist
                    If .MaxRows > 0 Then
                        .Row = .MaxRows : .Col = enmPickList.ItemCode
                        If .Text.Trim = "" Then Return
                    End If
                    Call AddBlankRow()
                    .Row = .MaxRows
                    .Col = enmPickList.ItemHlp
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    .Focus()
                    If sprPicklist.ActiveCol = enmPickList.Sch_date Then
                        If sprPicklist.MaxRows = sprPicklist.ActiveRow Then
                            Call AddBlankRow()
                            .Col = enmPickList.ItemHlp
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                        End If
                    End If
                End With
            Else
                If sprPicklist.ActiveCol = enmPickList.ItemCode AndAlso e.keyCode = Keys.F1 Then
                    sprPicklist_ButtonClicked(sprPicklist, New AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent(enmPickList.ItemHlp, sprPicklist.ActiveRow, 1))
                End If
            End If
            Exit Sub

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        Try
            txtCustCode.Text = ""
            sprPicklist.MaxRows = 0
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnGeneratePicklist_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGeneratePicklist.Click
        Dim introw As Integer
        Dim strSql As String
        Dim strItem As String
        Dim strDrgno As String
        Dim schdate As String
        Dim schtime As String
        Dim Qty As String
        Dim dockcode As String
        Dim container As String
        Dim wh_code As String
        Dim strPicklist As String
        Dim oCMD As SqlCommand
        Try
            If ValidateData() = True Then
                oCMD = New SqlCommand
                oCMD.Connection = SqlConnectionclass.GetConnection
                oCMD.Transaction = oCMD.Connection.BeginTransaction
                strSql = "DELETE FROM TMP_PICKLIST_FORD_HMRS WHERE IP_ADDRESS='" & gstrIpaddressWinSck & "' AND UNIT_CODE='" & gStrUnitId & "'"
                SqlConnectionclass.ExecuteNonQuery(strSql)
                With sprPicklist
                    For introw = 1 To .MaxRows
                        .Row = introw
                        .Col = enmPickList.ItemCode
                        strItem = .Text
                        .Col = enmPickList.Cust_drgNo
                        strDrgno = .Text
                        .Col = enmPickList.Sch_date
                        schdate = .Text
                        .Col = enmPickList.Sch_time
                        schtime = .Text
                        .Col = enmPickList.Qty
                        Qty = .Text
                        .Col = enmPickList.Dock_Code
                        dockcode = .Text
                        .Col = enmPickList.Container
                        container = .Text
                        .Col = enmPickList.Plant_Code
                        wh_code = .Text
                        strSql = " SET DATEFORMAT 'DMY' insert into TMP_PICKLIST_FORD_HMRS "
                        strSql = strSql & "([CUSTOMER_CODE],[ITEM_CODE],[CUST_DRGNO],[SCH_DATE],[SCH_TIME],[QTY],[DOCKCODE],[CONTAINER],[WH_CODE],[IP_ADDRESS],[UNIT_CODE])"
                        strSql = strSql & "values("
                        strSql = strSql & "'" & txtCustCode.Text & "','" & strItem & "','" & strDrgno & "','" & schdate & "','" & schtime & "'," & Qty & ",'" & dockcode & "' , "
                        strSql = strSql & " '" & container & "', '" & wh_code & "', '" & gstrIpaddressWinSck & "','" & gStrUnitId & "')"
                        SqlConnectionclass.ExecuteNonQuery(strSql)
                    Next


                    With oCMD
                        .CommandType = CommandType.StoredProcedure
                        .CommandText = "USP_GENERATEMANUALFORDHMRSPICKLIST"
                        .CommandTimeout = 0
                        .Parameters.Clear()
                        .Parameters.AddWithValue("@UNIT_CODE", gStrUnitId)
                        .Parameters.AddWithValue("@IP_ADDRESS", gstrIpaddressWinSck)
                        '     .Parameters.AddWithValue("@PICKLISTNO", "").Direction = ParameterDirection.Output
                        .Parameters.Add("@PICKLISTNOSTRING", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                        .ExecuteNonQuery()
                        strPicklist = .Parameters("@PICKLISTNOSTRING").Value.ToString
                        If Len(RTrim(LTrim(strPicklist))) <> 0 Then
                            strPicklist = strPicklist.Remove(strPicklist.Length - 1)
                        End If
                    End With
                End With
                oCMD.Transaction.Commit()
                If Len(RTrim(LTrim(strPicklist))) <> 0 Then
                    MsgBox("Picklist No. " + Chr(13) + strPicklist + Chr(13) + " generated sucessfully.", MsgBoxStyle.Information, ResolveResString(100))
                Else
                    MsgBox("Picklist already exists with this data.", MsgBoxStyle.Information, ResolveResString(100))
                End If
                sprPicklist.MaxRows = 0
                txtCustCode.Text = ""
            End If
        Catch ex As Exception
            oCMD.Transaction.Rollback()
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class





