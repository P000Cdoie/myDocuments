Imports System.Data.SqlClient
'*********************************************************************************************************************
'Copyright(c)       - MIND
'Name of Module     - BSR FIFO Authorization
'Name of Form       - FRMMKTTRN0110  , BSR FIFO Authorization
'Created by         - Ashish sharma
'Created Date       - 29 AUG 2019
'description        - BSR FIFO Authorization (New Development)
'*********************************************************************************************************************
Public Class FRMMKTTRN0110
    Private Enum enumBarcodeGrid
        TICK = 1
        SNO
        ITEM_CODE
        BARCODE
        FIFO_DATE
        REQUEST_USER_ID
        REMARKS
    End Enum
    Private Sub FRMMKTTRN0110_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        txtItemCode.Focus()
    End Sub
    Private Sub FRMMKTTRN0110_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Call FitToClient(Me, GrpMain, ctlHeader, GrpBoxButtons, 600)
            Me.MdiParent = mdifrmMain
            InitializeSpread()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Try
            FillBarcodes()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            If MessageBox.Show("Are you sure you want close?", "eMPro", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = Windows.Forms.DialogResult.Yes Then
                Me.Close()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            txtItemCode.Text = String.Empty
            lblItemName.Text = String.Empty
            fspBarcode.MaxRows = 0
            rdbCheckAll.Checked = False
            rdbUncheckAll.Checked = False
            txtItemCode.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub FillBarcodes()
        Dim dt As New DataTable
        Dim dtBarcodes As New DataTable
        Try
            Dim sqlCmd As New SqlCommand
            dtBarcodes.Columns.Add("ID", GetType(System.Int64))
            dtBarcodes.Columns.Add("BARCODE", GetType(System.String))
            dtBarcodes.Columns.Add("REMARKS", GetType(System.String))
            With sqlCmd
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .CommandText = "USP_BSR_SKIP_FIFO_AUTH_REJECT"
                .Parameters.Clear()
                .Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
                .Parameters.AddWithValue("@USER_ID", mP_User)
                .Parameters.AddWithValue("@UDT_BSR_SKIP_FIFO_AUTH_REJECT", dtBarcodes)
                If Not String.IsNullOrEmpty(txtItemCode.Text.Trim()) Then
                    .Parameters.AddWithValue("@ITEM_CODE", txtItemCode.Text.Trim())
                End If
                .Parameters.AddWithValue("@OPERATION", "SELECT")
                .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                dt = SqlConnectionclass.GetDataTable(sqlCmd)
                If Convert.ToString(.Parameters("@MESSAGE").Value) <> "" Then
                    MsgBox(Convert.ToString(.Parameters("@MESSAGE").Value), MsgBoxStyle.Exclamation, ResolveResString(100))
                Else
                    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                        InitializeSpread()
                        For i As Integer = 0 To dt.Rows.Count - 1
                            With fspBarcode
                                AddRow()
                                .SetText(enumBarcodeGrid.TICK, i + 1, False)
                                .SetText(enumBarcodeGrid.SNO, i + 1, dt.Rows(i).Item("ID").ToString())
                                .SetText(enumBarcodeGrid.ITEM_CODE, i + 1, dt.Rows(i).Item("ITEM_CODE"))
                                .SetText(enumBarcodeGrid.BARCODE, i + 1, dt.Rows(i).Item("BARCODE"))
                                .SetText(enumBarcodeGrid.FIFO_DATE, i + 1, dt.Rows(i).Item("FIFO_DATE"))
                                .SetText(enumBarcodeGrid.REQUEST_USER_ID, i + 1, dt.Rows(i).Item("ENT_BY"))
                            End With
                        Next
                    End If
                End If
            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            dt.Dispose()
            dtBarcodes.Dispose()
        End Try
    End Sub

    Private Sub InitializeSpread()
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            With Me.fspBarcode
                .MaxRows = 0
                .MaxCols = [Enum].GetValues(GetType(enumBarcodeGrid)).Length
                .set_RowHeight(0, 20)
                .Row = 0 : .Col = enumBarcodeGrid.TICK : .Text = "Select" : .set_ColWidth(.Col, 5) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBarcodeGrid.SNO : .Text = "Id" : .set_ColWidth(.Col, 5) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .ColHidden = True
                .Row = 0 : .Col = enumBarcodeGrid.ITEM_CODE : .Text = "Item Code" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBarcodeGrid.BARCODE : .Text = "Barcode" : .set_ColWidth(.Col, 32) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBarcodeGrid.FIFO_DATE : .Text = "FIFO Date" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBarcodeGrid.REQUEST_USER_ID : .Text = "Requestor UserId" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = enumBarcodeGrid.REMARKS : .Text = "Remarks [F1]" : .set_ColWidth(.Col, 22) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            End With
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Public Sub AddRow()
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            With fspBarcode
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows : .Col = enumBarcodeGrid.TICK : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumBarcodeGrid.SNO : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumBarcodeGrid.ITEM_CODE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumBarcodeGrid.BARCODE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumBarcodeGrid.FIFO_DATE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumBarcodeGrid.REQUEST_USER_ID : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = enumBarcodeGrid.REMARKS : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
            End With
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub AuthorizeReject(ByVal strOperation As String)
        Dim dtBarcodes As New DataTable
        Try
            dtBarcodes.Columns.Add("ID", GetType(System.Int64))
            dtBarcodes.Columns.Add("BARCODE", GetType(System.String))
            dtBarcodes.Columns.Add("REMARKS", GetType(System.String))
            Dim drBarcode As DataRow
            Dim isChecked As String = String.Empty
            With fspBarcode
                For i As Integer = 1 To .MaxRows
                    .Row = i
                    .Col = enumBarcodeGrid.TICK
                    isChecked = Convert.ToString(.Value)
                    If isChecked = "1" Then
                        drBarcode = dtBarcodes.NewRow()
                        .Row = i
                        .Col = enumBarcodeGrid.SNO
                        drBarcode("ID") = Convert.ToInt64(.Text)
                        .Row = i
                        .Col = enumBarcodeGrid.BARCODE
                        drBarcode("BARCODE") = .Text.Trim()
                        .Row = i
                        .Col = enumBarcodeGrid.REMARKS
                        drBarcode("REMARKS") = .Text.Trim()
                        dtBarcodes.Rows.Add(drBarcode)
                    End If
                Next
            End With
            If dtBarcodes IsNot Nothing AndAlso dtBarcodes.Rows.Count > 0 Then
                Dim sqlCmd As New SqlCommand
                With sqlCmd
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 0
                    .CommandText = "USP_BSR_SKIP_FIFO_AUTH_REJECT"
                    .Parameters.Clear()
                    .Parameters.AddWithValue("@UNIT_CODE", gstrUnitId)
                    .Parameters.AddWithValue("@USER_ID", mP_User)
                    .Parameters.AddWithValue("@UDT_BSR_SKIP_FIFO_AUTH_REJECT", dtBarcodes)
                    If strOperation = "A" Then
                        .Parameters.AddWithValue("@OPERATION", "AUTHORIZE")
                    Else
                        .Parameters.AddWithValue("@OPERATION", "REJECT")
                    End If
                    .Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                    SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                    If Convert.ToString(.Parameters("@MESSAGE").Value) <> "" Then
                        MsgBox(Convert.ToString(.Parameters("@MESSAGE").Value), MsgBoxStyle.Exclamation, ResolveResString(100))
                    Else
                        If strOperation = "A" Then
                            MsgBox("Selected Barcode(s) authorized successfully.")
                        Else
                            MsgBox("Selected Barcode(s) rejected successfully.")
                        End If
                        btnCancel.PerformClick()
                    End If
                End With
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            dtBarcodes.Dispose()
        End Try
    End Sub

    Private Function ValidateRows(ByVal strOperation As String) As Boolean
        Dim result As Boolean = False
        Dim chkSelect As String = String.Empty
        Dim countCheck As Integer = 0
        Try
            If fspBarcode Is Nothing OrElse fspBarcode.MaxRows = 0 Then
                MessageBox.Show("No barcode for authorize/reject.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                result = True
                Return result
            End If
            With fspBarcode
                For i As Integer = 1 To .MaxRows
                    .Row = i
                    .Col = enumBarcodeGrid.TICK
                    chkSelect = Convert.ToString(.Value)
                    If chkSelect = "1" Then
                        countCheck += 1
                        Exit For
                    End If
                Next
                If countCheck = 0 Then
                    MessageBox.Show("Please select atleast one barcode for authorize/reject.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    result = True
                    Return result
                End If

                For i As Integer = 1 To .MaxRows
                    .Row = i
                    .Col = enumBarcodeGrid.TICK
                    chkSelect = Convert.ToString(.Value)
                    If chkSelect = "1" Then
                        .Row = i
                        .Col = enumBarcodeGrid.REMARKS
                        If String.IsNullOrEmpty(.Text.Trim()) Then
                            MessageBox.Show("Please select [Remarks].", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            result = True
                            Return result
                        End If
                    End If
                Next
                For i As Integer = 1 To .MaxRows
                    .Row = i
                    .Col = enumBarcodeGrid.TICK
                    chkSelect = Convert.ToString(.Value)
                    If chkSelect = "1" Then
                        .Row = i
                        .Col = enumBarcodeGrid.REMARKS
                        If strOperation = "A" AndAlso .Text.Trim().ToUpper() = "TAG AVAILABLE NOW" Then
                            MessageBox.Show("Kindly select valid [Remarks] for authorization.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            result = True
                            Return result
                        ElseIf strOperation = "R" AndAlso .Text.Trim().ToUpper() <> "TAG AVAILABLE NOW" Then
                            MessageBox.Show("Kindly select valid [Remarks] for rejection.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            result = True
                            Return result
                        End If
                    End If
                Next
            End With
            Return result
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function

    Private Sub rdbCheckAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbCheckAll.CheckedChanged
        Try
            CheckUncheckAll("1")
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub rdbUncheckAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbUncheckAll.CheckedChanged
        Try
            CheckUncheckAll("0")
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub CheckUncheckAll(ByVal tick As String)
        If fspBarcode Is Nothing OrElse fspBarcode.MaxRows = 0 Then
            rdbCheckAll.Checked = False
            rdbUncheckAll.Checked = False
            Exit Sub
        End If
        With fspBarcode
            For i As Integer = 1 To .MaxRows
                .Row = i
                .Col = enumBarcodeGrid.TICK
                .Value = tick
            Next
        End With
    End Sub

    Private Sub cmdItemCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdItemCode.Click
        Dim strHelp() As String = Nothing
        Dim strSql As String = String.Empty
        Try
            strSql = "SELECT ITEM_CODE,ITEM_NAME FROM DBO.UDF_BSR_FIFO_AUTH_ITEM_HELP('" & gstrUnitId & "') ORDER BY ITEM_CODE"
            strHelp = Me.ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "Customer(s) Help")
            If Not (UBound(strHelp) <= 0) Then
                If (Len(strHelp(0)) >= 1) And strHelp(0) = "0" Then
                    MessageBox.Show("No record To Display", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    txtItemCode.Text = String.Empty
                    Exit Sub
                Else
                    txtItemCode.Text = strHelp(0).Trim
                    lblItemName.Text = strHelp(1).Trim
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtItemCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtItemCode.KeyDown
        Try
            Dim objText As TextBox = DirectCast(sender, TextBox)
            If e.KeyCode = Keys.Delete Then
                objText.Text = String.Empty
            ElseIf e.KeyCode = Keys.F1 Then
                cmdItemCode_Click(sender, e)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtItemCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtItemCode.TextChanged
        Try
            lblItemName.Text = String.Empty
            fspBarcode.MaxRows = 0
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub fspBarcode_KeyDownEvent(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles fspBarcode.KeyDownEvent
        Dim strQuery As String
        Dim strHelp() As String
        Dim intLoopCounter As Integer = 0
        Dim isChecked As String = String.Empty
        Try
            If e.keyCode = Keys.F1 Then
                fspBarcode.Row = fspBarcode.ActiveRow
                fspBarcode.Col = enumBarcodeGrid.TICK
                isChecked = Convert.ToString(fspBarcode.Value)
                If isChecked = "0" Then Exit Sub
                If fspBarcode.ActiveCol = enumBarcodeGrid.REMARKS Then
                    With Me.fspBarcode
                        intLoopCounter = Me.fspBarcode.ActiveRow
                        strQuery = "SELECT DESCR REMARK_CODE,DESCR REMARKS FROM LISTS (NOLOCK) WHERE KEY1='BSR' AND KEY2='SKIP_APPROVER_REMARKS' AND UNIT_CODE='" & gstrUnitId & "' ORDER BY CODE"
                        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Remark(s)")
                        If UBound(strHelp) > 0 Then
                            If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
                                MsgBox("Remark(s) Not Available.", MsgBoxStyle.Information, ResolveResString(100))
                                Exit Sub
                            End If
                            If IsNothing(strHelp) = False Then
                                intLoopCounter = Me.fspBarcode.ActiveRow
                                .SetText(enumBarcodeGrid.REMARKS, intLoopCounter, Trim(strHelp(0)).ToString())
                            End If
                        End If
                    End With
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnAuthorize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAuthorize.Click
        Try
            If ValidateRows("A") Then
                Exit Sub
            End If
            If MessageBox.Show("Are you sure you want to authorize selected barcode(s)?", "eMPro", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = Windows.Forms.DialogResult.Yes Then
                AuthorizeReject("A")
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnReject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReject.Click
        Try
            If ValidateRows("R") Then
                Exit Sub
            End If
            If MessageBox.Show("Are you sure you want to reject selected barcode(s)?", "eMPro", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly, False) = Windows.Forms.DialogResult.Yes Then
                AuthorizeReject("R")
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
End Class