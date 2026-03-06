Imports System.Data.SqlClient
Public Class FRMMKTTRN0084B
    Public gItem_Code As String = String.Empty
    Public gRate As Double = 0.0
    Public gcust_code As String = String.Empty
    Public g_newrate As Double = 0.0
    Public g_mode As String = String.Empty
    Public g_ProvDoc_No As String = String.Empty
    Dim Col_Reason As Object = Nothing, Col_Remarks As Object = Nothing, Col_Effect As Object = Nothing, Col_Value As Object = Nothing
    Dim intloopcounter As Int16 = 0
    Dim trans As Boolean = False

    Private Enum ENUMPRICECHANGEDETAILS
        Col_Reason = 1
        Col_Remarks = 2
        Col_Effect = 3
        Col_Value = 4
    End Enum
    Private Sub FRMMKTTRN0084B_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            lbl_item_code.Text = gItem_Code
            lbl_OldRate.Text = gRate
            ItemDesc(gItem_Code)
            Me.BringToFront()
            If g_mode = "ADD" Then
                If checkdata() = True Then

                Else
                    InitializeSpread()
                    AddRow()
                End If
            End If
            If g_mode = "VIEW" Then
                GETDATA("VIEW")
                LockGrid()
                btn_OK.Enabled = False
            End If
            If g_mode = "EDIT" Then
                GETDATA("VIEW")
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Protected Overrides Sub WndProc(ByRef message As Message)
        Const WM_SYSCOMMAND As Integer = &H112
        Const SC_MOVE As Integer = &HF010

        Select Case message.Msg
            Case WM_SYSCOMMAND
                Dim command As Integer = message.WParam.ToInt32() And &HFFF0
                If command = SC_MOVE Then
                    Return
                End If
                Exit Select
        End Select

        MyBase.WndProc(message)
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_EXIT.Click
        Try
            'If AxfpSpread1.MaxRows = 0 Then
            '    g_newrate = 0.0
            'Else
            '    If trans = True Then
            If (lbl_newrate.Text.Trim = "") Then
                g_newrate = 0.0
            Else
                g_newrate = Convert.ToDouble(lbl_newrate.Text.Trim)
            End If

            '    End If
            'End If
            Me.Close()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub InitializeSpread()
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            With Me.AxfpSpread1
                .MaxRows = 0
                .MaxCols = ENUMPRICECHANGEDETAILS.Col_Value
                .set_RowHeight(0, 20)

                .Row = 0 : .Col = ENUMPRICECHANGEDETAILS.Col_Reason : .Text = "REASON (F1)" : .set_ColWidth(.Col, 20) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMPRICECHANGEDETAILS.Col_Remarks : .Text = "REMARKS" : .set_ColWidth(.Col, 35) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMPRICECHANGEDETAILS.Col_Effect : .Text = "EFFECT" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMPRICECHANGEDETAILS.Col_Value : .Text = "VALUE" : .set_ColWidth(.Col, 15) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)

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
            With AxfpSpread1
                .MaxRows = .MaxRows + 1

                .Row = .MaxRows : .Col = ENUMPRICECHANGEDETAILS.Col_Reason : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMPRICECHANGEDETAILS.Col_Remarks : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .Lock = False : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid) : .TypeMaxEditLen = 100
                .Row = .MaxRows : .Col = ENUMPRICECHANGEDETAILS.Col_Effect : .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox : .TypeComboBoxList = "-" + Chr(9) + "+" : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMPRICECHANGEDETAILS.Col_Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .Lock = False : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeFloatMin = 0 : .TypeFloatDecimalPlaces = 4 : .BorderStyle = FPSpreadADO.BorderStyleConstants.BorderStyleFixedSingle : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
            End With

            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub AxfpSpread1_KeyPressEvent(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles AxfpSpread1.KeyPressEvent
        If e.keyAscii = 14 Then
            AddRow()
            With Me.AxfpSpread1

                .Row = .MaxRows : .Col = ENUMPRICECHANGEDETAILS.Col_Reason
                .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()

            End With
            Exit Sub
        End If

        If e.keyAscii = 4 Then
            If AxfpSpread1.MaxRows > 1 Then

                With AxfpSpread1

                    .Row = AxfpSpread1.ActiveRow

                    .Action = FPSpreadADO.ActionConstants.ActionDeleteRow
                    .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow

                    .Row = .MaxRows : .Col = ENUMPRICECHANGEDETAILS.Col_Reason
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()

                    .MaxRows = .MaxRows - 1

                End With
                'If AxfpSpread1.MaxRows = 0 Then
                '    lbl_newrate.Text = ""
                '    lbl_PriceChange.Text = ""
                'End If

            Else
                Exit Sub
            End If
        End If

    End Sub
    Private Sub AxfpSpread1_KeyDownEvent(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles AxfpSpread1.KeyDownEvent
        Dim strQuery As String = String.Empty
        Dim strHelp() As String
        Dim sqlCmd As New SqlCommand()
        Try
            With Me.AxfpSpread1
                If e.keyCode = Keys.F1 And g_mode.ToString <> "VIEW" Then
                    If AxfpSpread1.ActiveCol = ENUMPRICECHANGEDETAILS.Col_Reason Then
                        strQuery = "SELECT KEY2 AS REASON,ISNULL(CODE,'N/A') AS CodeType FROM LISTS WHERE UNIT_CODE='" + gstrUNITID + "' AND KEY1 ='SALESPROVISION'"
                        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Reason")
                        If UBound(strHelp) > 0 Then
                            If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
                                MsgBox("Reason Not Available.", MsgBoxStyle.Information, ResolveResString(100))
                                Exit Sub
                            End If
                            If IsNothing(strHelp) = False Then
                                With Me.AxfpSpread1
                                    For intloopcounter = 1 To .MaxRows
                                        Col_Reason = Nothing
                                        .GetText(ENUMPRICECHANGEDETAILS.Col_Reason, intloopcounter, Col_Reason)
                                        If IsNothing(Col_Reason) = True Then Col_Reason = String.Empty
                                        If (Col_Reason.ToString() = Trim(strHelp(0)).ToString()) Then
                                            MsgBox("Reason Already Selected.", MsgBoxStyle.Information, ResolveResString(100))
                                            Exit Sub
                                        End If
                                    Next
                                End With
                                .SetText(ENUMPRICECHANGEDETAILS.Col_Reason, AxfpSpread1.ActiveRow, Trim(strHelp(0)).ToString())
                            End If
                        End If
                    End If
                End If
            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            If sqlCmd.Connection.State = ConnectionState.Open Then sqlCmd.Connection.Close()
            sqlCmd.Connection.Dispose()
            sqlCmd.Dispose()
        End Try
    End Sub
    Private Sub ItemDesc(ByRef item_code As String)

        Dim sqlCmd As New SqlCommand()
        Dim SqlAdp As SqlDataAdapter
        Dim DSCONTRACTDTL As DataSet

        SqlAdp = New SqlDataAdapter
        DSCONTRACTDTL = New DataSet

        Try

            With sqlCmd
                .CommandText = "USP_PRICE_CHANGE_DETAILS"
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .Connection = SqlConnectionclass.GetConnection()
                .Parameters.Clear()
                .Parameters.AddWithValue("@MODE", "ITEM")
                .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                .Parameters.AddWithValue("@CUSTOMER_CODE", gcust_code)
                .Parameters.AddWithValue("@ITEM_CODE", item_code)
                .Parameters.Add(New SqlParameter("@p_ERROR", SqlDbType.VarChar, 200, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                SqlAdp.SelectCommand = sqlCmd
                SqlAdp.Fill(DSCONTRACTDTL)
                .Dispose()
                If Not String.IsNullOrEmpty(.Parameters("@p_ERROR").Value) Then
                    MessageBox.Show(Convert.ToString(sqlCmd.Parameters("@p_ERROR").Value), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                End If

                If DSCONTRACTDTL.Tables.Count > 0 Then
                    If DSCONTRACTDTL.Tables(0).Rows.Count > 0 Then
                        lblitemdesc.Text = DSCONTRACTDTL.Tables(0).Rows(0).Item("DESCRIPTION").ToString()
                    End If
                    If DSCONTRACTDTL.Tables(1).Rows.Count > 0 Then
                        lbl_Custitemcode.Text = DSCONTRACTDTL.Tables(1).Rows(0).Item("CUST_DRGNO").ToString()
                    End If
                End If
            End With

        Catch ex As Exception
            RaiseException(ex)
        Finally
            If sqlCmd.Connection.State = ConnectionState.Open Then sqlCmd.Connection.Close()
            sqlCmd.Connection.Dispose()
            sqlCmd.Dispose()
            SqlAdp.Dispose()
            DSCONTRACTDTL.Dispose()
        End Try
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_OK.Click
        Try
            If g_mode.ToString() = "ADD" Or g_mode.ToString() = "EDIT" Then
                If Me.AxfpSpread1.MaxRows = 0 Then
                    Dim strquery As String = String.Empty
                    Dim SqlCmd As New SqlCommand
                    With SqlCmd
                        .Connection = SqlConnectionclass.GetConnection()
                        .CommandText = "DELETE FROM TMP_PRICECHANGE_DTL WHERE rate=@RATE AND unitcode=@UNIT_CODE AND itemcode=@ITEM_CODE and CUSTOMEr_CODE=@CUSTOMER_CODE and ipaddress=@IPADDRESS"
                        .CommandType = CommandType.Text
                        .CommandTimeout = 0
                        .Parameters.Clear()
                        .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                        .Parameters.AddWithValue("@ITEM_CODE", gItem_Code)
                        .Parameters.AddWithValue("@RATE", gRate)
                        .Parameters.AddWithValue("@CUSTOMER_CODE", gcust_code)
                        .Parameters.AddWithValue("@IPADDRESS", gstrIpaddressWinSck)
                        .ExecuteNonQuery()
                    End With
                    'MessageBox.Show("Record updated successfully.")
                    lbl_newrate.Text = ""
                    lbl_PriceChange.Text = ""
                    g_newrate = 0.0
                    Me.Close()
                End If
            End If

            If Me.AxfpSpread1.MaxRows > 0 Then
                If ValidData() = True Then
                    SaveData()
                    Calculation()
                    g_newrate = Convert.ToDouble(lbl_newrate.Text.Trim)
                    Me.Close()
                Else
                    With Me.AxfpSpread1
                        .Row = .MaxRows : .Col = ENUMPRICECHANGEDETAILS.Col_Reason
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                    End With
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub SaveData()
        Dim data_farspreadGrid As New DataTable  '  Contract dtl table data
        Dim sqlCmd As New SqlCommand()

        Try
            data_farspreadGrid.Columns.Add("REASON", GetType(String))
            data_farspreadGrid.Columns.Add("REMARKS", GetType(String))
            data_farspreadGrid.Columns.Add("EFFECT", GetType(String))
            data_farspreadGrid.Columns.Add("VALUE", GetType(Double))

            With Me.AxfpSpread1
                For intloopcounter = 1 To .MaxRows
                    Dim data_newRow As DataRow = data_farspreadGrid.NewRow

                    Col_Reason = Nothing
                    .GetText(ENUMPRICECHANGEDETAILS.Col_Reason, intloopcounter, Col_Reason)
                    If IsNothing(Col_Reason) = True Then Col_Reason = String.Empty

                    Col_Remarks = Nothing
                    .GetText(ENUMPRICECHANGEDETAILS.Col_Remarks, intloopcounter, Col_Remarks)
                    If IsNothing(Col_Remarks) = True Then Col_Remarks = String.Empty

                    Col_Effect = Nothing
                    .GetText(ENUMPRICECHANGEDETAILS.Col_Effect, intloopcounter, Col_Effect)
                    If IsNothing(Col_Effect) = True Then Col_Effect = String.Empty

                    Col_Value = Nothing
                    .GetText(ENUMPRICECHANGEDETAILS.Col_Value, intloopcounter, Col_Value)
                    If IsNothing(Col_Value) = True Then Col_Value = 0.0

                    If (Col_Reason.ToString() <> "") Then

                        data_newRow("REASON") = Col_Reason.ToString()
                        data_newRow("REMARKS") = Col_Remarks.ToString()
                        data_newRow("EFFECT") = Col_Effect.ToString()
                        data_newRow("VALUE") = Convert.ToDouble(Col_Value)

                        data_farspreadGrid.Rows.Add(data_newRow)

                    End If

                Next
            End With

            With sqlCmd
                .CommandText = "USP_PRICE_CHANGE_DETAILS"
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .Connection = SqlConnectionclass.GetConnection()
                .Parameters.Clear()
                .Parameters.AddWithValue("@MODE", "SAVETEMP")
                .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                .Parameters.AddWithValue("@RATE", gRate)
                .Parameters.AddWithValue("@NEWRATE", Convert.ToDouble(lbl_newrate.Text.Trim))
                .Parameters.AddWithValue("@ITEM_CODE", gItem_Code)
                .Parameters.AddWithValue("@ITEM_DESC", lblitemdesc.Text.Trim.ToString())
                .Parameters.AddWithValue("@CUSTOMER_CODE", gcust_code)
                .Parameters.AddWithValue("@CUSTITEM_CODE", lbl_Custitemcode.Text.Trim.ToString())
                .Parameters.AddWithValue("@IPADDRESS", gstrIpaddressWinSck)
                .Parameters.AddWithValue("@PRICE_DTLS", data_farspreadGrid)
                .Parameters.Add(New SqlParameter("@p_ERROR", SqlDbType.VarChar, 200, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))

                .ExecuteNonQuery()
                .Dispose()
                If Not String.IsNullOrEmpty(.Parameters("@p_ERROR").Value) Then
                    MessageBox.Show(Convert.ToString(sqlCmd.Parameters("@p_ERROR").Value), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                End If
                trans = True
            End With

        Catch ex As Exception
            RaiseException(ex)
        Finally
            If sqlCmd.Connection.State = ConnectionState.Open Then sqlCmd.Connection.Close()
            sqlCmd.Connection.Dispose()
            sqlCmd.Dispose()
            data_farspreadGrid.Dispose()

        End Try
    End Sub
    Private Function ValidData() As Boolean
        Try
            With Me.AxfpSpread1
                For intloopcounter = 1 To .MaxRows
                    Col_Reason = Nothing
                    .GetText(ENUMPRICECHANGEDETAILS.Col_Reason, intloopcounter, Col_Reason)
                    If IsNothing(Col_Reason) = True Then Col_Reason = String.Empty
                    If Col_Reason.ToString() = "" Then
                        MessageBox.Show("Kindly Select the Reason for Row no." + intloopcounter.ToString(), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Return False
                    End If

                    Col_Effect = Nothing
                    .GetText(ENUMPRICECHANGEDETAILS.Col_Effect, intloopcounter, Col_Effect)
                    If IsNothing(Col_Effect) = True Then Col_Effect = String.Empty
                    If Col_Effect.ToString() = "" Then
                        MessageBox.Show("Kindly Select the Effect for Row no." + intloopcounter.ToString(), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Return False
                    End If


                    Col_Value = Nothing
                    .GetText(ENUMPRICECHANGEDETAILS.Col_Value, intloopcounter, Col_Value)
                    If IsNothing(Col_Value) = True Then Col_Value = 0.0
                    If Convert.ToDouble(Col_Value) = 0.0 Then
                        MessageBox.Show("Kindly Select the Value for Row no." + intloopcounter.ToString(), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Return False
                    End If

                    Col_Remarks = Nothing
                    .GetText(ENUMPRICECHANGEDETAILS.Col_Remarks, intloopcounter, Col_Remarks)
                    If IsNothing(Col_Remarks) = True Then Col_Remarks = String.Empty
                    If Col_Remarks.ToString() = "" Then
                        MessageBox.Show("Kindly Enter the Remark for Row no." + intloopcounter.ToString(), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        Return False
                    End If
                Next
            End With
            Return True
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function
    Private Sub Calculation()
        Try
            Dim positivevalue As Double = 0.0
            Dim negativevalue As Double = 0.0

            With Me.AxfpSpread1
                For intloopcounter = 1 To .MaxRows

                    Col_Effect = Nothing
                    .GetText(ENUMPRICECHANGEDETAILS.Col_Effect, intloopcounter, Col_Effect)
                    If IsNothing(Col_Effect) = True Then Col_Effect = String.Empty

                    Col_Value = Nothing
                    .GetText(ENUMPRICECHANGEDETAILS.Col_Value, intloopcounter, Col_Value)
                    If IsNothing(Col_Value) = True Then Col_Value = 0.0

                    If Col_Effect.ToString = "+" Then
                        positivevalue = positivevalue + Convert.ToDouble(Col_Value)
                    End If
                    If Col_Effect.ToString = "-" Then
                        negativevalue = negativevalue + Convert.ToDouble(Col_Value)
                    End If
                Next
                lbl_PriceChange.Text = positivevalue - negativevalue
                lbl_newrate.Text = Convert.ToDouble(lbl_OldRate.Text.Trim()) + Convert.ToDouble(lbl_PriceChange.Text.Trim())
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Function checkdata() As Boolean

        Dim sqlCmd As New SqlCommand()
        Dim SqlAdp As SqlDataAdapter
        Dim DSCONTRACTDTL As DataSet

        SqlAdp = New SqlDataAdapter
        DSCONTRACTDTL = New DataSet
        Try

            With sqlCmd
                .CommandText = "USP_PRICE_CHANGE_DETAILS"
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .Connection = SqlConnectionclass.GetConnection()
                .Parameters.Clear()
                .Parameters.AddWithValue("@MODE", "VIEWTEMP")
                .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                .Parameters.AddWithValue("@ITEM_CODE", gItem_Code)
                .Parameters.AddWithValue("@RATE", gRate)
                .Parameters.AddWithValue("@CUSTOMER_CODE", gcust_code)
                .Parameters.AddWithValue("@IPADDRESS", gstrIpaddressWinSck)
                .Parameters.Add(New SqlParameter("@p_ERROR", SqlDbType.VarChar, 200, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                SqlAdp.SelectCommand = sqlCmd
                SqlAdp.Fill(DSCONTRACTDTL)
                .Dispose()
                If Not String.IsNullOrEmpty(.Parameters("@p_ERROR").Value) Then
                    MessageBox.Show(Convert.ToString(sqlCmd.Parameters("@p_ERROR").Value), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                End If
            End With
            If DSCONTRACTDTL.Tables(0).Rows.Count > 0 Then
                InitializeSpread()
                With Me.AxfpSpread1
                    For intloopcounter = 0 To DSCONTRACTDTL.Tables(0).Rows.Count - 1
                        AddRow()
                        .SetText(ENUMPRICECHANGEDETAILS.Col_Reason, intloopcounter + 1, DSCONTRACTDTL.Tables(0).Rows(intloopcounter).Item("REASON").ToString.Trim)
                        .SetText(ENUMPRICECHANGEDETAILS.Col_Remarks, intloopcounter + 1, DSCONTRACTDTL.Tables(0).Rows(intloopcounter).Item("REMARKS").ToString.Trim)
                        .SetText(ENUMPRICECHANGEDETAILS.Col_Effect, intloopcounter + 1, DSCONTRACTDTL.Tables(0).Rows(intloopcounter).Item("EFFECT").ToString.Trim)
                        .SetText(ENUMPRICECHANGEDETAILS.Col_Value, intloopcounter + 1, Convert.ToDouble(DSCONTRACTDTL.Tables(0).Rows(intloopcounter).Item("VALUE")))
                    Next
                End With
                ' lbl_newrate.Text = Convert.ToDouble(DSCONTRACTDTL.Tables(0).Rows(0).Item("VALUE"))
                Calculation()
            End If
            If DSCONTRACTDTL.Tables(0).Rows.Count = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            If sqlCmd.Connection.State = ConnectionState.Open Then sqlCmd.Connection.Close()
            sqlCmd.Connection.Dispose()
            sqlCmd.Dispose()
            SqlAdp.Dispose()
            DSCONTRACTDTL.Dispose()
        End Try
    End Function
    Private Sub AxfpSpread1_EditChange(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles AxfpSpread1.EditChange
        Try
            If Me.AxfpSpread1.ActiveCol = ENUMPRICECHANGEDETAILS.Col_Value Then
                With Me.AxfpSpread1
                    Col_Reason = Nothing
                    .GetText(ENUMPRICECHANGEDETAILS.Col_Reason, .ActiveRow, Col_Reason)
                    If IsNothing(Col_Reason) = True Then Col_Reason = String.Empty
                    If Col_Reason.ToString() = "" Then
                        MessageBox.Show("Kindly Select the Reason for Row no." + .ActiveRow.ToString(), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        .SetText(ENUMPRICECHANGEDETAILS.Col_Value, .ActiveRow, 0.0)
                        Exit Sub
                    End If

                    Col_Effect = Nothing
                    .GetText(ENUMPRICECHANGEDETAILS.Col_Effect, .ActiveRow, Col_Effect)
                    If IsNothing(Col_Effect) = True Then Col_Effect = String.Empty
                    If Col_Effect.ToString() = "" Then
                        MessageBox.Show("Kindly Select the Effect for Row no." + .ActiveRow.ToString(), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        .SetText(ENUMPRICECHANGEDETAILS.Col_Value, .ActiveRow, 0.0)
                        Exit Sub
                    End If
                End With
                Calculation()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub AxfpSpread1_ComboSelChange(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ComboSelChangeEvent) Handles AxfpSpread1.ComboSelChange
        Try
            If Me.AxfpSpread1.ActiveCol = ENUMPRICECHANGEDETAILS.Col_Effect Then
                With Me.AxfpSpread1
                    Col_Reason = Nothing
                    .GetText(ENUMPRICECHANGEDETAILS.Col_Reason, .ActiveRow, Col_Reason)
                    If IsNothing(Col_Reason) = True Then Col_Reason = String.Empty
                    If Col_Reason.ToString() = "" Then
                        MessageBox.Show("Kindly Select the Reason for Row no." + .ActiveRow.ToString(), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                        .SetText(ENUMPRICECHANGEDETAILS.Col_Effect, .ActiveRow, "")
                        Exit Sub
                    End If
                    Calculation()
                End With
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub GETDATA(ByVal mode As String)
        Dim sqlCmd As New SqlCommand()
        Dim SqlAdp As SqlDataAdapter
        Dim DSCONTRACTDTL As DataSet

        SqlAdp = New SqlDataAdapter
        DSCONTRACTDTL = New DataSet
        Try

            With sqlCmd
                .CommandText = "USP_PRICE_CHANGE_DETAILS"
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .Connection = SqlConnectionclass.GetConnection()
                .Parameters.Clear()
                .Parameters.AddWithValue("@MODE", mode)
                .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                .Parameters.AddWithValue("@ITEM_CODE", gItem_Code)
                .Parameters.AddWithValue("@RATE", gRate)
                .Parameters.AddWithValue("@CUSTOMER_CODE", gcust_code)
                .Parameters.AddWithValue("@IPADDRESS", g_ProvDoc_No.ToString())  ' here in ip address i am passing doc no
                .Parameters.Add(New SqlParameter("@p_ERROR", SqlDbType.VarChar, 200, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                SqlAdp.SelectCommand = sqlCmd
                SqlAdp.Fill(DSCONTRACTDTL)
                .Dispose()
                If Not String.IsNullOrEmpty(.Parameters("@p_ERROR").Value) Then
                    MessageBox.Show(Convert.ToString(sqlCmd.Parameters("@p_ERROR").Value), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                End If
            End With
            If DSCONTRACTDTL.Tables(0).Rows.Count > 0 Then
                InitializeSpread()
                With Me.AxfpSpread1
                    For intloopcounter = 0 To DSCONTRACTDTL.Tables(0).Rows.Count - 1
                        AddRow()
                        .SetText(ENUMPRICECHANGEDETAILS.Col_Reason, intloopcounter + 1, DSCONTRACTDTL.Tables(0).Rows(intloopcounter).Item("REASON").ToString.Trim)
                        .SetText(ENUMPRICECHANGEDETAILS.Col_Remarks, intloopcounter + 1, DSCONTRACTDTL.Tables(0).Rows(intloopcounter).Item("REMARKS").ToString.Trim)
                        .SetText(ENUMPRICECHANGEDETAILS.Col_Effect, intloopcounter + 1, DSCONTRACTDTL.Tables(0).Rows(intloopcounter).Item("EFFECT").ToString.Trim)
                        .SetText(ENUMPRICECHANGEDETAILS.Col_Value, intloopcounter + 1, Convert.ToDouble(DSCONTRACTDTL.Tables(0).Rows(intloopcounter).Item("VALUE")))
                    Next
                End With
                ' lbl_newrate.Text = Convert.ToDouble(DSCONTRACTDTL.Tables(0).Rows(0).Item("VALUE"))
                Calculation()
            End If
            If DSCONTRACTDTL.Tables(0).Rows.Count = 0 Then
                InitializeSpread()
                AddRow()
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            If sqlCmd.Connection.State = ConnectionState.Open Then sqlCmd.Connection.Close()
            sqlCmd.Connection.Dispose()
            sqlCmd.Dispose()
            SqlAdp.Dispose()
            DSCONTRACTDTL.Dispose()
        End Try
    End Sub
    Private Sub LockGrid()
        Try
            With Me.AxfpSpread1
                .BlockMode = True
                .Row = 1
                .Row2 = .MaxRows
                .Col = 1
                .Col2 = .MaxCols
                .Lock = True
                .BlockMode = False
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
End Class