Option Strict Off
Option Explicit On
Imports System.Data.SqlClient
Imports system
Public Class frmMKTMST0020
    '-----------------------------------------------------------------------
    ' Copyright (c)         :MIND Ltd.
    ' Form Name             :frmMKTMST0020
    ' Function Name         :To MAP MARUTI PO WITH EMPOWER SO
    ' Created By            :Siddharth Ranjan
    ' Created On            :31-MAY-2010
    ' Modify Date           :NIL
    ' Revision History      :-
    'Revised By             :Shubhra Verma
    'Revised On             :21 Apr 2011
    'Description            :Multi Unit Changes
    'Modified by Shubra Verma on 25/Apr/2011 for multiunit change 
    '-----------------------------------------------------------------------
    Dim mintIndex As Short
    Dim mstrItemList As String
    Private Enum enumGrid
        col_Item_Code = 1
        col_Customer_part = 2
        col_Empower_SO = 3
        col_Amendment_no = 4
        col_Maruti_PO = 5
    End Enum
    Private Sub CmdGrpEdit_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtngrptwo.ButtonClickEventArgs) Handles CmdGrpEdit.ButtonClick
        Select Case Me.CmdGrpEdit.Mode
            Case "S"
                If SaveData() Then
                    MsgBox("Mapping Saved Successfuly.", MsgBoxStyle.Information, ResolveResString(100))
                    txtCustomerCode.Text = ""
                    txtCustomerCode_TextChanged(txtCustomerCode, New System.EventArgs())
                    CmdGrpEdit.Revert()
                End If
            Case ""
                txtCustomerCode.Text = ""
                Call txtCustomerCode_TextChanged(txtCustomerCode, New System.EventArgs())
                CmdGrpEdit.Revert()
            Case "E"
                If fpsSpread.MaxRows = 0 Then
                    MsgBox("No Data To Edit.", MsgBoxStyle.Information, ResolveResString(100))
                    CmdGrpEdit.Revert()
                    Exit Sub
                End If
                With fpsSpread
                    .Col = enumGrid.col_Maruti_PO
                    .Col2 = enumGrid.col_Maruti_PO
                    .Row = 1
                    .Row2 = .MaxRows
                    .BlockMode = True
                    .Lock = False
                    .BlockMode = False
                End With
            Case "X"
                Me.Close()
        End Select
    End Sub
    Private Function SaveData() As Boolean
        On Error GoTo ErrHandler
        Dim strSql As String
        Dim intcounter As Int32
        Dim objItem_code As Object
        Dim objcust_drg_no As Object
        Dim objcust_ref As Object
        Dim objcust_PO_NO As Object
        Dim objAmend_NO As Object
        Dim objSqlConn As SqlConnection
        Dim objComm As SqlCommand
        Dim objTrans As SqlTransaction
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
        If fpsSpread.MaxRows = 0 Then
            MsgBox("No Data To Update.", MsgBoxStyle.Information, ResolveResString(100))
            SaveData = False
            Exit Function
        End If
        objSqlConn = SqlConnectionclass.GetConnection()
        objTrans = objSqlConn.BeginTransaction
        With fpsSpread
            For intcounter = 1 To .MaxRows
                objItem_code = Nothing
                .GetText(enumGrid.col_Item_Code, intcounter, objItem_code)
                objcust_drg_no = Nothing
                .GetText(enumGrid.col_Customer_part, intcounter, objcust_drg_no)
                objcust_ref = Nothing
                .GetText(enumGrid.col_Empower_SO, intcounter, objcust_ref)
                objAmend_NO = Nothing
                .GetText(enumGrid.col_Amendment_no, intcounter, objAmend_NO)
                objcust_PO_NO = Nothing
                .GetText(enumGrid.col_Maruti_PO, intcounter, objcust_PO_NO)
                If objcust_PO_NO.ToString.Length = 0 Then
                    objcust_PO_NO = DBNull.Value
                End If
                strSql = "IF EXISTS(SELECT TOP 1 1 FROM PO_MAPPING_DATA WHERE UNIT_CODE = '" & gstrUNITID & "' AND ACCOUNT_CODE = '" & txtCustomerCode.Text & "' AND ITEM_CODE = '" & objItem_code.ToString & "' AND CUST_REF = '" & objcust_ref.ToString & "' AND CUST_DRGNO = '" & objcust_drg_no.ToString & "' AND AMENDMENT_NO = '" & objAmend_NO.ToString & "')" & _
                        " UPDATE PO_MAPPING_DATA SET CUST_PO_REF = '" & objcust_PO_NO & "' WHERE UNIT_CODE = '" & gstrUNITID & "' AND ACCOUNT_CODE = '" & txtCustomerCode.Text & "' AND ITEM_CODE = '" & objItem_code.ToString & "' AND CUST_REF = '" & objcust_ref.ToString & "' AND CUST_DRGNO = '" & objcust_drg_no.ToString & "' AND AMENDMENT_NO = '" & objAmend_NO.ToString & "'" & _
                        " ELSE" & _
                        " INSERT PO_MAPPING_DATA(ACCOUNT_CODE, ITEM_CODE, CUST_DRGNO, CUST_REF, AMENDMENT_NO, CUST_PO_REF, ENT_DT, ENT_USERID, UPD_DT, UPD_USERID,UNIT_CODE)" & _
                        " VALUES('" & txtCustomerCode.Text & "','" & objItem_code.ToString & "','" & objcust_drg_no.ToString & "','" & objcust_ref.ToString & "','" & objAmend_NO.ToString & "','" & objcust_PO_NO.ToString & "',CONVERT(DATETIME,(CONVERT(VARCHAR(12),GETDATE())),106),'" & mP_User & "',CONVERT(DATETIME,(CONVERT(VARCHAR(12),GETDATE())),106),'" & mP_User & "','" & gstrUNITID & "')"
                objComm = New SqlCommand(strSql, objSqlConn)
                objComm.CommandType = CommandType.Text
                objComm.Transaction = objTrans
                objComm.ExecuteNonQuery()

            Next intcounter
            objTrans.Commit()
            objComm = Nothing
            objSqlConn = Nothing
        End With
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        SaveData = True
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        SaveData = False
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        objComm = Nothing
        objSqlConn = Nothing
    End Function
    Private Sub frmMKTMST0020_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0020_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0020_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTMST0020_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlFormHeader_Click(ctlFormHeader, New System.EventArgs())
        End If
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ctlFormHeader_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0020_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{tab}")
        End If
        GoTo EventExitSub 'To prevent the execution of errhandler
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub
    Private Sub frmMKTMST0020_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000
        On Error GoTo ErrHandler ' Error Handler
        If KeyCode = System.Windows.Forms.Keys.Escape And Shift = 0 Then Me.Close()
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0020_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        On Error GoTo ErrHandler ' Error Handler
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader.Tag)
        FitToClient(Me, fraMain, ctlFormHeader, CmdGrpEdit, 0)
        InitializeGrid()
        optItemCode.Enabled = False
        optDescription.Enabled = False
        cmdCustomerHlp.Enabled = True
        fpsSpread.Enabled = False
        txtCustomerCode.Focus()
        optAll.Checked = True
        Me.optUnMapped.Checked = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub InitializeGrid()
        On Error GoTo ErrHandler ' Error Handler
        With fpsSpread
            .MaxCols = 5
            .MaxRows = 0
            .set_RowHeight(0, 570)
            .BackColor = Color.Beige
            .UserResize = FPSpreadADO.UserResizeConstants.UserResizeNone
            .ShadowColor = Color.WhiteSmoke
            .Row = 0
            .Col = enumGrid.col_Item_Code
            .Value = "Item Code"
            .set_ColWidth(enumGrid.col_Item_Code, 1500)
            .FontBold = True
            .Col = enumGrid.col_Customer_part
            .Value = "Customer Part Code"
            .set_ColWidth(enumGrid.col_Customer_part, 2300)
            .FontBold = True
            .Col = enumGrid.col_Empower_SO
            .Value = "SO No."
            .set_ColWidth(enumGrid.col_Empower_SO, 1500)
            .FontBold = True
            .Col = enumGrid.col_Amendment_no
            .Value = "Amendment No."
            .set_ColWidth(enumGrid.col_Amendment_no, 1400)
            .FontBold = True
            .Col = enumGrid.col_Maruti_PO
            .Value = "SO Mapped With"
            .set_ColWidth(enumGrid.col_Maruti_PO, 1500)
            .FontBold = True
            .AutoClipboard = False
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdCustomerHlp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCustomerHlp.Click
        Dim strQry As String 'Holds the query to geneate for work order help button
        Dim strWO() As String 'Holds the array returned by showlist
        On Error GoTo ErrHandler
        If Trim(Me.txtCustomerCode.Text) = "" Then
            strQry = "SELECT A.CUSTOMER_CODE, B.CUST_NAME FROM PO_AUTO_AMEND_CUSTOMER_MST A INNER JOIN CUSTOMER_MST B ON A.UNIT_CODE = B.UNIT_CODE AND A.UNIT_CODE = '" & gstrUNITID & "' AND A.CUSTOMER_CODE = B.CUSTOMER_CODE"
            strWO = Me.ctlEMPHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry, "Customer Code")
        Else
            strQry = "SELECT A.CUSTOMER_CODE, B.CUST_NAME FROM PO_AUTO_AMEND_CUSTOMER_MST A INNER JOIN CUSTOMER_MST B ON A.UNIT_CODE = B.UNIT_CODE AND A.CUSTOMER_CODE = B.CUSTOMER_CODE WHERE A.UNIT_CODE = '" & gstrUNITID & "' AND A.CUSTOMER_CODE = '" & Trim(txtCustomerCode.Text) & "'"
            strWO = Me.ctlEMPHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry, "Customer Code")
        End If
        If Not (UBound(strWO) <= 0) Then
            If (Len(strWO(0)) >= 1) And strWO(0) = "0" Then
                MsgBox("Invalid Customer Code", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                txtCustomerCode.Text = ""
                txtCustomerCode.Focus()
                Exit Sub
            Else
                Me.txtCustomerCode.Text = strWO(0)
                Me.lblDescription.Text = strWO(1)
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub optSelected_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optSelected.CheckedChanged
        On Error GoTo ErrHandler
        If optSelected.Checked Then
            If (txtCustomerCode.Text = "") Then
                MsgBox("Customer Is Blank. Please Select Customer.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
            fpsSpread.MaxRows = 0
            optSelected.Checked = True
            Me.LvwItem.Enabled = True
            Me.LvwItem.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            Call FillItems()
            Me.LvwItem.Focus()
            optItemCode.Enabled = True
            optItemCode.Checked = True
            optDescription.Enabled = True
            txtSearched.Enabled = True
            txtSearched.Focus()
            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        End If
        Exit Sub
ErrHandler:
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub optSelected_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles optSelected.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar) 'If user presses enter, the focus will be set in the next control.
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Me.LvwItem.Focus()
            Case System.Windows.Forms.Keys.Escape
                Me.Close()
        End Select
        GoTo EventExitSub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        gblnCancelUnload = True
EventExitSub:
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub
    Private Sub FillItems()
        'On Error GoTo ErrHandler ' & will display the corresponding Item codes.
        Dim strsql As String
        Dim rsLocItem As New ADODB.Recordset
        Dim lngLoop As Integer, lngrecords As Int32, lngcount1 As Int32
        Dim LstItem As System.Windows.Forms.ListViewItem
        strsql = String.Empty
        Me.LvwItem.Columns.Clear()
        Me.LvwItem.Items.Clear()
        Me.LvwItem.View = System.Windows.Forms.View.Details
        Me.LvwItem.GridLines = True
        Me.LvwItem.CheckBoxes = True
        strsql = "SELECT * FROM DBO.UDF_PO_MAPPING_CUSTOMER_HLP('" & txtCustomerCode.Text.Trim & "','" & gstrUNITID & "') ORDER BY ITEM_CODE"
        Call rsLocItem.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        For lngLoop = 1 To rsLocItem.Fields.Count
            Me.LvwItem.Columns.Add(rsLocItem.Fields(lngLoop - 1).Name)
        Next
        lngrecords = rsLocItem.RecordCount
        If lngrecords > 0 Then
            rsLocItem.MoveFirst()
            For lngLoop = 0 To LvwItem.Items.Count - 1
                Me.LvwItem.Items.Clear()
            Next
            If Not (rsLocItem.BOF And rsLocItem.EOF) Then
                For lngcount1 = 1 To lngrecords
                    LstItem = Me.LvwItem.Items.Add(Trim(rsLocItem.Fields("Item_Code").Value))
                    For lngLoop = 1 To rsLocItem.Fields.Count - 1
                        If LstItem.SubItems.Count > lngLoop Then
                            LstItem.SubItems(lngLoop).Text = rsLocItem.Fields(lngLoop).Value
                        Else
                            LstItem.SubItems.Insert(lngLoop, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsLocItem.Fields(lngLoop).Value))
                        End If
                    Next
                    rsLocItem.MoveNext()
                Next lngcount1
                rsLocItem.Close()
                rsLocItem = Nothing
            End If
        Else
            If rsLocItem.State = 1 Then rsLocItem.Close()
            Exit Sub
        End If
        Call ListHeaders((Me.LvwItem))
        Me.LvwItem.Columns.Item(0).Width = 120
        Me.LvwItem.Columns.Item(1).Width = 180

        Exit Sub
ErrHandler:
        If Err.Number = 3021 Then
            Resume Next
        Else
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
            gblnCancelUnload = True
        End If
    End Sub
    Private Sub optAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAll.CheckedChanged
        If optAll.Checked Then 'If user selects AllItem Control,then the listing of items will be disabled.
            On Error GoTo ErrHandler
            Me.LvwItem.Enabled = False
            Me.LvwItem.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED_LIST_VIEW)
            Me.LvwItem.Items.Clear()
            Me.LvwItem.Columns.Clear()
            fpsSpread.MaxRows = 0
            optItemCode.Enabled = False
            optDescription.Enabled = False
            txtSearched.Text = ""
            txtSearched.Enabled = False
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        gblnCancelUnload = True
    End Sub
    Private Sub txtCustomerCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Dim Shifts As Short = e.KeyData \ &H10000
        If e.KeyCode = Keys.F1 And Shifts = 0 Then
            cmdCustomerHlp_Click(cmdCustomerHlp, New System.EventArgs())
        End If
    End Sub
    Private Sub txtCustomerCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.TextChanged
        Me.LvwItem.Enabled = False
        Me.LvwItem.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED_LIST_VIEW)
        Me.LvwItem.Items.Clear()
        Me.LvwItem.Columns.Clear()
        fpsSpread.MaxRows = 0
        optItemCode.Enabled = False
        optDescription.Enabled = False
        txtSearched.Text = ""
        txtSearched.Enabled = False
        lblDescription.Text = ""
        optAll.Checked = True
    End Sub
    Private Sub txtCustomerCode_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
        If txtCustomerCode.Text.Trim <> "" And lblDescription.Text.Length = 0 Then
            cmdCustomerHlp_Click(cmdCustomerHlp, New System.EventArgs())
        End If
    End Sub
    Private Sub txtSearched_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSearched.TextChanged
        On Error GoTo ErrHandler
        Call SearchItem()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub SearchItem()
        '---------------------------------------------------------------------
        'Created By     -   Shruti Khanna\(Name Changed - Nitin Sood)
        '---------------------------------------------------------------------
        Dim itmFound As System.Windows.Forms.ListViewItem ' FoundItem variable.
        On Error GoTo ErrHandler
        itmFound = SearchText((txtSearched.Text), optDescription, LvwItem)
        If itmFound Is Nothing Then ' If no match,
            Exit Sub
        Else
            itmFound.EnsureVisible() ' Scroll ListView to show found ListItem.
            itmFound.Selected = True ' Select the ListItem.
            ' Return focus to the control to see selection.
            LvwItem.Enabled = True
            If Len(txtSearched.Text) > 0 Then itmFound.Font = VB6.FontChangeBold(itmFound.Font, True)
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub CheckItem() 'this function will check, if user has selected any itemcode or not
        On Error GoTo ErrHandler
        Dim intCount As Short
        mstrItemList = ""
        If Me.optSelected.Checked Then
            For intCount = 0 To Me.LvwItem.Items.Count - 1
                If Me.LvwItem.Items.Item(intCount).Checked = True Then mstrItemList = mstrItemList & Me.LvwItem.Items.Item(intCount).Text & ";"
            Next intCount
        End If
        If mstrItemList.Length > 0 Then
            mstrItemList = Mid(mstrItemList, 1, Len(mstrItemList) - 1)
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdShowData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdShowData.Click
        On Error GoTo ErrHandler
        If txtCustomerCode.Text.Length = 0 Then
            MsgBox("Please Select Customer.", MsgBoxStyle.Information, ResolveResString(100))
            txtCustomerCode.Focus()
            Exit Sub
        End If
        Dim objSQLConn As SqlConnection
        Dim objReader As SqlDataReader
        Dim objCommand As SqlCommand
        Dim STRSQL As String
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
        If optSelected.Checked Then
            Call CheckItem()
        Else
            mstrItemList = ""
        End If
        STRSQL = "SELECT * FROM DBO.UDF_PO_MAPPING_DATA('" & txtCustomerCode.Text.Trim & "','" & mstrItemList & "','" & IIf((optUnMapped.Checked), "U", (IIf((optMapped.Checked), "M", "A"))) & "','" & gstrUNITID & "')"
        objSQLConn = SqlConnectionclass.GetConnection()
        objCommand = New SqlCommand(STRSQL, objSQLConn)
        objReader = objCommand.ExecuteReader()
        If objReader.HasRows Then
            fpsSpread.Enabled = True
            fpsSpread.MaxRows = 0
            With fpsSpread
                While objReader.Read
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .set_RowHeight(.Row, 250)
                    .SetText(enumGrid.col_Item_Code, .Row, objReader.GetValue(0))
                    .SetText(enumGrid.col_Customer_part, .Row, objReader.GetValue(1))
                    .SetText(enumGrid.col_Empower_SO, .Row, objReader.GetValue(2))
                    .SetText(enumGrid.col_Amendment_no, .Row, objReader.GetValue(3))
                    .SetText(enumGrid.col_Maruti_PO, .Row, objReader.GetValue(4))
                End While
                objReader = Nothing
                objSQLConn.Close()
                objSQLConn = Nothing
                .Col = enumGrid.col_Item_Code
                .Col2 = enumGrid.col_Maruti_PO
                .Row = 1
                .Row2 = .MaxRows
                .BlockMode = True
                .Lock = True
                .BlockMode = False
            End With
        Else
            MsgBox("Data Not Found For Selected Criteria.", MsgBoxStyle.Information, ResolveResString(100))
        End If
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        objReader = Nothing
        If objSQLConn.State = ConnectionState.Open Then
            objSQLConn.Close()
        End If
        objSQLConn = Nothing
    End Sub
    Private Sub optAllDisplay_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optAllDisplay.CheckedChanged
        fpsSpread.MaxRows = 0
    End Sub
    Private Sub optMapped_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optMapped.CheckedChanged
        fpsSpread.MaxRows = 0
    End Sub
    Private Sub optUnMapped_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optUnMapped.CheckedChanged
        fpsSpread.MaxRows = 0
    End Sub
End Class