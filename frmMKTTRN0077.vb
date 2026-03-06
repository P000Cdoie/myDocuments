Option Strict Off
Option Explicit On
Imports System.Data.SqlClient
Imports System.Data
Public Class frmMKTTRN0077
#Region "Comments"
    '***************************************************************************************
    'Copyright       : MIND Ltd.
    'Module          : frmMKTTRN0077 - Intra Material Transfer Note Cancellation
    'Description     : Intra Material Transfer Note Cancellation
    '------------------------------------------------------------------------------
    'REVISED BY     :   VINOD SINGH
    'REVISION DATE  :   22 DEC 2014
    'ISSUE ID       :   10729940-DUPLICATE ROWS IN IMTN CANCELLATION
    '***************************************************************************************
#End Region
#Region "Form level variable Declarations"
    Dim MINTFORMINDEX As Integer
    Dim SqlCmd As SqlCommand
    Dim SqlDR As SqlDataReader
    Dim SqlAdp As SqlDataAdapter
    Dim mStrSql As String = String.Empty
    Dim intRowCount As Integer = 0
    Dim mStrMSG As String = String.Empty
    Private Enum ENUMDETAILS
        VAL_ITEMCODE = 1
        VAL_CUSTPARTNO = 2
        VAL_RATE = 3
        VAL_QUANTITY = 4
        VAL_CUSTPARTNODESC = 5
        VAL_CURRSTOCK = 6
        VAL_LOCATION = 7
        VAL_ITEMGRPCODE = 8
        VAL_ITEMGLCODE = 9
        VAL_ITEMSLCODE = 10
    End Enum
#End Region
#Region "FormEvents"
    Private Sub frmMKTTRN0077_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            mdifrmMain.CheckFormName = MINTFORMINDEX
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub frmMKTTRN0077_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub
    Private Sub frmMKTTRN0077_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.MdiParent = prjMPower.mdifrmMain
            MINTFORMINDEX = mdifrmMain.AddFormNameToWindowList(ctlHeader.Tag)
            Call FitToClient(Me, Me.GrpIntra, ctlHeader, CmdButtonGrp, 500)
            DTPIMNT.Format = DateTimePickerFormat.Custom
            DTPIMNT.CustomFormat = gstrDateFormat
            DTPIMNT.Value = GetServerDate()
            DTPIMNT.Enabled = False
            DtpBasicDueDate.Format = DateTimePickerFormat.Custom
            DtpBasicDueDate.CustomFormat = gstrDateFormat
            DtpBasicDueDate.Value = GetServerDate()
            DtpBasicDueDate.Enabled = False
            DtpBasicDueDate.Visible = False
            DtpBasicPayDate.Format = DateTimePickerFormat.Custom
            DtpBasicPayDate.CustomFormat = gstrDateFormat
            DtpBasicPayDate.Value = GetServerDate()
            DtpBasicPayDate.Enabled = False
            DtpBasicPayDate.Visible = False
            CmdIMNTNo.Image = My.Resources.ico111.ToBitmap
            lblUnitCode.Text = gstrUNITID
            lblUnitDesc.Text = gstrUNITDESC
            InitializeSpreed()
            AddBlankRow()
            RefreshForm()
            CmdButtonGrp.Caption(0) = "New"
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
#End Region
#Region "Sub Routines"
    Private Sub InitializeSpreed()
        Try
            With Me.SprIMNTGrid
                .MaxRows = 0
                .MaxCols = 10
                .Row = 0 : .Col = ENUMDETAILS.VAL_ITEMCODE : .Text = "Item Code" : .set_ColWidth(ENUMDETAILS.VAL_ITEMCODE, 1000)
                .Row = 0 : .Col = ENUMDETAILS.VAL_CUSTPARTNO : .Text = "Cust Part No" : .set_ColWidth(ENUMDETAILS.VAL_CUSTPARTNO, 1500)
                .Row = 0 : .Col = ENUMDETAILS.VAL_RATE : .Text = "Rate (Per Unit)" : .set_ColWidth(ENUMDETAILS.VAL_RATE, 1000)
                .Row = 0 : .Col = ENUMDETAILS.VAL_QUANTITY : .Text = "Quantity" : .set_ColWidth(ENUMDETAILS.VAL_QUANTITY, 1000)
                .Row = 0 : .Col = ENUMDETAILS.VAL_CUSTPARTNODESC : .Text = "Cust Part Desc" : .set_ColWidth(ENUMDETAILS.VAL_CUSTPARTNODESC, 2000)
                .Row = 0 : .Col = ENUMDETAILS.VAL_CURRSTOCK : .Text = "Current Stock" : .set_ColWidth(ENUMDETAILS.VAL_CURRSTOCK, 1500)
                .Row = 0 : .Col = ENUMDETAILS.VAL_LOCATION : .Text = "Location Code" : .set_ColWidth(ENUMDETAILS.VAL_LOCATION, 1500)
                .Row = 0 : .Col = ENUMDETAILS.VAL_ITEMGRPCODE : .Text = "Item Inv Code" : .set_ColWidth(ENUMDETAILS.VAL_ITEMGRPCODE, 1500)
                .Row = 0 : .Col = ENUMDETAILS.VAL_ITEMGLCODE : .Text = "GL Code" : .set_ColWidth(ENUMDETAILS.VAL_ITEMGLCODE, 1500)
                .Row = 0 : .Col = ENUMDETAILS.VAL_ITEMSLCODE : .Text = "SL Code" : .set_ColWidth(ENUMDETAILS.VAL_ITEMSLCODE, 1500)
                .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
    Public Sub AddBlankRow()
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            With Me.SprIMNTGrid
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows : .Col = ENUMDETAILS.VAL_ITEMCODE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                .Row = .MaxRows : .Col = ENUMDETAILS.VAL_CUSTPARTNO : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .Row = .MaxRows : .Col = ENUMDETAILS.VAL_RATE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Row = .MaxRows : .Col = ENUMDETAILS.VAL_QUANTITY : .CellType = FPSpreadADO.CellTypeConstants.CellTypeNumber : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Row = .MaxRows : .Col = ENUMDETAILS.VAL_CUSTPARTNODESC : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .Row = .MaxRows : .Col = ENUMDETAILS.VAL_CURRSTOCK : .CellType = FPSpreadADO.CellTypeConstants.CellTypeNumber : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .Row = .MaxRows : .Col = ENUMDETAILS.VAL_LOCATION : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                .Row = .MaxRows : .Col = ENUMDETAILS.VAL_ITEMGRPCODE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                .Row = .MaxRows : .Col = ENUMDETAILS.VAL_ITEMGLCODE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                .Row = .MaxRows : .Col = ENUMDETAILS.VAL_ITEMSLCODE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub RefreshForm()
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            txtIMNTNo.Text = "" : txtIMNTNo.Enabled = False : txtIMNTNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            CmdIMNTNo.Enabled = False
            DTPIMNT.Enabled = False : DTPIMNT.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtCustomerCode.Text = ""
            lblCustomerDesc.Text = ""
            txtSoNo.Text = ""
            lblAmendementNo.Text = ""
            lblARDrNote.Text = ""
            SprIMNTGrid.Enabled = True : SprIMNTGrid.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            SprIMNTGrid.MaxRows = 0
            lblReceivingUnitCode.Text = ""
            lblReceivingUnitName.Text = ""
            txtCancelRemark.Enabled = False : txtCancelRemark.Text = "" : txtCancelRemark.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)

            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub ShowHelp(ByVal pstrQuery As String)
        Try
            Dim StrHelp() As String
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
            With IMNTHelp
                .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
                .ConnectAsUser = gstrCONNECTIONUSER
                .ConnectThroughDSN = gstrCONNECTIONDSN
                .ConnectWithPWD = gstrCONNECTIONPASSWORD
            End With
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            StrHelp = IMNTHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, pstrQuery, ResolveResString(100), 1)
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            If UBound(StrHelp) <> -1 Then
                If StrHelp(0) <> "0" Then
                    txtIMNTNo.Text = StrHelp(0)
                    DTPIMNT.Value = StrHelp(1)

                Else
                    MessageBox.Show("Record not found.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                End If

            End If
        Catch EX As Exception
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            MessageBox.Show(EX.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
#End Region
#Region "Form Control Events"
    Private Sub CmdIMNTNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdIMNTNo.Click
        Try
            mStrSql = "SELECT DOC_NO,DOC_DATE FROM IMTN_HDR WHERE UNIT_CODE = '" & gstrUNITID & "' AND NOT AUTH_CODE IS NULL AND CANCEL = 0 and DOC_DATE = CONVERT(VARCHAR(12),GETDATE(),106) "
            Call ShowHelp(mStrSql)
            If txtIMNTNo.Text.ToString.Trim.Length > 0 Then
                If PopulateData() = False Then
                    MessageBox.Show("Record not found", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub txtIMNTNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtIMNTNo.KeyDown
        Try
            Dim KeyCode As Short = e.KeyCode
            Dim Shift As Short = e.KeyData \ &H10000
            If Shift <> 0 Then Exit Sub
            If KeyCode = System.Windows.Forms.Keys.F1 Then
                If sender.Equals(txtIMNTNo) Then
                    Call CmdIMNTNo_Click(CmdIMNTNo, New System.EventArgs())
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub txtIMNTNo_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtIMNTNo.Validating
        Try
            If sender.Equals(txtIMNTNo) Then
                If txtIMNTNo.Text.ToString.Trim.Length > 0 Then
                    If PopulateData() = False Then
                        MessageBox.Show("Records not found.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
                        txtIMNTNo.Text = ""
                        txtIMNTNo.Focus()
                    Else
                        CmdButtonGrp.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
        End Try
    End Sub
    Private Sub txtIMNTNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtIMNTNo.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Try
            If sender.Equals(txtIMNTNo) Or sender.Equals(txtCustomerCode) Or sender.Equals(txtSoNo) Then
                If KeyAscii = 13 Then
                    System.Windows.Forms.SendKeys.Send("{tab}")
                ElseIf KeyAscii = 187 Or KeyAscii = 166 Or KeyAscii = 164 Or KeyAscii = 172 Then
                    KeyAscii = 0
                End If
                If Not (e.KeyChar >= "a" And e.KeyChar <= "z") And Not (e.KeyChar >= "A" And e.KeyChar <= "Z") And Not (e.KeyChar >= "0" And e.KeyChar <= "9") And Not KeyAscii = 8 Then
                    e.KeyChar = ""
                    KeyAscii = 0
                End If
            End If

            GoTo EventExitSub

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
        End Try
EventExitSub:
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtIMNTNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIMNTNo.TextChanged
        Try
            If txtIMNTNo.Text.ToString.Trim.Length = 0 Then
                Call RefreshForm()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
    Private Sub CmdButtonGrp_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtnEditGrp.ButtonClickEventArgs) Handles CmdButtonGrp.ButtonClick
        Try
            Select Case e.Button

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                    txtIMNTNo.Enabled = True : txtIMNTNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    CmdIMNTNo.Enabled = True
                    txtCancelRemark.Enabled = True : txtCancelRemark.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    txtCustomerCode.Focus()
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                    If Validate_data() = True Then
                        If SaveData() = False Then
                            MessageBox.Show("Due to some error record could not saved please contact to administrator.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                        Else
                            MessageBox.Show("Intra material transfer note has been canceled.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                            RefreshForm()
                            CmdButtonGrp.Revert()
                        End If
                    Else
                        MessageBox.Show(mStrMSG, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                    End If

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                    RefreshForm()
                    CmdButtonGrp.Revert()
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                    Me.Close()
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
#End Region
#Region "Functions"
    Private Function SaveData() As Boolean
        Try
            Try
                SqlCmd = New SqlCommand
                With SqlCmd
                    .Connection = SqlConnectionclass.GetConnection()
                    .CommandType = CommandType.StoredProcedure
                    .CommandText = "USP_CANCEL_INTRAMATERIALTRANSFERNOTE"
                    .CommandTimeout = 0
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID.Trim
                    .Parameters.Add("@Customer_Code", SqlDbType.VarChar, 10).Value = txtCustomerCode.Text.Trim
                    .Parameters.Add("@DOC_NO", SqlDbType.VarChar, 15).Value = txtIMNTNo.Text.ToString.Trim
                    .Parameters.Add("@CANCELREMARKS", SqlDbType.VarChar, 250).Value = txtCancelRemark.Text.Trim
                    .Parameters.Add("@MPUSER", SqlDbType.VarChar, 16).Value = mP_User.Trim
                    SqlCmd.ExecuteNonQuery()
                End With
                SqlCmd.Dispose()
                Return True
            Catch SqlEx As SqlException
                MsgBox(SqlEx.Message, MsgBoxStyle.Critical, ResolveResString(100))
                Return False
            Finally
                If IsNothing(SqlCmd) = False Then
                    If SqlCmd.Connection.State = ConnectionState.Open Then
                        SqlCmd.Connection.Close()
                    End If
                    SqlCmd.Dispose()
                End If
            End Try
            Return True
        Catch ex As SqlException
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
            Return False
        Finally
        End Try

    End Function
    Private Function PopulateData() As Boolean
        Dim intLoopcounter As Integer = 1
        Dim sqlAdp As SqlDataAdapter = Nothing
        Dim sqlDataSet As DataSet = Nothing
        Dim sqlCmd As New SqlCommand
        Dim dtIMTN As New DataTable
        Try
            Call InitializeSpreed()
            'ISSUE ID 10729940
            'mStrSql = "select H.RECV_UNIT_CODE,h.Doc_Date,h.Customer_Code,h.Cust_Ref,h.Amendment_No,h.AR_DrNote,h.AR_DrDate, " & _
            '        " h.Booked_FinanceValue,d.Item_Code,d.Cust_DrgNo,d.Item_rate,d.MTN_Qty,d.Location_Code, " & _
            '        " INV.invGld_invGLGrpId,INV.invGld_glCode,INV.invGld_slCode,C.CUST_NAME,CUST.Cust_Drg_Desc, h.Cancel,h.CancelRemarks " & _
            '        " from IMTN_Hdr h inner join IMTN_Dtl d on h.Unit_Code = d.Unit_code and h.Doc_No = d.Doc_no  " & _
            '        " and h.Doc_Type = d.Doc_Type inner join item_mst i  " & _
            '        " on d.Item_Code = i.Item_Code and d.Unit_code = i.UNIT_CODE " & _
            '        " inner join fin_invGLGrpDtl inv on i.GlGrp_code = inv.invGld_invGLGrpId  " & _
            '        " and i.UNIT_CODE = inv.UNIT_CODE " & _
            '        " INNER JOIN CUSTOMER_MST C ON h.Customer_Code = C.CUSTOMER_CODE AND H.UNIT_CODE = C.UNIT_CODE " & _
            '        " INNER JOIN Cust_ord_dtl CUST ON D.ITEM_CODE = CUST.ITEM_CODE AND D.Cust_DrgNo = CUST.Cust_DrgNo " & _
            '        " AND H.CUSTOMER_CODE = CUST.ACCOUNT_CODE AND  h.Cust_Ref = cust.Cust_Ref AND H.UNIT_CODE = CUST.UNIT_CODE " & _
            '        " where inv.invGld_prpsCode = 'StockTrans' " & _
            '        " AND H.DOC_NO = '" & txtIMNTNo.Text.ToString.Trim & "' " & _
            '        " AND H.UNIT_CODE = '" & gstrUNITID.Trim & "'  AND H.CANCEL = 0 "
            With sqlCmd
                .CommandText = "USP_IMTN_CANCELLATION_ITEM_DETAIL"
                .CommandType = CommandType.StoredProcedure
                .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                .Parameters.Add("@IMTN_NO", SqlDbType.Int).Value = Val(txtIMNTNo.Text.Trim)
                dtIMTN = SqlConnectionclass.GetDataTable(sqlCmd)
            End With
            'sqlAdp = New SqlDataAdapter(mStrSql, SqlConnectionclass.GetConnection)
            'sqlDataSet = New DataSet
            'sqlAdp.Fill(sqlDataSet)
            With SprIMNTGrid
                If dtIMTN.Rows.Count > 0 Then
                    DTPIMNT.Value = dtIMTN.Rows(0)("Doc_Date").ToString
                    txtCustomerCode.Text = dtIMTN.Rows(0)("Customer_Code").ToString
                    lblCustomerDesc.Text = dtIMTN.Rows(0)("CUST_NAME").ToString
                    txtSoNo.Text = dtIMTN.Rows(0)("Cust_Ref").ToString
                    lblAmendementNo.Text = dtIMTN.Rows(0)("Amendment_No").ToString
                    lblARDrNote.Text = dtIMTN.Rows(0)("AR_DrNote").ToString
                    lblReceivingUnitCode.Text = dtIMTN.Rows(0)("RECV_UNIT_CODE").ToString
                    lblReceivingUnitName.Text = FetchUnitName(lblReceivingUnitCode.Text)
                    If dtIMTN.Rows(0)("cancel") = True Then
                        'ChkCancel.Checked = True : ChkCancel.Enabled = False
                        txtCancelRemark.Text = dtIMTN.Rows(0)("CancelRemarks").ToString : txtCancelRemark.Enabled = False
                    Else
                        'ChkCancel.Checked = False : ChkCancel.Enabled = True
                        txtCancelRemark.Text = "" : txtCancelRemark.Enabled = True
                    End If

                    For intRowCount = 0 To dtIMTN.Rows.Count - 1
                        Call AddBlankRow()
                        .SetText(ENUMDETAILS.VAL_ITEMCODE, intLoopcounter, dtIMTN.Rows(intRowCount)("Item_Code").ToString)
                        .SetText(ENUMDETAILS.VAL_CUSTPARTNO, intLoopcounter, dtIMTN.Rows(intRowCount)("Cust_DrgNo").ToString)
                        .SetText(ENUMDETAILS.VAL_RATE, intLoopcounter, Convert.ToDecimal(dtIMTN.Rows(intRowCount)("Item_rate")).ToString)
                        .SetText(ENUMDETAILS.VAL_QUANTITY, intLoopcounter, Convert.ToDecimal(dtIMTN.Rows(intRowCount)("MTN_Qty")).ToString)
                        .SetText(ENUMDETAILS.VAL_CUSTPARTNODESC, intLoopcounter, dtIMTN.Rows(intRowCount)("Cust_Drg_Desc").ToString())
                        .SetText(ENUMDETAILS.VAL_CURRSTOCK, intLoopcounter, Convert.ToString(GetStockInHand(dtIMTN.Rows(intRowCount)("Location_Code").ToString, dtIMTN.Rows(intRowCount)("Item_Code").ToString)))
                        .SetText(ENUMDETAILS.VAL_LOCATION, intLoopcounter, dtIMTN.Rows(intRowCount)("Location_Code").ToString)
                        .SetText(ENUMDETAILS.VAL_ITEMGRPCODE, intLoopcounter, dtIMTN.Rows(intRowCount)("invGld_invGLGrpId").ToString)
                        .SetText(ENUMDETAILS.VAL_ITEMGLCODE, intLoopcounter, dtIMTN.Rows(intRowCount)("invGld_glCode").ToString)
                        .SetText(ENUMDETAILS.VAL_ITEMSLCODE, intLoopcounter, dtIMTN.Rows(intRowCount)("invGld_SlCode").ToString)
                        intLoopcounter = intLoopcounter + 1
                        PopulateData = True
                    Next
                    SprIMNTGrid.Enabled = True
                Else
                    Return False
                End If
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
            Return False
            If IsNothing(sqlCmd) = False Then sqlCmd.Dispose()
            If IsNothing(dtIMTN) = False Then dtIMTN.Dispose()
        Finally
        End Try
    End Function
    Private Function FetchUnitName(ByVal pstrUnitID As String) As String
        Try
            mStrSql = "SELECT UNT_UNITNAME  FROM GEN_UNITMASTER WHERE Unt_CodeID='" + pstrUnitID + "' "
            Return Convert.ToString(SqlConnectionclass.ExecuteScalar(SqlConnectionclass.GetConnection, CommandType.Text, mStrSql))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, ResolveResString(100))
            Return String.Empty
        End Try
    End Function
    Private Function Validate_data() As Boolean
        Try
            Dim intCtr As Integer = 0

            If txtIMNTNo.Text.Trim.Length = 0 Then
                mStrMSG = "Please select To Intra Material Transfer Note first."
                Return False
            End If
            mStrMSG = ""
            'If ChkCancel.Checked = False Then
            '    mStrMSG = "Please check cancel check box."
            '    Return False
            'End If
            If txtCancelRemark.Text.Trim.Length = 0 Then
                mStrMSG = "Please enter cancel remarks."
                Return False
            End If
            If mStrMSG.Trim.Length = 0 Then
                Return True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
            Return False
        End Try
    End Function
#End Region
End Class