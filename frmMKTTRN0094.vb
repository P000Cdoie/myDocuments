'******************************************************************************************************************************************************
'COPYRIGHT       : MIND LTD.
'MODULE          : FRMMKTTRN0094 - FTS FG PICK LIST GENERATION 
'AUTHOR          : PRASHANT RAJPAL

Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class frmMKTTRN0094
    Inherits System.Windows.Forms.Form
    Private mlngFormTag As Integer
    Private mServerDate As String
    Dim mstrUpdateDocumentNoSQL As String
    Dim mblnsalesorderpicklist As Boolean
    Private Enum EnmGridCol
        Mark = 1
        ItemCode = 2
        DrawingNo = 3
        Drg_Desc = 4
        Salesorder = 5
        AmendmentNo = 6
        BalQuantity = 7
        CurrentStock = 8
        ScheduleQuantity = 9
        Pending_WIPPick_Currentdate = 10
        Quantity_to_be_Invoice = 11
    End Enum
    Private Sub cmdDocCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDocCode.Click
        Dim StrDocHelp As String
        Try
            StrDocHelp = ShowList(0, Len(txtDocCode.Text), txtDocCode.Text, "Doc_No", "" & DateColumnNameInShowList("Trans_date") & " as Trans_Date", "FTS_FG_PICKLIST_INV_HDR", " ", "FTS INVOICE PICKLIST SERIES ")
            If StrDocHelp = "-1" Then

                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                With Me.txtDocCode
                    .Enabled = True
                    .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    .Focus()
                End With
                cmdDocCode.Enabled = True

            ElseIf StrDocHelp = "" Then
                Me.txtDocCode.Focus()
            Else
                Me.txtDocCode.Text = StrDocHelp
                txtDocCode_Validating(txtDocCode, New System.ComponentModel.CancelEventArgs(False))
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub ctlHeader_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ctlHeader.Click
        Try
            Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub FRMMKTTRN0082_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        Try
            mdifrmMain.CheckFormName = mlngFormTag
            frmModules.NodeFontBold(Me.Tag) = True
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub FRMMKTTRN0082_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        Try
            frmModules.NodeFontBold(Me.Tag) = False
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub FRMMKTTRN0082_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown

        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Try
            If (KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0) Then Call ctlHeader_ClickEvent(ctlHeader, New System.EventArgs())
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub
    Private Sub FRMMKTTRN0082_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim strQry As String

        Try
            mlngFormTag = mdifrmMain.AddFormNameToWindowList(Me.ctlHeader.Tag)
            Call FitToClient(Me, fraMain, ctlHeader, CmdBUTTON, 400)
            CmdBUTTON.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(fraMain.Left) + (VB6.PixelsToTwipsX(fraMain.Width) / 4))
            Call EnableControls(False, Me, True)
            SpItemDetails.Enabled = True
            mServerDate = getDateForDB(GetServerDate())
            Call InitializeControls()
            Call AddColumnsInSpread()
            txtunitcode.Text = gstrUNITID
            With txtunitcode
                .Enabled = False
                .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            End With
            '10623079
            strQry = "Select ALLOW_FTS_SALESORDER_PICKLIST from sales_parameter where unit_code='" & gstrUNITID & "'"
            mblnsalesorderpicklist = SqlConnectionclass.ExecuteScalar(strQry).ToString()
            '10623079
            If mblnsalesorderpicklist = True Then
                txtSalesOrder.Visible = True
                txtamendment.Visible = True
                cmdsono.Visible = True
                lblsalesorder.Visible = True
            Else
                txtSalesOrder.Visible = False
                txtamendment.Visible = False
                cmdsono.Visible = False
                lblsalesorder.Visible = False
            End If
            '10623079
            With txtinvoicetype
                .Enabled = False
                .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            End With

            With txtsubtype
                .Enabled = False
                .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            End With
            txtsearch.Enabled = True
            txtsearch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)

            With txtDocCode
                .Enabled = True
                .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                .Focus()
            End With

            With Me.txtinvoicetype
                .Enabled = False
                .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                .Focus()
            End With
            cmdDocCode.Enabled = True
            gblnCancelUnload = False

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub InitializeControls()
        '*************************************************************************************************************************************************************
        'PURPOSE     : TO INITALIZE CONTROLS :UNIT CODE , DATE IT SHOULD BE OF TODAYS DATE ,BY DEFAULT VIEW MODE WHERE ONLY DOCUMENT FIELD IS ENABLED
        '*************************************************************************************************************************************************************

        Try
            With txtunitcode
                .Enabled = False
                .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            End With

            lblDisplayDate.Text = VB6.Format(mServerDate, gstrDateFormat)
            With CmdBUTTON
                .Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
            End With
            Call CmdBUTTON.ShowButtons(True, False, True, False)
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub FRMMKTTRN0082_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason

        Try
            Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
            If UnloadMode >= 0 And UnloadMode <= 5 Then
                If CmdBUTTON.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    enmValue = ConfirmWindow(10055, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Or enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then

                        ElseIf enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Then
                            gblnCancelUnload = False : gblnFormAddEdit = False
                        End If
                    Else

                        gblnCancelUnload = True : gblnFormAddEdit = True
                        Me.ActiveControl.Focus()
                    End If
                Else

                    Exit Sub
                End If
            End If
            If gblnCancelUnload = True Then eventArgs.Cancel = True

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub FRMMKTTRN0082_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            Me.Dispose()
            frmModules.NodeFontBold(Me.Tag) = False
            mdifrmMain.RemoveFormNameFromWindowList = mlngFormTag

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub SpItemDetails_Change(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ChangeEvent) Handles SpItemDetails.Change
        Try
            Select Case e.col

                Case EnmGridCol.Mark
                    With SpItemDetails
                        .Row = e.row
                        .Col = EnmGridCol.Mark
                        If CBool(.Value) Then
                            .Col = EnmGridCol.Quantity_to_be_Invoice : .Col2 = EnmGridCol.Quantity_to_be_Invoice : .BlockMode = True : .Col = EnmGridCol.Quantity_to_be_Invoice
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .BlockMode = False
                        Else
                            .Col = EnmGridCol.Quantity_to_be_Invoice : .Col2 = EnmGridCol.Quantity_to_be_Invoice : .BlockMode = True : .Col = EnmGridCol.Quantity_to_be_Invoice
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .BlockMode = False
                        End If
                    End With
            End Select

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub
    Private Sub Spitemdetails_DblClick(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_DblClickEvent) Handles SpItemDetails.DblClick
        Try
            With SpItemDetails
                If (e.col = 0 And e.row > 0) And (CmdBUTTON.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD) Then
                    .Row = e.row
                    .Action = FPSpreadADO.ActionConstants.ActionDeleteRow
                    .MaxRows = .MaxRows - 1
                End If
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub txtDocCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDocCode.TextChanged
        Try
            Select Case CmdBUTTON.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    If (txtDocCode.Text.Trim.Length = 0) Then
                        Call RefreshControls()
                    End If
            End Select

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub txtDocCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDocCode.Enter
        Try
            With Me.txtDocCode
                .SelectionStart = 0
                .SelectionLength = Len(txtDocCode.Text)
            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub txtDocCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtDocCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Try
            If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
                Call cmdDocCode_Click(cmdDocCode, New System.EventArgs())
            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub txtunitcode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtunitcode.TextChanged

        Try
            Select Case CmdBUTTON.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    If (txtunitcode.Text.Trim.Length = 0) Then
                        Call RefreshControls()
                        lblunitdesc.Text = String.Empty
                        With txtDocCode
                            .Clear()
                            .Enabled = False
                            .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        End With
                        cmdDocCode.Enabled = False
                    End If
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    Call RefreshControls()
                    lblunitdesc.Text = String.Empty
            End Select

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub RefreshControls()
        '******************************************************
        'PURPOSE     : TO REFRESH THE CONTROLS WITH BLANK DATA
        '******************************************************
        Try
            With SpItemDetails
                .MaxRows = 0
            End With
            lblDisplayDate.Text = VB6.Format(mServerDate, gstrDateFormat)
            txtCustomerCode.Text = String.Empty
            lblCustomerdesc.Text = String.Empty
            lblsubtypedescription.Text = String.Empty
            lblInvoicedescription.Text = String.Empty
            txtDocCode.Text = String.Empty
            txtSalesOrder.Text = String.Empty


            With CmdBUTTON
                .Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub txtunitcode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtunitcode.Enter
        Try

            With Me.txtunitcode
                .SelectionStart = 0
                .SelectionLength = Len(txtunitcode.Text)
            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Function InsertIntoPickListTable() As Boolean
        '**********************************************************************************************************************************
        'FUNCTION :MAILY USED TO INSERT THE DATA IN WIP_FG_PICKLIST HDR AND DTL, AND ALS UPDATE SERIES FROM DOCUMENTTYPE_MST WITH DOC TYPE 9990
        '**********************************************************************************************************************************
        InsertIntoPickListTable = False
        Dim strQry As String
        Dim sqlConn As New SqlConnection
        Dim sqlCmd As SqlCommand
        Dim sqlTrans As SqlTransaction
        Dim sqlRdr As SqlDataReader
        Dim strinvoicetype As String
        Dim intLoopCounter As Integer
        Dim Documentno As String
        Dim StrPartCode As String
        Dim StrSalesorder As String
        Dim StrAmendmentno As String
        Dim stritemcode As String
        Dim dblquantity As Double
        Dim blnTran As Boolean
        Dim StrDrgdesc As String
        Dim SALESORDERLIST As String
        Dim FinanceYearNotation As String
        Dim GenerateDocumentno As String
        strinvoicetype = "INV"
        Try
            sqlConn = SqlConnectionclass.GetConnection
            sqlTrans = sqlConn.BeginTransaction
            blnTran = True

            sqlCmd = New SqlCommand()
            sqlCmd.Connection = sqlConn
            sqlCmd.Transaction = sqlTrans

            strQry = "SELECT Fin_Year_Notation from Financial_Year_tb where  UNIT_CODE='" & gstrUNITID & "' AND '" & VB6.Format(lblDisplayDate.Text, "DD MMM YYYY") & "'  BETWEEN FIN_START_DATE AND FIN_END_DATE "
            FinanceYearNotation = SqlConnectionclass.ExecuteScalar(strQry).ToString()

            strQry = "SELECT ISNULL(CURRENT_NO,0)+1 AS CURRENT_NO,FIN_START_DATE,FIN_END_DATE FROM DOCUMENTTYPE_MST WHERE DOC_TYPE=9995 AND '" & VB6.Format(lblDisplayDate.Text, "DD MMM YYYY") & "'  BETWEEN FIN_START_DATE AND FIN_END_DATE AND UNIT_CODE = '" & gstrUNITID & "'"
            sqlCmd.CommandType = CommandType.Text
            sqlCmd.CommandText = strQry

            sqlRdr = sqlCmd.ExecuteReader
            If sqlRdr.HasRows = True Then
                While sqlRdr.Read
                    Documentno = sqlRdr("CURRENT_NO").ToString
                End While
            Else
                If sqlRdr.IsClosed = False Then sqlRdr.Close()
                sqlTrans.Rollback()
                blnTran = False
                If sqlConn.State = ConnectionState.Open Then sqlConn.Close()
                MsgBox("Document number cannot be generated. Document series not defined", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                Exit Function
            End If

            GenerateDocumentno = CStr(FinanceYearNotation) + CStr(Documentno)

            If sqlRdr.IsClosed = False Then sqlRdr.Close()

            strQry = "INSERT INTO FTS_FG_PICKLIST_INV_HDR(UNIT_CODE,DOC_TYPE,DOC_NO,INVOICE_TYPE,CUSTOMER_CODE,TRANS_DATE,TRANS_TIME,ENT_DT,ENT_USERID,UPD_DT,UPD_USERID )"
            strQry += " VALUES ('" & gstrUNITID & "',9995 " & ",'" & GenerateDocumentno & "','" & strinvoicetype & "','" & txtCustomerCode.Text.Trim & "',"
            strQry += "CONVERT(DATETIME, CONVERT(VARCHAR(11), GETDATE(), 106), 106) ,SUBSTRING(CONVERT(VARCHAR(20),GETDATE()),13,LEN(GETDATE())),"
            strQry += "GETDATE(),'" & Trim(mP_User) & "', GETDATE(),'" & Trim(mP_User) & "')"

            sqlCmd.CommandText = strQry
            sqlCmd.ExecuteNonQuery()

            With SpItemDetails
                For intLoopCounter = 1 To .MaxRows
                    .Row = intLoopCounter
                    .Col = EnmGridCol.Mark

                    If CBool(.Value) Then
                        .Row = intLoopCounter
                        .Col = EnmGridCol.DrawingNo
                        StrPartCode = .Text.Trim
                        .Col = EnmGridCol.Drg_Desc
                        StrDrgdesc = .Text.Trim
                        .Col = EnmGridCol.ItemCode
                        stritemcode = .Text.Trim
                        .Col = EnmGridCol.Quantity_to_be_Invoice
                        dblquantity = Val(.Text)
                        .Row = intLoopCounter
                        .Col = EnmGridCol.Salesorder
                        StrSalesorder = .Text.Trim
                        .Col = EnmGridCol.AmendmentNo
                        StrAmendmentno = .Text.Trim
                        strQry = "INSERT INTO FTS_FG_PICKLIST_INV_DTL(UNIT_CODE,DOC_TYPE,DOC_NO,CUST_ITEM_CODE,CUST_ITEM_DESC,ITEM_CODE,QUANTITY,SO_NO,AMENDMENT_NO,ENT_DT,ENT_USERID)"
                        strQry += "VALUES ('" & gstrUNITID & "',9995 " & ",'" & GenerateDocumentno & "','" & StrPartCode & "','" & StrDrgdesc & "','" & stritemcode & "'," & dblquantity
                        strQry += ",'" & StrSalesorder & "','" & StrAmendmentno & "',  GETDATE(),'" & Trim(mP_User) & "')"
                        sqlCmd.CommandText = strQry
                        sqlCmd.ExecuteNonQuery()
                    End If
                Next
            End With

            strQry = "UPDATE DOCUMENTTYPE_MST SET CURRENT_NO=" & Val(CStr(Documentno)) & " WHERE UNIT_CODE='" & gstrUNITID & "'AND DOC_TYPE=9995 AND  GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE  "
            sqlCmd.CommandType = CommandType.Text
            sqlCmd.CommandText = strQry
            sqlCmd.ExecuteNonQuery()

            sqlTrans.Commit()
            blnTran = False
            sqlConn.Close()
            InsertIntoPickListTable = True
            txtDocCode.Text = Documentno
            MsgBox("Record successfully saved with new Document no [" & GenerateDocumentno & "]", MsgBoxStyle.Information, ResolveResString(100))

        Catch ex As Exception
            If blnTran Then sqlTrans.Rollback()
            If sqlConn.State = ConnectionState.Open Then sqlConn.Close()
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        Finally

            If IsNothing(sqlTrans) = False Then
                If blnTran = True Then
                    sqlTrans.Rollback()
                    blnTran = False
                End If
            End If
            If IsNothing(sqlCmd) = False Then
                sqlCmd.Dispose()
            End If
            If IsNothing(sqlConn) = False Then
                If sqlConn.State = ConnectionState.Open Then sqlConn.Close()
            End If
        End Try
    End Function
    Private Sub Spitemdetails_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles SpItemDetails.KeyPressEvent
        Try
            Select Case CmdBUTTON.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    Select Case e.keyAscii
                        Case 34, 39, 96
                            e.keyAscii = 0
                    End Select
            End Select

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub cmdcustcodehelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdcustcodehelp.Click
        Dim strCustHelp As String
        Try
            If txtinvoicetype.Text.Length <= 0 Then
                MsgBox("Please Select First Invoice Type ", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                txtinvoicetype.Focus()
                Exit Sub
            End If

            Select Case CmdBUTTON.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD

                    strCustHelp = ShowList(0, Len(txtunitcode.Text), , "CUSTOMER_CODE", "CUST_NAME", "CUSTOMER_MST", " AND UNIT_CODE ='" & gstrUNITID & "' AND FTS_DISPATCHADVICE_CUSTOMER =1 ", "CUSTOMER HELP", , , , , )
                    If strCustHelp = "-1" Then
                        txtCustomerCode.Focus()
                        Exit Sub

                    ElseIf strCustHelp = "" Then
                        txtCustomerCode.Focus()
                    Else
                        txtCustomerCode.Text = strCustHelp
                        lblCustomerdesc.Text = GetQryOutput("SELECT CUST_NAME FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & strCustHelp & "' AND FTS_DISPATCHADVICE_CUSTOMER =1")
                    End If
            End Select

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub txtCustomerCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.TextChanged
        Try
            With SpItemDetails
                .MaxRows = 0
                If Me.txtCustomerCode.Text.Length = 0 Then Me.lblCustomerdesc.Text = String.Empty : Me.txtSalesOrder.Text = String.Empty
            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub txtCustomerCode_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
        Dim strQry As String
        Dim dtCustomer As New DataTable
        If txtCustomerCode.Text.Trim = "" Then Exit Sub

        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            If CmdBUTTON.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                strQry = "SELECT CUST_NAME FROM CUSTOMER_MST  "
                strQry += " WHERE CUSTOMER_CODE= '" & Me.txtCustomerCode.Text.Trim & "' and unit_code= '" & gstrUNITID & "' AND FTS_DISPATCHADVICE_CUSTOMER =1 "
            End If
            Call GetData(strQry, dtCustomer)
            If dtCustomer.Rows.Count = 0 Then
                MsgBox("Invalid Customer Code for FTS ", MsgBoxStyle.Information, ResolveResString(100))
                Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
                e.Cancel = True
                Exit Sub
            End If

            lblCustomerdesc.Text = dtCustomer.Rows(0).Item("cust_name")
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)

        Catch EX As Exception
            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub
    Private Sub GetData(ByVal strQry As String, ByRef dt As DataTable)
        Dim ConnString As String = "Server=" & gstrCONNECTIONSERVER & ";Database=" & gstrDatabaseName & ";user=" & gstrCONNECTIONUSER & ";password=" & gstrCONNECTIONPASSWORD & " "
        Dim oDataAdapter As SqlDataAdapter
        Try
            oDataAdapter = New SqlDataAdapter(strQry, ConnString)
            oDataAdapter.Fill(dt)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub txtinvoicetype_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtinvoicetype.TextChanged
        Try

            With SpItemDetails
                .MaxRows = 0
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub CmdDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdDetails.Click
        Try
            Call FILLSODETAILS()
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Public Function FILLSODETAILS() As Boolean
        '*******************************************************************************************************************************************************
        'PURPOSE     : TO FILL THE DETAIL DATA OF SELECTED CUSTOMER , DETAIL DATA MEANS PENDING SO AND PENDING SCHEDULE QUANTITY ITEM WISE AND ALSO STOCK AT 01B1
        'RETURN TYPE :BOOLEAN (TRUE MEANS ANY ITEM IS PENDING FOR INVOICE FOR TODAY AS PER SCHEDULE AND SO WISE )
        '********************************************************************************************************************************************************

        Dim sqlConn As New SqlConnection
        Dim sqlCmd As New SqlCommand
        Dim sqlTrans As SqlTransaction = Nothing
        Dim dblinvoiceQty As String
        Dim dblpendingqty As String
        Dim dblPendingpicklist_Today As String
        Dim strQry As String
        Dim blnTran As Boolean
        Dim oda As New SqlDataAdapter
        Dim ods As New DataSet
        Dim Stritemcode As String
        Dim StrCustDrgno As String
        Dim StrCustDrgdesc As String
        Dim StrCustref As String
        Dim StrAmendmenno As String
        Dim lngloop As Long
        Try
            If Me.txtCustomerCode.Text.Length <= 0 Then
                MsgBox("Please Select First Customer ", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                txtCustomerCode.Text = String.Empty
                txtCustomerCode.Focus()
                Exit Function
            End If
            If mblnsalesorderpicklist = True Then
                If Me.txtSalesOrder.Text.Length <= 0 Then
                    MsgBox("Please Select sales order ", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                    txtSalesOrder.Text = String.Empty
                    txtSalesOrder.Focus()
                    Exit Function
                End If
            End If

            FILLSODETAILS = True
            sqlConn = SqlConnectionclass.GetConnection
            'sqlTrans = sqlConn.BeginTransaction
            blnTran = True

            sqlCmd = New SqlCommand()
            sqlCmd.Connection = sqlConn
            'sqlCmd.Transaction = sqlTrans

            With sqlCmd
                .Parameters.Clear()
                .CommandType = CommandType.StoredProcedure
                .CommandText = "USP_FTS_PICKLIST_GENERATE_DATA"
                .Parameters.AddWithValue("@Unit_Code", gstrUNITID)
                .Parameters.AddWithValue("@CUST_CODE", txtCustomerCode.Text.Trim)
                .Parameters.AddWithValue("@INVTYPE", txtinvoicetype.Text.Trim)
                .Parameters.AddWithValue("@INVSUBTYPE", "F")
                .Parameters.AddWithValue("@DATE", VB6.Format(lblDisplayDate.Text))
                .Parameters.AddWithValue("@SONO", txtSalesOrder.Text.Trim)
                .Parameters.AddWithValue("@AMENDMENTNO", txtamendment.Text.Trim)
                .Parameters.Add("@MSG", SqlDbType.VarChar, 5000).Direction = ParameterDirection.Output
                .ExecuteNonQuery()

                Dim strMsg As String = String.Empty
                strMsg = Convert.ToString(.Parameters("@MSG").Value)
                If Len(strMsg) > 0 Then
                    MsgBox(strMsg, MsgBoxStyle.Exclamation, ResolveResString(100))
                    Exit Function
                End If
            End With
            oda.SelectCommand = sqlCmd
            oda.Fill(ods)
            If ods.Tables(0).Rows.Count <= 0 Then
                MessageBox.Show("No Data found ", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                Exit Function
            End If
            Call AddColumnsInSpread()

            With SpItemDetails
                For lngloop = 0 To ods.Tables(0).Rows.Count - 1
                    Call AddBlankRow()
                    Stritemcode = CStr(ods.Tables(0).Rows(lngloop).Item("ITEM_CODE"))
                    StrCustDrgno = CStr(ods.Tables(0).Rows(lngloop).Item("CUST_DRGNO"))
                    StrCustDrgdesc = CStr(ods.Tables(0).Rows(lngloop).Item("CUST_DRG_DESC"))
                    StrCustref = CStr(ods.Tables(0).Rows(lngloop).Item("CUST_REF"))
                    StrAmendmenno = CStr(ods.Tables(0).Rows(lngloop).Item("AMENDMENT_NO"))

                    Call .SetText(EnmGridCol.ItemCode, .MaxRows, Stritemcode)
                    Call .SetText(EnmGridCol.DrawingNo, .MaxRows, StrCustDrgno)
                    Call .SetText(EnmGridCol.Drg_Desc, .MaxRows, StrCustDrgdesc)
                    Call .SetText(EnmGridCol.Salesorder, .MaxRows, StrCustref)
                    Call .SetText(EnmGridCol.AmendmentNo, .MaxRows, StrAmendmenno)

                    If CInt(ods.Tables(0).Rows(lngloop).Item("BALANCE_QTY")) > 0 Then
                        Call .SetText(EnmGridCol.BalQuantity, .MaxRows, CInt(ods.Tables(0).Rows(lngloop).Item("BALANCE_QTY")))
                    Else
                        Call .SetText(EnmGridCol.BalQuantity, .MaxRows, 0)
                    End If
                    Call .SetText(EnmGridCol.CurrentStock, .MaxRows, CInt(ods.Tables(0).Rows(lngloop).Item("cur_bal")))
                    SqlConnectionclass.ExecuteReader("set dateformat 'dmy'")
                    strQry = "SELECT DBO.UDF_FTS_CURRENTMONTHINVOICE('" & gstrUNITID & "','" & txtCustomerCode.Text.Trim & "','" & StrCustDrgno & "','" & Stritemcode & "')"
                    dblinvoiceQty = SqlConnectionclass.ExecuteScalar(strQry).ToString()

                    strQry = "SELECT DBO.UDF_FTS_CURRENTDATE_PENDING_PICKLIST('" & gstrUNITID & "','" & txtCustomerCode.Text.Trim & "','" & StrCustDrgno & "','" & Stritemcode & "')"
                    dblPendingpicklist_Today = SqlConnectionclass.ExecuteScalar(strQry).ToString()

                    Call .SetText(EnmGridCol.Pending_WIPPick_Currentdate, .MaxRows, dblPendingpicklist_Today)
                    dblpendingqty = CInt(ods.Tables(0).Rows(lngloop).Item("PENDING_SCHEDULE")) - dblinvoiceQty - dblPendingpicklist_Today
                    Call .SetText(EnmGridCol.ScheduleQuantity, .MaxRows, dblpendingqty)

                    If CInt(ods.Tables(0).Rows(lngloop).Item("BALANCE_QTY")) <= 0 Then
                        .Col = -1
                        .BackColor = System.Drawing.ColorTranslator.FromOle(&HC0FFFF)
                    Else
                        .BackColor = System.Drawing.Color.White
                    End If
                Next
                .Row = 1 : .Col = EnmGridCol.ItemCode : .Row2 = .MaxRows : .Col2 = EnmGridCol.Pending_WIPPick_Currentdate : .BlockMode = True : .Lock = True : .BlockMode = False

            End With


        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    Private Sub AddColumnsInSpread()
        '****************************************************
        'PURPOSE     : TO SET THE HEADING AND COLUMN OF GRID 
        '****************************************************
        Try
            With SpItemDetails
                .MaxRows = 0
                .MaxCols = EnmGridCol.Quantity_to_be_Invoice
                .Row = 0
                .Font = Me.Font
                .ColsFrozen = EnmGridCol.DrawingNo
                .UnitType = FPSpreadADO.UnitTypeConstants.UnitTypeTwips
                .Col = EnmGridCol.Mark : .Text = "Mark" : .set_ColWidth(EnmGridCol.Mark, 400)
                .Col = EnmGridCol.ItemCode : .Text = "Item Code" : .set_ColWidth(EnmGridCol.ItemCode, 1500)
                .Col = EnmGridCol.DrawingNo : .Text = "Part No." : .set_ColWidth(EnmGridCol.DrawingNo, 1700)
                .Col = EnmGridCol.Drg_Desc : .Text = "Descrption." : .set_ColWidth(EnmGridCol.Drg_Desc, 2100)
                .Col = EnmGridCol.Salesorder : .Text = "Reference No." : .set_ColWidth(EnmGridCol.Salesorder, 1300) ' .ColHidden = True
                .Col = EnmGridCol.AmendmentNo : .Text = "Amendment No." : .set_ColWidth(EnmGridCol.AmendmentNo, 1000) ': .ColHidden = True
                .Col = EnmGridCol.BalQuantity : .Text = "SO Quantity" : .set_ColWidth(EnmGridCol.BalQuantity, 1000)
                .Col = EnmGridCol.CurrentStock : .Text = " Stock" : .set_ColWidth(EnmGridCol.CurrentStock, 750)
                .Col = EnmGridCol.ScheduleQuantity : .Text = "Pending Schedule" : .set_ColWidth(EnmGridCol.ScheduleQuantity, 1400)
                .Col = EnmGridCol.Pending_WIPPick_Currentdate : .Text = "Dispatch Advice Generated" : .set_ColWidth(EnmGridCol.Pending_WIPPick_Currentdate, 2400)
                .Col = EnmGridCol.Quantity_to_be_Invoice : .Text = "Quantity to be Invoice" : .set_ColWidth(EnmGridCol.Quantity_to_be_Invoice, 1600)
                '.Col = EnmGridCol.ItemCode : .Col2 = EnmGridCol.Pending_WIPPick_Currentdate : .BlockMode = True : .Lock = True : .BlockMode = False
                ''.Col = EnmGridCol.Quantity_to_be_Invoice : .Col2 = EnmGridCol.Quantity_to_be_Invoice : .BlockMode = True : .Lock = False : .BlockMode = False
                '.Col = EnmGridCol.Mark : .Row = .MaxRows : .Col2 = EnmGridCol.Mark : .Row2 = .MaxRows : .BlockMode = True : .ColHidden = False : .BlockMode = False
                '.Col = EnmGridCol.BalQuantity : .Row = .MaxRows : .Col2 = EnmGridCol.ScheduleQuantity : .Row2 = .MaxRows : .BlockMode = True : .ColHidden = False : .BlockMode = False
            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub CmdBUTTON_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles CmdBUTTON.ButtonClick
        Dim mblndatadeletion_req As Boolean
        Dim strQry As String
        Try
            Select Case e.Button
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                    Call EnableControls(False, Me, False)
                    SpItemDetails.Enabled = True
                    Call RefreshControls()
                    SpItemDetails.Col = 0
                    SpItemDetails.Row = 0

                    Call InitializeControls()
                    Call AddColumnsInSpread()
                    With SpItemDetails
                        .Col = EnmGridCol.Salesorder : .Row = 1 : .Col2 = EnmGridCol.Pending_WIPPick_Currentdate : .Row2 = .MaxRows : .BlockMode = True : .ColHidden = False : .BlockMode = False
                    End With
                    txtunitcode.Text = gstrUNITID
                    With txtCustomerCode
                        .Enabled = True
                        .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        .Focus()
                    End With
                    With txtSalesOrder
                        .Enabled = True
                        .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    End With

                    cmdcustcodehelp.Enabled = True
                    cmdsono.Enabled = True
                    CmdDetails.Enabled = True
                    txtinvoicetype.Text = "INV"
                    lblInvoicedescription.Text = "NORMAL INVOICE"
                    txtsubtype.Text = "F"
                    lblsubtypedescription.Text = "FINISHED GOODS"
                    optPartNo.Enabled = True
                    optPartNo.Checked = True
                    txtsearch.Text = ""
                    optItemCode.Enabled = True
                    txtsearch.Enabled = True
                    txtsearch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                    Call FRMMKTTRN0082_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                    Me.Close()

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                    optPartNo.Enabled = True
                    optPartNo.Checked = True
                    optItemCode.Enabled = True

                    If (ValidatebeforeSave() = False) Then
                        If InsertIntoPickListTable() Then
                            Call EnableControls(False, Me, True)
                            Call Me.CmdBUTTON.Revert() : gblnCancelUnload = False : gblnFormAddEdit = False
                            SpItemDetails.Enabled = True
                            Call RefreshControls()
                            With Me.txtDocCode
                                .Enabled = True
                                .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                                .Focus()
                            End With
                            cmdDocCode.Enabled = True
                        End If
                    End If

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
                    If Len(txtDocCode.Text) > 0 And SpItemDetails.MaxRows > 0 Then
                        mblndatadeletion_req = DataExist("SELECT TOP 1 1 FROM FTS_LABEL_ISSUE  WHERE  UNIT_CODE='" & gstrUNITID & "' AND DOC_TYPE=9995 AND TEMP_DISPATCH_ADVICE_NO ='" & txtDocCode.Text.Trim & "'")
                        If mblndatadeletion_req = False Then
                            If ConfirmWindow(10003, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                                If DeleteRecord() = False Then
                                    ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
                                    Exit Sub
                                End If
                                CmdBUTTON.Revert()
                            End If
                        Else
                            MsgBox("Scanning is in progress against this Disaptch Advice, Not possible to delete ", MsgBoxStyle.Information, ResolveResString(100))
                        End If
                    Else
                    End If
            End Select

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub FRMMKTTRN0082_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Try
            Select Case KeyAscii
                Case System.Windows.Forms.Keys.Escape
                    If (Me.CmdBUTTON.Mode) <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                        If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                            Call Me.CmdBUTTON.Revert() : gblnCancelUnload = False : gblnFormAddEdit = False
                            Call EnableControls(False, Me, True)
                            SpItemDetails.Enabled = True

                            With txtCustomerCode
                                .Enabled = False
                                .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                                .Focus()
                            End With

                            With Me.txtDocCode
                                .Enabled = True
                                .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                                .Focus()
                            End With

                            With Me.txtinvoicetype
                                .Enabled = False
                                .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                                .Focus()
                            End With

                            cmdDocCode.Enabled = True
                            Call RefreshControls()
                        Else
                            Me.ActiveControl.Focus()
                        End If
                    End If
            End Select

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub AddBlankRow()
        '****************************************************
        'PURPOSE     : TO ADD BALNK ROW IN GRID 
        '****************************************************
        Try
            With SpItemDetails
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .set_RowHeight(.MaxRows, 300)
                .Col = EnmGridCol.Mark
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter

                .Col = EnmGridCol.ItemCode
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter

                .Col = EnmGridCol.DrawingNo
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter

                .Col = EnmGridCol.Drg_Desc
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter

                .Col = EnmGridCol.Salesorder
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter

                .Col = EnmGridCol.AmendmentNo
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter

                .Col = EnmGridCol.BalQuantity
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter

                .Col = EnmGridCol.ScheduleQuantity
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter

                .Col = EnmGridCol.CurrentStock
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter


                .Col = EnmGridCol.Pending_WIPPick_Currentdate
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter

                .Col = EnmGridCol.Quantity_to_be_Invoice
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter

            End With

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Function ValidatebeforeSave() As Boolean
        '******************************************************************************************************************

        'PURPOSE     : TO VALIDATE BEFORE SAVING (UNIT CODE,INVOICE TYPE AND CUSTOMER CODE VALIDATION , QUANTITY SHOULD
        'BE LESS THAN STOCK,SCHEDULE AND BALANCE SALES ORDER 
        'RETURNS :TRUE OR FALSE : TRUE MEANS VALIDATION FAILED , FALSE MEANS OK
        '******************************************************************************************************************

        Try

            Dim intLoopCounter As Short
            Dim lstrControls As String
            Dim lNo As Integer
            Dim lctrFocus As System.Windows.Forms.Control
            Dim dblpendingschedule As Double
            Dim dblquantity As Double
            Dim dblstock As Double
            Dim dblSOQuantity As Double
            Dim intcounter As Integer
            Dim Stritemcode As String
            Dim StrCustDrgno As String
            Dim strQry As String
            Dim strcustref As String
            Dim SALESORDERLIST As String
            ValidatebeforeSave = False

            lNo = 1
            lstrControls = ResolveResString(10059)
            lctrFocus = Nothing
            With SpItemDetails

                If (txtunitcode.Text.Trim.Length = 0) Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Unit Code."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtunitcode
                    End If
                    ValidatebeforeSave = True
                End If

                If (txtCustomerCode.Text.Trim.Length = 0) Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Customer Code."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtCustomerCode
                    End If
                    ValidatebeforeSave = True
                End If

                For intLoopCounter = 1 To .MaxRows
                    .Row = intLoopCounter
                    .Col = EnmGridCol.Mark
                    If CBool(.Value) Then
                        .Col = EnmGridCol.Quantity_to_be_Invoice
                        .Row = intLoopCounter
                        If .Value = "" Or .Value = "0" Then
                            lstrControls = lstrControls & vbCrLf & lNo & ". Entered some value in Quantity Column."
                            ValidatebeforeSave = True
                            .Col = EnmGridCol.Quantity_to_be_Invoice : .Row = intLoopCounter : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            Exit For
                        End If

                    End If
                Next
                'ASAS'

                For intLoopCounter = 1 To .MaxRows
                    .Row = intLoopCounter
                    .Col = EnmGridCol.Mark
                    If CBool(.Value) Then
                        .Row = intLoopCounter
                        .Col = EnmGridCol.ItemCode
                        Stritemcode = .Text.Trim
                        .Col = EnmGridCol.DrawingNo
                        StrCustDrgno = .Text.Trim
                        .Col = EnmGridCol.Salesorder
                        strcustref = .Text.Trim

                        strQry = "Select dbo.UDF_CHECK_ACTIVE_SO_ITEM('" & gstrUNITID & "','" & txtCustomerCode.Text & "','" & strcustref & "','" & StrCustDrgno & "','" & Stritemcode & "')"
                        SALESORDERLIST = SqlConnectionclass.ExecuteScalar(strQry).ToString()


                        If SALESORDERLIST <> "" Then
                            If Len(SALESORDERLIST) >= 1 Then
                                lstrControls = lstrControls & vbCrLf & lNo & ".More than one sales order Active for item code." & Stritemcode
                                lNo = lNo + 1
                                ValidatebeforeSave = True
                                .Col = EnmGridCol.Quantity_to_be_Invoice : .Row = intLoopCounter : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                Exit For
                            End If
                        End If
                    End If
                Next
                'ASAS'

                For intLoopCounter = 1 To .MaxRows
                    .Row = intLoopCounter
                    .Col = EnmGridCol.Mark
                    If CBool(.Value) Then
                        .Col = EnmGridCol.Quantity_to_be_Invoice
                        dblquantity = Val(.Text)


                        .Col = EnmGridCol.CurrentStock
                        dblstock = Val(.Text)

                        .Col = EnmGridCol.BalQuantity
                        dblSOQuantity = Val(.Text)

                        .Col = EnmGridCol.ScheduleQuantity
                        dblpendingschedule = Val(.Text)

                        If dblquantity > dblstock Then
                            lstrControls = lstrControls & vbCrLf & lNo & ". Entered Value greater than Stock."
                            lNo = lNo + 1
                            ValidatebeforeSave = True
                            .Col = EnmGridCol.Quantity_to_be_Invoice : .Row = intLoopCounter : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            Exit For
                        End If

                        If dblSOQuantity > 0 And dblquantity > dblSOQuantity Then
                            lstrControls = lstrControls & vbCrLf & lNo & ". Entered Value greater than Sales order Quantity."
                            lNo = lNo + 1
                            ValidatebeforeSave = True
                            .Col = EnmGridCol.Quantity_to_be_Invoice : .Row = intLoopCounter : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            Exit For
                        End If

                        If dblquantity > dblpendingschedule Then
                            lstrControls = lstrControls & vbCrLf & lNo & ". Entered Value greater than Pending Schedule ."
                            lNo = lNo + 1
                            ValidatebeforeSave = True
                            .Col = EnmGridCol.Quantity_to_be_Invoice : .Row = intLoopCounter : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            Exit For
                        End If

                    End If
                Next

                intcounter = 0
                For intLoopCounter = 1 To .MaxRows
                    .Row = intLoopCounter
                    .Col = EnmGridCol.Mark
                    If CBool(.Value) Then
                        intcounter = 1
                        Exit For
                    End If
                Next
                If intcounter = 0 Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Select at least one row. "
                    lNo = lNo + 1
                    ValidatebeforeSave = True

                End If
            End With
            If (ValidatebeforeSave = True) Then
                MsgBox(lstrControls, MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Function

    Private Sub SpItemDetails_LeaveCell(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles SpItemDetails.LeaveCell
        Dim lstrReturnVal As String
        Dim IntquantitytobeInvoice As Integer
        Dim intstock As Integer
        Dim intSchedule As Integer
        If e.newRow = -1 Then Exit Sub
        Try


            lstrReturnVal = ""
            If CmdBUTTON.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                With SpItemDetails

                    .Row = e.row
                    .Col = EnmGridCol.Mark
                    If CBool(.Value) Then
                        .Col = EnmGridCol.Quantity_to_be_Invoice
                        .Row = e.row
                        IntquantitytobeInvoice = Val(.Value)

                        .Col = EnmGridCol.CurrentStock
                        .Row = e.row
                        intstock = Val(.Value)

                        .Col = EnmGridCol.ScheduleQuantity
                        .Row = e.row
                        intSchedule = Val(.Value)

                        If (IntquantitytobeInvoice > intstock) Then
                            .Row = e.row
                            .Col = e.col
                            .Text = ""
                            MsgBox("Entered Quantity can't exceed than stock ", MsgBoxStyle.Critical, ResolveResString(100))
                            Exit Sub
                        End If
                        If (IntquantitytobeInvoice > intSchedule) Then
                            .Row = e.row
                            .Col = e.col
                            .Text = ""
                            MsgBox("Entered Quantity can't exceed than Pending Schedule ", MsgBoxStyle.Critical, ResolveResString(100))
                            Exit Sub
                        End If
                    End If
                End With
            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub FillDataInSpread()
        '**********************************************************************************************************************************
        'PURPOSE     : TO FILL THE DETAIL DATA IN VIEW MODE , USER CAN VIEW ALL THE EXISTING DETAILS OF PICK LIST ONLY FOR TODAYS DATE
        '**********************************************************************************************************************************
        Dim strQry As String = String.Empty
        Dim dtCust As DataTable
        Dim lvwItem As ListViewItem
        Dim StrPartCode As String
        Dim StrPartDescription As String
        Dim StrItemCode As String
        Dim StrQuantity As String

        If txtDocCode.Text.Trim = "" Then Exit Sub
        Try
            Select Case CmdBUTTON.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    strQry = "Select * From FTS_FG_PICKLIST_INV_HDR HD INNER JOIN FTS_FG_PICKLIST_INV_DTL DT ON HD.UNIT_CODE=DT.UNIT_CODE "
                    strQry += " AND HD.DOC_TYPE=DT.DOC_TYPE AND HD.DOC_NO=DT.DOC_NO AND HD.UNIT_CODE='" & gstrUNITID & "' AND HD.Doc_Type = 9995"
                    strQry += "  AND HD.Doc_No = '" & txtDocCode.Text.Trim & "'"

                    dtCust = SqlConnectionclass.GetDataTable(strQry)
                    If dtCust.Rows.Count > 0 Then

                        txtunitcode.Text = gstrUNITID
                        txtinvoicetype.Text = "INV"
                        lblInvoicedescription.Text = "NORMAL INVOICE"
                        txtsubtype.Text = "F"
                        lblsubtypedescription.Text = "FINISHED GOODS"


                        With SpItemDetails

                            For Each Row As DataRow In dtCust.Rows

                                Call AddBlankRow()
                                .Col = EnmGridCol.DrawingNo
                                .Text = Convert.ToString(Row("CUST_ITEM_CODE"))
                                StrPartCode = .Text.Trim

                                .Col = EnmGridCol.Drg_Desc
                                .Text = Convert.ToString(Row("CUST_ITEM_DESC"))
                                StrPartDescription = .Text.Trim


                                .Col = EnmGridCol.ItemCode
                                .Text = Convert.ToString(Row("ITEM_CODE"))
                                StrItemCode = .Text.Trim

                                .Col = EnmGridCol.Quantity_to_be_Invoice
                                .Text = Convert.ToString(Row("quantity"))
                                StrQuantity = .Text.Trim
                            Next
                            If CmdBUTTON.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                                '.Col = EnmGridCol.Mark : .Row = .MaxRows : .Col2 = EnmGridCol.Quantity_to_be_Invoice : .Row2 = .MaxRows : .BlockMode = True : .Lock = True : .BlockMode = False
                                '.Col = EnmGridCol.Mark : .Row = .MaxRows : .Col2 = EnmGridCol.Mark : .Row2 = .MaxRows : .BlockMode = True : .ColHidden = True : .BlockMode = False
                                .Col = EnmGridCol.Salesorder : .Row = 1 : .Col2 = EnmGridCol.Pending_WIPPick_Currentdate : .Row2 = .MaxRows : .BlockMode = True : .ColHidden = True : .BlockMode = False
                                .Enabled = False
                                End If 
                        End With


                    End If

            End Select

        Catch ex As Exception
            Throw ex
        Finally
            If IsNothing(dtCust) = False Then dtCust.Dispose()
            dtCust = Nothing

        End Try
    End Sub
    Private Function DeleteRecord() As Boolean
        '*****************************************************************************************************************
        'PURPOSE     : TO DELETE THE EXISTING RECORD , ( BUT ONLY THOSE PICK LIST WHOSE WHOLE PICK LIST IS UNSCANNED (FULL PENDING FOR SCAN  )
        '*****************************************************************************************************************

        Dim strQry As String
        Dim sqlConn As New SqlConnection
        Dim sqlCmd As SqlCommand
        Dim sqlTrans As SqlTransaction = Nothing
        Dim blnTran As Boolean

        DeleteRecord = False
        Try
            sqlConn = SqlConnectionclass.GetConnection
            sqlTrans = sqlConn.BeginTransaction
            blnTran = True

            sqlCmd = New SqlCommand()
            sqlCmd.Connection = sqlConn
            sqlCmd.Transaction = sqlTrans

            strQry = "DELETE FROM FTS_FG_PICKLIST_INV_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND DOC_NO='" & txtDocCode.Text.Trim & "'"
            sqlCmd.CommandType = CommandType.Text
            sqlCmd.CommandText = strQry
            sqlCmd.ExecuteNonQuery()

            strQry = "DELETE FROM FTS_FG_PICKLIST_INV_HDR WHERE UNIT_CODE='" + gstrUNITID + "' AND DOC_NO='" & txtDocCode.Text.Trim & "'"
            sqlCmd.CommandType = CommandType.Text
            sqlCmd.CommandText = strQry
            sqlCmd.ExecuteNonQuery()

            sqlTrans.Commit()
            blnTran = False
            sqlConn.Close()
            DeleteRecord = True
            MsgBox("Record Deleted successfully .", MsgBoxStyle.Information, ResolveResString(100))
            Call Me.CmdBUTTON.Revert() : gblnCancelUnload = False : gblnFormAddEdit = False
            Call EnableControls(False, Me, True)
            Call RefreshControls()
            SpItemDetails.Enabled = True
            With Me.txtDocCode
                .Enabled = True
                .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                .Focus()
            End With
            cmdDocCode.Enabled = True

        Catch ex As Exception
            If blnTran Then sqlTrans.Rollback()
            If sqlConn.State = ConnectionState.Open Then sqlConn.Close()
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Function
    Private Sub txtDocCode_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDocCode.Validating
        Dim strQry As String
        Dim dtdocno As New DataTable
        If txtDocCode.Text.Trim = "" Then Exit Sub

        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            If CmdBUTTON.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                strQry = "SELECT DOC_NO,CUSTOMER_CODE FROM FTS_FG_PICKLIST_INV_HDR WHERE UNIT_CODE='" & gstrUNITID & "' "
                strQry += "AND DOC_TYPE = 9995 AND DOC_NO = '" & txtDocCode.Text.Trim & "'"
            End If
            Call GetData(strQry, dtdocno)
            If dtdocno.Rows.Count > 0 Then
                txtDocCode.Text = dtdocno.Rows(0).Item("DOC_NO")
                txtCustomerCode.Text = dtdocno.Rows(0).Item("customer_code")
                Call AddColumnsInSpread()
                Call FillDataInSpread()
                optPartNo.Enabled = True
                optPartNo.Checked = True
                txtsearch.Text = ""
                optItemCode.Enabled = True
                txtsearch.Enabled = True
                txtsearch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            Else
                Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
                e.Cancel = True
                Exit Sub
            End If
        Catch EX As Exception
            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub
    Private Sub txtsearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtsearch.TextChanged
        Try
            Call SearchItem()

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub
    Private Sub SearchItem()
        '*******************************************************************************
        'PURPOSE     : TO SEARCH THE ITEM CODE AND PART CODE  : GRID ITEMS 
        '******************************************************************************

        Dim intCount As Short
        Try
            With SpItemDetails
                .Row = -1
                .Col = -1
                .Font = VB6.FontChangeBold(.Font, False)
                If optPartNo.Checked Then
                    .Col = EnmGridCol.DrawingNo
                End If

                If optItemCode.Checked Then
                    .Col = EnmGridCol.ItemCode
                End If

                For intCount = 1 To .MaxRows
                    .Row = intCount
                    If UCase(Mid(.Text, 1, Len(txtsearch.Text))) = UCase(txtsearch.Text) Then
                        .TopRow = .Row
                        .Col = -1
                        .Font = VB6.FontChangeBold(.Font, True)
                        Exit Sub
                    End If
                Next
            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub cmdSo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim StrSOHelp As String
        Dim STRSQL As String
        Try

            StrSOHelp = ShowList(0, Len(txtSalesOrder.Text), txtSalesOrder.Text, "Cust_ref", "" & DateColumnNameInShowList("Amendment_No") & " as Amendment_No", "VW_FTS_ALL_ACTIVESALESORDER", "AND ACCOUNT_CODE='" & Me.txtCustomerCode.Text.Trim & "'  ", "SALES ORDER LIST")
            If StrSOHelp = "-1" Then

                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                With Me.txtSalesOrder
                    .Enabled = True
                    .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    .Focus()
                End With
                cmdDocCode.Enabled = True

            ElseIf StrSOHelp = "" Then
                Me.txtSalesOrder.Focus()
            Else
                Me.txtSalesOrder.Text = StrSOHelp
                txtSalesOrder_Validating(txtSalesOrder, New System.ComponentModel.CancelEventArgs(False))
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtSalesOrder_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSalesOrder.TextChanged
        Try
            With SpItemDetails
                .MaxRows = 0
                Me.txtamendment.Text = String.Empty
            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub txtSalesOrder_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtSalesOrder.Validating
        Dim strQry As String
        Dim dtsalesorder As New DataTable
        If txtSalesOrder.Text.Trim = "" Then Exit Sub

        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            If CmdBUTTON.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                strQry = "SELECT CUST_REF FROM VW_FTS_ALL_ACTIVESALESORDER  "
                strQry += " WHERE ACCOUNT_CODE= '" & Me.txtCustomerCode.Text.Trim & "' and unit_code= '" & gstrUNITID & "' AND CUST_REF='" & Me.txtSalesOrder.Text.Trim & "'"
            End If
            Call GetData(strQry, dtsalesorder)
            If dtsalesorder.Rows.Count = 0 Then
                MsgBox("No Sales Order exists for FTS ITEMS ", MsgBoxStyle.Information, ResolveResString(100))
                Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
                txtSalesOrder.Text = String.Empty
                txtamendment.Text = String.Empty
                e.Cancel = True
                Exit Sub
            End If

            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)

        Catch EX As Exception
            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Private Sub cmdsono_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdsono.Click
        Dim StrSOHelp As String
        Dim STRSALESORDER() As String
        Try

            StrSOHelp = "SELECT CUST_REF, AMENDMENT_NO FROM VW_FTS_ALL_ACTIVESALESORDER WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & Me.txtCustomerCode.Text.Trim & "'  "

            STRSALESORDER = Me.ctlhelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, StrSOHelp, "Sales Order No")

            If UBound(STRSALESORDER) <= 0 Then Exit Sub

            If STRSALESORDER(0) = "0" Then
                MsgBox("No Sales Order Available To Display", MsgBoxStyle.Information, "eMPro") : txtSalesOrder.Text = "" : txtSalesOrder.Focus() : Exit Sub
            Else
                txtSalesOrder.Text = STRSALESORDER(0)
                txtamendment.Text = STRSALESORDER(1)
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub
End Class