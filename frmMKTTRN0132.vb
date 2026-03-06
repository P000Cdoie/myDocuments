'******************************************************************************************************************************************************
'COPYRIGHT       : MIND LTD.
'MODULE          : FRMMKTTRN0082 - WIP FG PICK LIST GENERATION 
'AUTHOR          : PRASHANT RAJPAL
'CREATION DATE   : 11-FEB 2014
'ISSUE ID        : 10533478 
'PURPOSE         : TO GENERATE THE PICK LIST AND PROVIDE DATA TO SCANNER TO SCAN (ONLY FOR THOSE ITEMS AND CUSTOMERS WHOSE WIP FLAG FUNCITONALITY IS ON
'*******************************************************************************************************************************************************
' REVISION DATE     : 01 JULY 2014
' REVISED BY        : PRASHANT RAJPAL
' ISSUE ID          : 10623079
' REVISION HISTORY  : WIP CHANGES FOR VACUFORM CHANGES 
'*******************************************************************************************************************************************************

Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class frmMKTTRN0132
    Inherits System.Windows.Forms.Form
    Private mlngFormTag As Integer
    Private mServerDate As String
    Dim mstrUpdateDocumentNoSQL As String
    Private Enum EnmGridCol
        Invoiceno = 1
        ItemCode = 2
        DrawingNo = 3
        Drg_Desc = 4
        Quantity = 5
        kanbanno = 6
        kanabnnoHelp = 7
        kanbannoqty = 8
        Sch_Date = 9
        sch_time = 10
        Unloc = 11
        UsLoc = 12
        BatchCode = 13
    End Enum
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
        Dim blnsalesorderpicklist As Boolean
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
            strQry = "Select ALLOW_WIPFG_SALESORDER_PICKLIST from sales_parameter where unit_code='" & gstrUNITID & "'"
            blnsalesorderpicklist = SqlConnectionclass.ExecuteScalar(strQry).ToString()
            '10623079
            If blnsalesorderpicklist = True Then
                txtInvoiceNo.Visible = False
                txtInvoiceDate.Visible = False
                cmdInvoiceNo.Visible = False
                lblkanbanno.Visible = False
            Else
                txtInvoiceNo.Visible = True
                txtInvoiceDate.Visible = True
                cmdInvoiceNo.Visible = True
                lblkanbanno.Visible = True
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


            With Me.txtinvoicetype
                .Enabled = False
                .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                .Focus()
            End With

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

    Private Sub SpItemDetails_ClickEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles SpItemDetails.ClickEvent
        Try
            Dim strvaritemcode As Object
            Dim strvarpartcode As Object
            Dim strSpdetail() As String
            If SpItemDetails.ActiveRow >= 1 And SpItemDetails.ActiveCol = EnmGridCol.kanabnnoHelp Then
                If SpItemDetails.ActiveCol = EnmGridCol.kanabnnoHelp Then
                    strvaritemcode = Nothing
                    Call SpItemDetails.GetText(EnmGridCol.ItemCode, e.row, strvaritemcode)
                    strvarpartcode = Nothing
                    Call SpItemDetails.GetText(EnmGridCol.DrawingNo, e.row, strvarpartcode)

                    strSpdetail = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT KANBANNO,SCH_DATE,SCH_TIME,QUANTITY,UNLOC,USLOC FROM UFN_PENDING_KANBAN_FOR_INVOICE_MAPPING('" & gstrUNITID & "','" & txtInvoiceNo.Text & "','" & txtCustomerCode.Text & "','" & strvaritemcode & "','" & strvarpartcode & "') ORDER BY SCH_DATE, SCH_TIME")
                    If Not (UBound(strSpdetail) = -1) Then
                        If (Len(strSpdetail(0)) >= 1) And strSpdetail(0) = "0" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            Call SpItemDetails.SetText(EnmGridCol.kanbanno, e.row, strSpdetail(0))
                            Call SpItemDetails.SetText(EnmGridCol.Sch_Date, e.row, strSpdetail(1))
                            Call SpItemDetails.SetText(EnmGridCol.sch_time, e.row, strSpdetail(2))
                            Call SpItemDetails.SetText(EnmGridCol.kanbannoqty, e.row, strSpdetail(3))
                            Call SpItemDetails.SetText(EnmGridCol.Unloc, e.row, strSpdetail(4))
                            Call SpItemDetails.SetText(EnmGridCol.UsLoc, e.row, strSpdetail(5))
                        End If
                    End If

                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
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
    Private Sub txtunitcode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtunitcode.TextChanged

        Try
            Select Case CmdBUTTON.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    If (txtunitcode.Text.Trim.Length = 0) Then
                        Call RefreshControls()
                        lblunitdesc.Text = String.Empty

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
            'txtDocCode.Text = String.Empty
            txtInvoiceNo.Text = String.Empty


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
    Private Function Savedata() As Boolean
        '**********************************************************************************************************************************
        'FUNCTION :MAILY USED TO INSERT THE DATA IN WIP_FG_PICKLIST HDR AND DTL, AND ALS UPDATE SERIES FROM DOCUMENTTYPE_MST WITH DOC TYPE 9990
        '**********************************************************************************************************************************
        Savedata = False
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
        Dim STRINVOICENO As String
        Dim STRKANBANNO As String

        strinvoicetype = "INV"
        Try
            sqlConn = SqlConnectionclass.GetConnection
            sqlTrans = sqlConn.BeginTransaction
            blnTran = True

            sqlCmd = New SqlCommand()
            sqlCmd.Connection = sqlConn
            sqlCmd.Transaction = sqlTrans

            With SpItemDetails
                For intLoopCounter = 1 To .MaxRows
                

                    .Row = intLoopCounter
                    .Col = EnmGridCol.Invoiceno
                    STRINVOICENO = .Text.Trim
                    .Row = intLoopCounter
                    .Col = EnmGridCol.DrawingNo
                    StrPartCode = .Text.Trim
                    .Col = EnmGridCol.Drg_Desc
                    StrDrgdesc = .Text.Trim
                    .Col = EnmGridCol.ItemCode
                    stritemcode = .Text.Trim
                    .Row = intLoopCounter
                    .Col = EnmGridCol.kanbanno
                    STRKANBANNO = .Text.Trim

                    strQry = "UPDATE SALES_DTL SET SRVDINO='" & STRKANBANNO & "',Upd_Userid = '" & mP_User & "' ,Upd_Dt = '" & getDateForDB(GetServerDate()) & "'WHERE UNIT_CODE='" & gstrUNITID & "'AND DOC_NO='" & STRINVOICENO & "' AND ITEM_CODE='" & stritemcode & "' AND CUST_ITEM_CODE='" & StrPartCode & "'"
                    sqlCmd.CommandType = CommandType.Text
                    sqlCmd.CommandText = strQry
                    sqlCmd.ExecuteNonQuery()


                Next
            End With

        
            sqlTrans.Commit()
            blnTran = False
            sqlConn.Close()
            Savedata = True
            'txtDocCode.Text = Documentno
            MsgBox("Record Updated successsfully ", MsgBoxStyle.Information, ResolveResString(100))

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

                    strCustHelp = ShowList(0, Len(txtunitcode.Text), , "CUSTOMER_CODE", "CUST_NAME", "CUSTOMER_MST", " AND UNIT_CODE ='" & gstrUNITID & "' AND MARUTI_KANBAN_WAREHOUSE_ENABLED =1 ", "CUSTOMER HELP", , , , , )
                    If strCustHelp = "-1" Then
                        'txtCustomerCode.Focus()
                        Exit Sub

                    ElseIf strCustHelp = "" Then
                        'txtCustomerCode.Focus()
                    Else
                        txtCustomerCode.Text = strCustHelp
                        lblCustomerdesc.Text = GetQryOutput("SELECT CUST_NAME FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & strCustHelp & "' AND MARUTI_KANBAN_WAREHOUSE_ENABLED=1")
                        txtInvoiceNo.Text = ""
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
                If Me.txtCustomerCode.Text.Length = 0 Then Me.lblCustomerdesc.Text = String.Empty : txtInvoiceNo.Text = String.Empty : lblInvoicedescription.Text = String.Empty
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
                strQry += " WHERE CUSTOMER_CODE= '" & Me.txtCustomerCode.Text.Trim & "' and unit_code= '" & gstrUNITID & "' AND MARUTI_KANBAN_WAREHOUSE_ENABLED =1 "
            End If
            Call GetData(strQry, dtCustomer)
            If dtCustomer.Rows.Count = 0 Then
                MsgBox("Invalid Customer Code ", MsgBoxStyle.Information, ResolveResString(100))
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
            Dim FIFIWISEINVOICELIST As String
            Dim strQry As String
            If Me.txtCustomerCode.Text.Length <= 0 Then
                MsgBox("Please Select First Customer ", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                txtCustomerCode.Text = String.Empty

                Exit Sub
            End If
            If Me.txtInvoiceNo.Text.Length <= 0 Then
                MsgBox("Please Select Invoice No.", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                txtInvoiceNo.Text = String.Empty
                '                txtCustomerCode.Focus()
                Exit Sub
            End If
            'FIFIWISEINVOICELIST = Find_Value("Select dbo.UDF_CHECK_PENDING_NAGAREINVOICE_FOR_MAPPING('" & gstrUNITID & "','" & txtCustomerCode.Text & "','" & txtInvoiceNo.Text & "')")

            strQry = "Select dbo.UDF_CHECK_PENDING_NAGAREINVOICE_FOR_MAPPING('" & gstrUNITID & "','" & txtCustomerCode.Text & "','" & txtInvoiceNo.Text & "')"
            FIFIWISEINVOICELIST = SqlConnectionclass.ExecuteScalar(strQry)

            If FIFIWISEINVOICELIST <> "" Then
                MsgBox("KINDLY MAP THE INOICE AS PER FIFO , PENDING INVOICE(S) : " & FIFIWISEINVOICELIST, MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If

            Call FILLKANBANDETAILS()
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Public Function FILLKANBANDETAILS() As Boolean
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
        Dim StrInvoiceNo As String
        Dim StrsalesQty As String
        Dim lngloop As Long
        Try
            If Me.txtCustomerCode.Text.Length <= 0 Then
                MsgBox("Please Select First Customer ", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                txtCustomerCode.Text = String.Empty
                '                txtCustomerCode.Focus()
                Exit Function
            End If

            FILLKANBANDETAILS = True
            sqlConn = SqlConnectionclass.GetConnection
            'sqlTrans = sqlConn.BeginTransaction
            blnTran = True

            sqlCmd = New SqlCommand()
            sqlCmd.Connection = sqlConn
            'sqlCmd.Transaction = sqlTrans

            With sqlCmd
                .Parameters.Clear()
                .CommandType = CommandType.StoredProcedure
                .CommandText = "USP_PENDINGKANBAN_MAPPING_DATA"
                .Parameters.AddWithValue("@Unit_Code", gstrUNITID)
                .Parameters.AddWithValue("@CUST_CODE", txtCustomerCode.Text.Trim)
                .Parameters.AddWithValue("@INVTYPE", txtinvoicetype.Text.Trim)
                .Parameters.AddWithValue("@INVSUBTYPE", "F")
                .Parameters.AddWithValue("@INVOICENO", txtInvoiceNo.Text)
                .Parameters.Add("@MSG", SqlDbType.VarChar, 5000).Direction = ParameterDirection.Output
                .ExecuteNonQuery()
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
                    StrCustDrgno = CStr(ods.Tables(0).Rows(lngloop).Item("CUST_ITEM_CODE"))
                    StrCustDrgdesc = CStr(ods.Tables(0).Rows(lngloop).Item("CUST_ITEM_DESC"))
                    StrInvoiceNo = CStr(ods.Tables(0).Rows(lngloop).Item("INVOICE_NO"))
                    StrsalesQty = CStr(ods.Tables(0).Rows(lngloop).Item("SALES_QUANTITY"))
                    'Str
                    Call .SetText(EnmGridCol.Invoiceno, .MaxRows, StrInvoiceNo)
                    Call .SetText(EnmGridCol.ItemCode, .MaxRows, Stritemcode)
                    Call .SetText(EnmGridCol.DrawingNo, .MaxRows, StrCustDrgno)
                    Call .SetText(EnmGridCol.Drg_Desc, .MaxRows, StrCustDrgdesc)
                    Call .SetText(EnmGridCol.Quantity, .MaxRows, StrsalesQty)



                Next
                .Row = 1 : .Col = EnmGridCol.Invoiceno : .Row2 = .MaxRows : .Col2 = EnmGridCol.BatchCode : .BlockMode = True : .Lock = True : .BlockMode = False

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
                .MaxCols = EnmGridCol.UsLoc
                .Row = 0
                .Font = Me.Font
                .ColsFrozen = EnmGridCol.DrawingNo
                .UnitType = FPSpreadADO.UnitTypeConstants.UnitTypeTwips
                .Col = EnmGridCol.Invoiceno : .Text = "Invoiceno" : .set_ColWidth(EnmGridCol.Invoiceno, 800)
                .Col = EnmGridCol.ItemCode : .Text = "Item Code" : .set_ColWidth(EnmGridCol.ItemCode, 1000)
                .Col = EnmGridCol.DrawingNo : .Text = "Part No." : .set_ColWidth(EnmGridCol.DrawingNo, 1200)
                .Col = EnmGridCol.Drg_Desc : .Text = "Descrption." : .set_ColWidth(EnmGridCol.Drg_Desc, 2200)
                .Col = EnmGridCol.Quantity : .Text = "Qty" : .set_ColWidth(EnmGridCol.Quantity, 800)
                .Col = EnmGridCol.kanbanno : .Text = "KANBANNO" : .set_ColWidth(EnmGridCol.kanbanno, 1500)
                .Col = EnmGridCol.kanabnnoHelp : .Text = "HELP" : .set_ColWidth(EnmGridCol.kanabnnoHelp, 600)
                .Col = EnmGridCol.kanbannoqty : .Text = "PENDING KABNANQTY" : .set_ColWidth(EnmGridCol.kanbannoqty, 700)
                .Col = EnmGridCol.Sch_Date : .Text = "SCH DATE " : .set_ColWidth(EnmGridCol.Sch_Date, 850)
                .Col = EnmGridCol.sch_time : .Text = "SCH TIME " : .set_ColWidth(EnmGridCol.sch_time, 800)
                .Col = EnmGridCol.Unloc : .Text = "UNLOC " : .set_ColWidth(EnmGridCol.Unloc, 700)
                .Col = EnmGridCol.UsLoc : .Text = " USLOC" : .set_ColWidth(EnmGridCol.UsLoc, 700)
                .Col = EnmGridCol.BatchCode : .Text = "BATCH CODE" : .set_ColWidth(EnmGridCol.BatchCode, 700)

                
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
                    Call AddColumnsInSpread()

                    txtunitcode.Text = gstrUNITID
                    With txtCustomerCode
                        .Enabled = False
                        .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        .Focus()
                    End With
                    With txtInvoiceNo
                        .Enabled = False
                        .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    End With

                    cmdcustcodehelp.Enabled = True
                    cmdInvoiceNo.Enabled = True
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
                        If Savedata() Then
                            Call EnableControls(False, Me, True)
                            Call Me.CmdBUTTON.Revert() : gblnCancelUnload = False : gblnFormAddEdit = False
                            SpItemDetails.Enabled = True
                            Call RefreshControls()

                        End If
                    End If

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE

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


                            With Me.txtinvoicetype
                                .Enabled = False
                                .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                                .Focus()
                            End With

                            'cmdDocCode.Enabled = True
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
                .Col = EnmGridCol.Invoiceno
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
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

                
                .Col = EnmGridCol.Quantity
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter

                .Col = EnmGridCol.kanbanno
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter


                .Col = EnmGridCol.kanabnnoHelp
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonPicture = My.Resources.ico111.ToBitmap

                .Col = EnmGridCol.kanbannoqty
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter

                .Col = EnmGridCol.Sch_Date
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter

                .Col = EnmGridCol.sch_time
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
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
            Dim dblKanbanQty As Double
            Dim strKanbanNo As String
            Dim intcounter As Integer
            Dim strQry As String
            Dim FIFIWISEINVOICELIST As String = String.Empty
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

                If (txtInvoiceNo.Text.Trim.Length = 0) Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Invoice No."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtInvoiceNo
                    End If
                    ValidatebeforeSave = True
                End If


                'FIFIWISEINVOICELIST = Find_Value("Select dbo.UDF_CHECK_PENDING_NAGAREINVOICE_FOR_MAPPING('" & gstrUNITID & "','" & txtCustomerCode.Text & "','" & txtInvoiceNo.Text & "')")
                strQry = "Select dbo.UDF_CHECK_PENDING_NAGAREINVOICE_FOR_MAPPING('" & gstrUNITID & "','" & txtCustomerCode.Text & "','" & txtInvoiceNo.Text & "')"

                FIFIWISEINVOICELIST = SqlConnectionclass.ExecuteScalar(strQry)

                If FIFIWISEINVOICELIST.ToString <> "" Then
                    MsgBox("KINDLY MAP THE INOICE AS PER FIFO , PENDING INVOICE(S) : " & FIFIWISEINVOICELIST, MsgBoxStyle.Information, ResolveResString(100))
                    ValidatebeforeSave = True
                End If


                For intLoopCounter = 1 To .MaxRows
                    .Row = intLoopCounter
                    .Col = EnmGridCol.Quantity
                    dblquantity = Val(.Text)


                    .Col = EnmGridCol.kanbannoqty
                    dblKanbanQty = Val(.Text)

                    
                    If dblquantity > dblKanbanQty Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". KanbanQty can't be less than Invoice Qty."
                        lNo = lNo + 1
                        ValidatebeforeSave = True
                        .Col = EnmGridCol.kanbannoqty : .Row = intLoopCounter : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        Exit For
                    End If

                    .Row = intLoopCounter
                    .Col = EnmGridCol.kanbanno
                    strKanbanNo = .Text
                    If strKanbanNo = "" Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". Kanban No can't be Empty."
                        lNo = lNo + 1
                        ValidatebeforeSave = True
                        .Col = EnmGridCol.kanabnnoHelp : .Row = intLoopCounter : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        Exit For
                    End If

                Next

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

        'If txtDocCode.Text.Trim = "" Then Exit Sub
        Try
            Select Case CmdBUTTON.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    strQry = "Select * From WIP_FG_PICKLIST_INV_HDR HD INNER JOIN WIP_FG_PICKLIST_INV_DTL DT ON HD.UNIT_CODE=DT.UNIT_CODE "
                    strQry += " AND HD.DOC_TYPE=DT.DOC_TYPE AND HD.DOC_NO=DT.DOC_NO AND HD.UNIT_CODE='" & gstrUNITID & "' AND HD.Doc_Type = 9990"
                    'strQry += "  AND HD.Doc_No = '" & txtDocCode.Text.Trim & "'"

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

                                
                            Next

                            .Enabled = False
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
    End Function
    Private Sub txtDocCode_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Dim strQry As String
        Dim dtdocno As New DataTable
        '    If txtDocCode.Text.Trim = "" Then Exit Sub

        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            If CmdBUTTON.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                strQry = "SELECT DOC_NO,CUSTOMER_CODE FROM WIP_FG_PICKLIST_INV_HDR WHERE UNIT_CODE='" & gstrUNITID & "' "
                'strQry += "AND DOC_TYPE = 9990 AND DOC_NO = '" & txtDocCode.Text.Trim & "'"
                strQry += "AND DOC_TYPE = 9990 "
            End If
            Call GetData(strQry, dtdocno)
            If dtdocno.Rows.Count > 0 Then
                'txtDocCode.Text = dtdocno.Rows(0).Item("DOC_NO")
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

            StrSOHelp = ShowList(0, Len(txtInvoiceNo.Text), txtInvoiceNo.Text, "Cust_ref", "" & DateColumnNameInShowList("Amendment_No") & " as Amendment_No", "VW_ALL_ACTIVESALESORDER", "AND ACCOUNT_CODE='" & Me.txtCustomerCode.Text.Trim & "'  ", "SALES ORDER LIST")
            If StrSOHelp = "-1" Then

                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                With Me.txtInvoiceNo
                    .Enabled = True
                    .BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    .Focus()
                End With
                'cmdDocCode.Enabled = True

            ElseIf StrSOHelp = "" Then
                Me.txtInvoiceNo.Focus()
            Else
                Me.txtInvoiceNo.Text = StrSOHelp
                txtSalesOrder_Validating(txtInvoiceNo, New System.ComponentModel.CancelEventArgs(False))
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub txtSalesOrder_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoiceNo.TextChanged
        Try
            With SpItemDetails
                .MaxRows = 0
                Me.txtInvoiceDate.Text = String.Empty
            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

    Private Sub txtSalesOrder_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtInvoiceNo.Validating
        Dim strQry As String
        Dim FIFIWISEINVOICELIST As String = String.Empty
        Dim dtsalesorder As New DataTable
        If txtInvoiceNo.Text.Trim = "" Then Exit Sub

        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            If CmdBUTTON.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                strQry = "SELECT * FROM saleschallan_dtl where unit_code='" & gstrUNITID & "' and doc_no='" & txtInvoiceNo.Text & "' and account_code='" & txtCustomerCode.Text & "'"
            End If
            Call GetData(strQry, dtsalesorder)
            If dtsalesorder.Rows.Count = 0 Then
                MsgBox("This InvoiceNo does not exist", MsgBoxStyle.Information, ResolveResString(100))
                Call ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
                txtInvoiceNo.Text = String.Empty
                txtInvoiceDate.Text = String.Empty
                e.Cancel = True
                Exit Sub
            End If

            'FIFIWISEINVOICELIST = Find_Value("Select dbo.UDF_CHECK_PENDING_NAGAREINVOICE_FOR_MAPPING('" & gstrUNITID & "','" & txtCustomerCode.Text & "','" & txtInvoiceNo.Text & "')")
            strQry = "Select dbo.UDF_CHECK_PENDING_NAGAREINVOICE_FOR_MAPPING('" & gstrUNITID & "','" & txtCustomerCode.Text & "','" & txtInvoiceNo.Text & "')"

            FIFIWISEINVOICELIST = SqlConnectionclass.ExecuteScalar(strQry)

            If FIFIWISEINVOICELIST <> "" Then
                MsgBox("KINDLY MAP THE INOICE AS PER FIFO , PENDING INVOICE(S) : " & FIFIWISEINVOICELIST, MsgBoxStyle.Information, ResolveResString(100))
                e.Cancel = True
                Exit Sub
            End If


            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)

        Catch EX As Exception
            MsgBox(EX.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Private Sub cmdKanbanNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdInvoiceNo.Click
        Dim StrHelp As String
        Dim STRKANBANNO() As String
        Try
            If txtCustomerCode.Text.Length <= 0 Then
                MsgBox("Please Select First Customer Code ", MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                txtInvoiceNo.Text = String.Empty
                'txtCustomerCode.Focus()
                Exit Sub
            End If

            StrHelp = "SELECT INVOICE_NO ,INVOICE_DATE,ITEM_CODE,CUST_ITEM_CODE,SALES_QUANTITY AS SALES_QTY FROM FN_INVOICE_KANBAN_PENDING('" & gstrUNITID & "','" & txtCustomerCode.Text & "' )"

            STRKANBANNO = Me.ctlhelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, StrHelp, "Invoice No")
            ctlhelp.AutoSizeMode = Windows.Forms.AutoSizeMode.GrowAndShrink




            If UBound(STRKANBANNO) <= 0 Then Exit Sub

            If STRKANBANNO(0) = "0" Then
                MsgBox("No Invoice Pending to Map", MsgBoxStyle.Information, "eMPro") : txtInvoiceNo.Text = "" : txtInvoiceNo.Focus() : Exit Sub
            Else
                txtInvoiceNo.Text = STRKANBANNO(0)
                txtInvoiceDate.Text = STRKANBANNO(1)
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub
End Class