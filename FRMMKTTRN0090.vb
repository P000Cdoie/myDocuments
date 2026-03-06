Imports System
Imports System.Data
Imports System.Data.SqlClient
'----------------------------------------------------
'Copyright(c)       - MIND
'Name of Module     - MARKETING CDR FUNCTIONALITY 
'Name of Form       - FRMMKTTRN0090.FRM  ,   INVOICE WISE ACTUAL AGREEMENT
'Created by         - Mayur Kumar
'Modified By        - 
'Created Date       - 16 SEP 2015
'description        - #10816097 -- New Forms Developed 
'*********************************************************************************************************************
Public Class FRMMKTTRN0090

#Region "GLOBAL VARIABLES"
    Dim SqlAdp As SqlDataAdapter
    Dim DSINVOICEDTL As DataSet
    Dim intLoopCounter As Int16 = 0
    Dim mstrErrMsg As String = String.Empty

    Private Enum ENUMINVOICEDETAILS
        VAL_INVOICENO = 1
        VAL_ITEMCODE
        VAL_ITEMDESC
        VAL_INVOICEQTY
        VAL_SHORTQTY
        VAL_CUSTOMERQTY
        VAL_REMARKS
    End Enum

    Dim VAL_ITEMCODE As Object = Nothing, VAL_ITEMDESC As Object = Nothing, VAL_INVOICEQTY As Object = Nothing
    Dim VAL_SHORTQTY As Object = Nothing, VAL_CUSTOMERQTY As Object = Nothing
    Dim VAL_REMARKS As Object = Nothing, VAL_INVOICENO As Object = Nothing

#End Region

#Region "Form Events"
    ' Form Load Event
    Private Sub FRMMKTTRN0091_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Call FitToClient(Me, grpMain, ctlHeader, grpcmdbtns)
            Me.MdiParent = mdifrmMain
            Me.grpcmdbtns.ShowButtons(True, True, True, False)

            cmdcustcodehelp.Enabled = False
            dtpFrm.Enabled = False
            dtpToDt.Enabled = False
            BtnHelpInvoiceNo.Enabled = False
            cmdContract.Enabled = True

            grpcmdbtns.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
            grpcmdbtns.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
            grpcmdbtns.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
            grpcmdbtns.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE) = True

            txt_ContractCode.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    ' Form Movement Abandoned
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
#End Region

#Region "Button Events"
    Private Sub grpcmdbtns_ButtonClick(ByVal Sender As System.Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles grpcmdbtns.ButtonClick
        Try
            Select Case e.Button
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                    Me.Close()
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                    cmdContract.Enabled = False
                    cmdcustcodehelp.Enabled = True
                    dtpFrm.Enabled = False
                    dtpToDt.Enabled = False
                    BtnHelpInvoiceNo.Enabled = False
                    fspr_Dtls.MaxRows = 0
                    fspr_Dtls.MaxCols = 0
                    txt_ContractCode.Text = ""
                    txtCustomerCode.Text = ""
                    txtInvoiceNo.Text = ""
                    lblCustomerdesc.Text = ""
                    'grpcmdbtns.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                    txtCustomerCode.Focus()
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE

                    If ValidData() = True Then
                        If SaveData() = True Then
                            MsgBox("Data Saved Successfully.", MsgBoxStyle.Information, ResolveResString(100))
                            Me.grpcmdbtns.Revert()
                            LockGrid()
                            '    'Initialize_controls()
                            grpcmdbtns.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                            grpcmdbtns.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                            'cmdButtonGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
                            grpcmdbtns.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                            grpcmdbtns.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE) = True
                            grpcmdbtns.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = True
                            dtpFrm.Enabled = False
                            dtpToDt.Enabled = False
                            BtnHelpInvoiceNo.Enabled = False
                            GETDOCUMENTDATA(Convert.ToInt32(Trim(txt_ContractCode.Text.Trim())))
                        End If
                    Else
                        MessageBox.Show(mstrErrMsg, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
                    End If

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                    Dim color As String = String.Empty
                    Dim countloop As Integer = 0
                    With Me.fspr_Dtls
                        For intLoopCounter = 1 To .MaxRows
                            .Row = intLoopCounter
                            .Row2 = intLoopCounter
                            .Col = 1
                            .Col2 = .MaxCols
                            color = .BackColor.ToString()
                            If color = "Color [Yellow]" Then
                                .BlockMode = True
                                .Row = intLoopCounter
                                .Row2 = intLoopCounter
                                .Col = ENUMINVOICEDETAILS.VAL_SHORTQTY
                                .Col2 = ENUMINVOICEDETAILS.VAL_SHORTQTY
                                .Lock = False
                                .BlockMode = False
                                countloop = 1
                            End If
                        Next
                    End With
                    If countloop = 0 Then
                        MessageBox.Show("No Item is available for EDIT!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
                        grpcmdbtns.Revert()
                    End If
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                    'If grpcmdbtns.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Or grpcmdbtns.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    fspr_Dtls.MaxRows = 0
                    fspr_Dtls.MaxCols = 0
                    txt_ContractCode.Text = ""
                    txtCustomerCode.Text = ""
                    txtInvoiceNo.Text = ""
                    lblCustomerdesc.Text = ""
                    BtnHelpInvoiceNo.Enabled = False
                    dtpFrm.Enabled = False
                    dtpToDt.Enabled = False
                    cmdcustcodehelp.Enabled = False
                    cmdContract.Enabled = True
                    txt_ContractCode.Focus()
                    grpcmdbtns.Revert()
                    grpcmdbtns.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    grpcmdbtns.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    'cmdButtonGrp.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
                    grpcmdbtns.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                    grpcmdbtns.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE) = True
                    grpcmdbtns.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                    'End If
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
                    MessageBox.Show("Delete Not Available!", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
                    Exit Sub
            End Select
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub cmdcustcodehelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdcustcodehelp.Click
        Try
            If grpcmdbtns.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then

                Dim strQuery As String
                Dim strHelp() As String

                strQuery = "SELECT CUSTOMER_CODE ,CUST_NAME  FROM CUSTOMER_MST WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE IS NOT NULL"

                strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Select Customer(s)")
                If UBound(strHelp) > 0 Then
                    If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
                        MsgBox("Customer Not Available.", MsgBoxStyle.Information, ResolveResString(100))
                        txtCustomerCode.Focus()
                        dtpFrm.Enabled = False
                        dtpToDt.Enabled = False
                        BtnHelpInvoiceNo.Enabled = False
                        Exit Sub
                    End If
                    If IsNothing(strHelp) = False Then
                        strQuery = String.Empty
                        Me.txtCustomerCode.Text = Trim(strHelp(0))
                        Me.lblCustomerdesc.Text = Trim(strHelp(1))
                        dtpFrm.Enabled = True
                        dtpToDt.Enabled = True
                        BtnHelpInvoiceNo.Enabled = True
                        dtpFrm.Focus()
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub BtnHelpInvoiceNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnHelpInvoiceNo.Click
        If grpcmdbtns.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            Try
                If txtCustomerCode.Text = "" Then
                    MsgBox("Customer Not Selected.", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If
                Dim strQuery As String
                Dim strHelp() As String

                strQuery = "SELECT DOC_NO ,CONVERT(VARCHAR(20),INVOICE_DATE,106) AS DOC_DATE  FROM SALESCHALLAN_DTL (NOLOCK) WHERE ACCOUNT_CODE ='" + txtCustomerCode.Text.Trim() + "' AND UNIT_CODE ='" + gstrUNITID + "' AND INVOICE_TYPE ='INV' and INVOICE_DATE BETWEEN '" + dtpFrm.Value.ToString("dd MMM yyyy") + "' and '" + dtpToDt.Value.ToString("dd MMM yyyy") + "'"

                strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Select Invoice(s)", 1, 1, "")
                If UBound(strHelp) > 0 Then
                    If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
                        MsgBox("Invoices Not Available.", MsgBoxStyle.Information, ResolveResString(100))
                        txtCustomerCode.Focus()
                        Exit Sub
                    End If
                    If IsNothing(strHelp) = False Then
                        strQuery = String.Empty
                        If strHelp.Length = 3 Then
                            strQuery = strHelp(2).Replace("|", ",")
                            txtInvoiceNo.Text = strQuery.ToString()
                            GETINVOICEDATA(txtInvoiceNo.Text.Trim.ToString())
                            'dtpFrm.Enabled = False
                            'dtpToDt.Enabled = False
                            cmdcustcodehelp.Enabled = False
                        End If
                    End If
                End If
            Catch ex As Exception
                RaiseException(ex)
            End Try
        End If
    End Sub
    Private Sub cmdContract_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdContract.Click
        Try
            If grpcmdbtns.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then

                Dim strQuery As String
                Dim strHelp() As String

                strQuery = "SELECT DISTINCT DOC_NO,CONVERT(varchar(20),DOC_DATE,106) as DOC_DATE FROM SHORTRECEIPT_HDR WHERE UNIT_CODE='" + gstrUNITID + "' ORDER BY DOC_DATE DESC"
                strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Document")
                If UBound(strHelp) > 0 Then
                    If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
                        MsgBox("Document Not Available.", MsgBoxStyle.Information, ResolveResString(100))
                        Exit Sub
                    End If
                    If IsNothing(strHelp) = False Then
                        txt_ContractCode.Text = Convert.ToInt32(Trim(strHelp(0)))
                        GETDOCUMENTDATA(Convert.ToInt32(Trim(strHelp(0))))
                        grpcmdbtns.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                        grpcmdbtns.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = True
                        'LockGrid()
                    End If

                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
#End Region

#Region "DATA EVENTS"

    Private Sub GETINVOICEDATA(ByRef invoices As String)
        Dim sqlCmd As New SqlCommand()
        SqlAdp = New SqlDataAdapter
        DSINVOICEDTL = New DataSet
        Try

            With sqlCmd
                .CommandText = "USP_CDR_ACTUAL_RECEIPT"
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .Connection = SqlConnectionclass.GetConnection()
                .Parameters.Clear()
                .Parameters.AddWithValue("@MODE", "ADD")
                .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                .Parameters.AddWithValue("@INVOICES", invoices)
                .Parameters.Add(New SqlParameter("@p_ERROR", SqlDbType.VarChar, 200, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                SqlAdp.SelectCommand = sqlCmd
                SqlAdp.Fill(DSINVOICEDTL)
                .Dispose()
                If Not String.IsNullOrEmpty(.Parameters("@p_ERROR").Value) Then
                    MessageBox.Show(Convert.ToString(sqlCmd.Parameters("@p_ERROR").Value), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                End If
            End With

            If DSINVOICEDTL.Tables.Count > 0 Then
                'DTL DATA
                If DSINVOICEDTL.Tables(0).Rows.Count > 0 Then
                    InitializeSpread_InvoiceDtls()
                    With Me.fspr_Dtls
                        For intLoopCounter = 0 To DSINVOICEDTL.Tables(0).Rows.Count - 1
                            AddRow()
                            .SetText(ENUMINVOICEDETAILS.VAL_INVOICENO, intLoopCounter + 1, DSINVOICEDTL.Tables(0).Rows(intLoopCounter).Item("INVOICE_NO").ToString.Trim)
                            .SetText(ENUMINVOICEDETAILS.VAL_ITEMCODE, intLoopCounter + 1, DSINVOICEDTL.Tables(0).Rows(intLoopCounter).Item("ITEM_CODE").ToString.Trim)
                            .SetText(ENUMINVOICEDETAILS.VAL_ITEMDESC, intLoopCounter + 1, DSINVOICEDTL.Tables(0).Rows(intLoopCounter).Item("ITEMDESC").ToString.Trim)
                            .SetText(ENUMINVOICEDETAILS.VAL_INVOICEQTY, intLoopCounter + 1, Convert.ToDouble(DSINVOICEDTL.Tables(0).Rows(intLoopCounter).Item("INVOICEQTY")))
                            .SetText(ENUMINVOICEDETAILS.VAL_SHORTQTY, intLoopCounter + 1, Convert.ToDouble(DSINVOICEDTL.Tables(0).Rows(intLoopCounter).Item("SHORTQTY")))
                            .SetText(ENUMINVOICEDETAILS.VAL_CUSTOMERQTY, intLoopCounter + 1, Convert.ToDouble(DSINVOICEDTL.Tables(0).Rows(intLoopCounter).Item("CUSTOMERQTY")))
                        Next
                    End With
                Else
                    fspr_Dtls.MaxRows = 0
                    fspr_Dtls.MaxCols = 0
                End If

                If DSINVOICEDTL.Tables(1).Rows.Count > 0 Then
                    txtInvoiceNo.Text = ""
                    For intLoopCounter = 0 To DSINVOICEDTL.Tables(1).Rows.Count - 1
                        txtInvoiceNo.Text = txtInvoiceNo.Text + DSINVOICEDTL.Tables(1).Rows(intLoopCounter).Item("INVOICE_NO").ToString() + ","
                    Next
                    txtInvoiceNo.Text = txtInvoiceNo.Text.Substring(0, txtInvoiceNo.Text.Trim.Length - 1)
                Else
                    txtInvoiceNo.Text = ""
                    MsgBox("No Item Available in Selected Invoice(s)")
                End If

            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally
            If sqlCmd.Connection.State = ConnectionState.Open Then sqlCmd.Connection.Close()
            sqlCmd.Connection.Dispose()
            sqlCmd.Dispose()
            DSINVOICEDTL.Clear()
        End Try

    End Sub
    Private Function ValidData() As Boolean
        Try

            'Dim msg_flag As Boolean = False
            Dim short_qty_total As Double = 0.0

            If txtCustomerCode.Text.ToString.Trim.Length = 0 Then
                mstrErrMsg = "Customer not selected!"
                Return False
            End If

            If txtInvoiceNo.Text.ToString.Trim.Length = 0 Then
                mstrErrMsg = "Invoices not selected!"
                Return False
            End If

            With fspr_Dtls
                VAL_SHORTQTY = Nothing
                For intLoopCounter = 1 To .MaxRows
                    .GetText(ENUMINVOICEDETAILS.VAL_SHORTQTY, intLoopCounter, VAL_SHORTQTY)
                    If IsNothing(VAL_SHORTQTY) = True Then VAL_SHORTQTY = 0.0
                    short_qty_total = short_qty_total + Convert.ToDouble(VAL_SHORTQTY)
                    VAL_SHORTQTY = Nothing
                Next
                If short_qty_total = 0.0 Then
                    mstrErrMsg = "No item Selected!"
                    Return False
                End If
            End With

            Return True
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Function
    Private Function SaveData() As Boolean
        Dim data_farspreadGrid As New DataTable  '  Contract dtl table data
        Dim sqlCmd As New SqlCommand()
        Try
            data_farspreadGrid.Columns.Add("INVOICE_NO", GetType(String))
            data_farspreadGrid.Columns.Add("ITEM_CODE", GetType(String))
            data_farspreadGrid.Columns.Add("ITEMDESC", GetType(String))
            data_farspreadGrid.Columns.Add("INVOICEQTY", GetType(Double))
            data_farspreadGrid.Columns.Add("SHORTQTY", GetType(Double))
            data_farspreadGrid.Columns.Add("CUSTOMERQTY", GetType(Double))
            data_farspreadGrid.Columns.Add("REMARKS", GetType(String))

            If grpcmdbtns.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then

                With fspr_Dtls
                    For intLoopCounter = 1 To .MaxRows
                        Dim data_newRow As DataRow = data_farspreadGrid.NewRow

                        VAL_INVOICENO = Nothing
                        .GetText(ENUMINVOICEDETAILS.VAL_INVOICENO, intLoopCounter, VAL_INVOICENO)
                        If IsNothing(VAL_INVOICENO) = True Then VAL_INVOICENO = String.Empty

                        VAL_ITEMCODE = Nothing
                        .GetText(ENUMINVOICEDETAILS.VAL_ITEMCODE, intLoopCounter, VAL_ITEMCODE)
                        If IsNothing(VAL_ITEMCODE) = True Then VAL_ITEMCODE = String.Empty

                        VAL_ITEMDESC = Nothing
                        .GetText(ENUMINVOICEDETAILS.VAL_ITEMDESC, intLoopCounter, VAL_ITEMDESC)
                        If IsNothing(VAL_ITEMDESC) = True Then VAL_ITEMDESC = String.Empty

                        VAL_INVOICEQTY = Nothing
                        .GetText(ENUMINVOICEDETAILS.VAL_INVOICEQTY, intLoopCounter, VAL_INVOICEQTY)
                        If IsNothing(VAL_INVOICEQTY) = True Then VAL_INVOICEQTY = 0.0

                        VAL_SHORTQTY = Nothing
                        .GetText(ENUMINVOICEDETAILS.VAL_SHORTQTY, intLoopCounter, VAL_SHORTQTY)
                        If IsNothing(VAL_SHORTQTY) = True Then VAL_SHORTQTY = 0.0

                        VAL_CUSTOMERQTY = Nothing
                        .GetText(ENUMINVOICEDETAILS.VAL_CUSTOMERQTY, intLoopCounter, VAL_CUSTOMERQTY)
                        If IsNothing(VAL_CUSTOMERQTY) = True Then VAL_CUSTOMERQTY = 0.0

                        VAL_REMARKS = Nothing
                        .GetText(ENUMINVOICEDETAILS.VAL_REMARKS, intLoopCounter, VAL_REMARKS)
                        If IsNothing(VAL_REMARKS) = True Then VAL_REMARKS = String.Empty

                        If (VAL_SHORTQTY > 0.0) Then

                            data_newRow("INVOICE_NO") = VAL_INVOICENO.ToString()
                            data_newRow("ITEM_CODE") = VAL_ITEMCODE.ToString()
                            data_newRow("ITEMDESC") = VAL_ITEMDESC.ToString()
                            data_newRow("INVOICEQTY") = Convert.ToDouble(VAL_INVOICEQTY)
                            data_newRow("SHORTQTY") = Convert.ToDouble(VAL_SHORTQTY)
                            data_newRow("CUSTOMERQTY") = Convert.ToDouble(VAL_CUSTOMERQTY)
                            data_newRow("REMARKS") = VAL_REMARKS.ToString()

                            data_farspreadGrid.Rows.Add(data_newRow)

                        End If

                    Next
                End With
            End If

            If grpcmdbtns.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                Dim color As String = String.Empty
                With fspr_Dtls
                    For intLoopCounter = 1 To .MaxRows
                        .Row = intLoopCounter
                        .Row2 = intLoopCounter
                        .Col = 1
                        .Col2 = .MaxCols
                        Color = .BackColor.ToString()
                        If Color = "Color [Yellow]" Then
                            Dim data_newRow As DataRow = data_farspreadGrid.NewRow

                            VAL_INVOICENO = Nothing
                            .GetText(ENUMINVOICEDETAILS.VAL_INVOICENO, intLoopCounter, VAL_INVOICENO)
                            If IsNothing(VAL_INVOICENO) = True Then VAL_INVOICENO = String.Empty

                            VAL_ITEMCODE = Nothing
                            .GetText(ENUMINVOICEDETAILS.VAL_ITEMCODE, intLoopCounter, VAL_ITEMCODE)
                            If IsNothing(VAL_ITEMCODE) = True Then VAL_ITEMCODE = String.Empty

                            VAL_ITEMDESC = Nothing
                            .GetText(ENUMINVOICEDETAILS.VAL_ITEMDESC, intLoopCounter, VAL_ITEMDESC)
                            If IsNothing(VAL_ITEMDESC) = True Then VAL_ITEMDESC = String.Empty

                            VAL_INVOICEQTY = Nothing
                            .GetText(ENUMINVOICEDETAILS.VAL_INVOICEQTY, intLoopCounter, VAL_INVOICEQTY)
                            If IsNothing(VAL_INVOICEQTY) = True Then VAL_INVOICEQTY = 0.0

                            VAL_SHORTQTY = Nothing
                            .GetText(ENUMINVOICEDETAILS.VAL_SHORTQTY, intLoopCounter, VAL_SHORTQTY)
                            If IsNothing(VAL_SHORTQTY) = True Then VAL_SHORTQTY = 0.0

                            VAL_CUSTOMERQTY = Nothing
                            .GetText(ENUMINVOICEDETAILS.VAL_CUSTOMERQTY, intLoopCounter, VAL_CUSTOMERQTY)
                            If IsNothing(VAL_CUSTOMERQTY) = True Then VAL_CUSTOMERQTY = 0.0

                            VAL_REMARKS = Nothing
                            .GetText(ENUMINVOICEDETAILS.VAL_REMARKS, intLoopCounter, VAL_REMARKS)
                            If IsNothing(VAL_REMARKS) = True Then VAL_REMARKS = String.Empty

                            If (VAL_SHORTQTY > 0.0) Then

                                data_newRow("INVOICE_NO") = VAL_INVOICENO.ToString()
                                data_newRow("ITEM_CODE") = VAL_ITEMCODE.ToString()
                                data_newRow("ITEMDESC") = VAL_ITEMDESC.ToString()
                                data_newRow("INVOICEQTY") = Convert.ToDouble(VAL_INVOICEQTY)
                                data_newRow("SHORTQTY") = Convert.ToDouble(VAL_SHORTQTY)
                                data_newRow("CUSTOMERQTY") = Convert.ToDouble(VAL_CUSTOMERQTY)
                                data_newRow("REMARKS") = VAL_REMARKS.ToString()

                                data_farspreadGrid.Rows.Add(data_newRow)

                            End If
                        End If
                    Next
                End With
            End If


            If data_farspreadGrid.Rows.Count = 0 Then
                MsgBox("No item selected!")
                Return False
            End If

            With sqlCmd
                .CommandText = "USP_CDR_ACTUAL_RECEIPT"
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .Connection = SqlConnectionclass.GetConnection()
                .Parameters.Clear()

                If grpcmdbtns.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    .Parameters.AddWithValue("@MODE", "SAVE")
                End If
                If grpcmdbtns.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                    .Parameters.AddWithValue("@MODE", "UPDATE")
                End If

                .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                .Parameters.AddWithValue("@CUSTOMER_CODE", txtCustomerCode.Text.Trim.ToString())
                .Parameters.AddWithValue("@CUSTOMER_DESC", lblCustomerdesc.Text.Trim.ToString())
                .Parameters.AddWithValue("@USER_ID", gstrUserIDSelected.ToString())
                .Parameters.AddWithValue("@INVOICE_DTLS", data_farspreadGrid)
                .Parameters.Add(New SqlParameter("@p_ERROR", SqlDbType.VarChar, 200, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))

                If grpcmdbtns.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    txt_ContractCode.Text = Convert.ToInt32(.ExecuteScalar)
                End If
                If grpcmdbtns.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                    .Parameters.AddWithValue("@DOCUMENTNO", Convert.ToInt32(txt_ContractCode.Text.Trim()))
                    .ExecuteNonQuery()
                End If
                .Dispose()
                If Not String.IsNullOrEmpty(.Parameters("@p_ERROR").Value) Then
                    MessageBox.Show(Convert.ToString(sqlCmd.Parameters("@p_ERROR").Value), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                End If
            End With

            fspr_Dtls.Lock = True
            Return True

        Catch ex As Exception
            RaiseException(ex)
        Finally
            If sqlCmd.Connection.State = ConnectionState.Open Then sqlCmd.Connection.Close()
            sqlCmd.Connection.Dispose()
            sqlCmd.Dispose()
            data_farspreadGrid.Dispose()

        End Try
    End Function
    Private Sub GETDOCUMENTDATA(ByRef doc_no As Integer)
        Dim sqlCmd As New SqlCommand()
        Dim color As String = String.Empty
        SqlAdp = New SqlDataAdapter
        DSINVOICEDTL = New DataSet
        Try
            With sqlCmd
                .CommandText = "USP_CDR_ACTUAL_RECEIPT"
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .Connection = SqlConnectionclass.GetConnection()
                .Parameters.Clear()
                .Parameters.AddWithValue("@MODE", "VIEW")
                .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                .Parameters.AddWithValue("@DOCUMENTNO", doc_no)
                .Parameters.Add(New SqlParameter("@p_ERROR", SqlDbType.VarChar, 200, ParameterDirection.InputOutput, True, 0, 0, "", DataRowVersion.Default, ""))
                SqlAdp.SelectCommand = sqlCmd
                SqlAdp.Fill(DSINVOICEDTL)
                .Dispose()
                If Not String.IsNullOrEmpty(.Parameters("@p_ERROR").Value) Then
                    MessageBox.Show(Convert.ToString(sqlCmd.Parameters("@p_ERROR").Value), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                End If
            End With

            If DSINVOICEDTL.Tables.Count > 0 Then
                ' HDR DATA 
                If DSINVOICEDTL.Tables(0).Rows.Count > 0 Then
                    txtCustomerCode.Text = DSINVOICEDTL.Tables(0).Rows(0).Item("CUSTOMER_CODE").ToString()
                    lblCustomerdesc.Text = DSINVOICEDTL.Tables(0).Rows(0).Item("Cust_Name").ToString()
                End If
                If DSINVOICEDTL.Tables(1).Rows.Count > 0 Then
                    txtInvoiceNo.Text = ""
                    For intLoopCounter = 0 To DSINVOICEDTL.Tables(1).Rows.Count - 1
                        txtInvoiceNo.Text = txtInvoiceNo.Text + DSINVOICEDTL.Tables(1).Rows(intLoopCounter).Item("INVOICE_NO").ToString() + ","
                    Next
                    txtInvoiceNo.Text = txtInvoiceNo.Text.Substring(0, txtInvoiceNo.Text.Trim.Length - 1)
                End If

                'DTL DATA

                If DSINVOICEDTL.Tables(2).Rows.Count > 0 Then
                    InitializeSpread_InvoiceDtls()
                    With Me.fspr_Dtls
                        For intLoopCounter = 0 To DSINVOICEDTL.Tables(2).Rows.Count - 1
                            AddRow()
                            .SetText(ENUMINVOICEDETAILS.VAL_INVOICENO, intLoopCounter + 1, DSINVOICEDTL.Tables(2).Rows(intLoopCounter).Item("INVOICE_NO").ToString.Trim)
                            .SetText(ENUMINVOICEDETAILS.VAL_ITEMCODE, intLoopCounter + 1, DSINVOICEDTL.Tables(2).Rows(intLoopCounter).Item("ITEM_CODE").ToString.Trim)
                            .SetText(ENUMINVOICEDETAILS.VAL_ITEMDESC, intLoopCounter + 1, DSINVOICEDTL.Tables(2).Rows(intLoopCounter).Item("ITEMDESC").ToString.Trim)
                            .SetText(ENUMINVOICEDETAILS.VAL_INVOICEQTY, intLoopCounter + 1, Convert.ToDouble(DSINVOICEDTL.Tables(2).Rows(intLoopCounter).Item("INVOICEQTY")))
                            .SetText(ENUMINVOICEDETAILS.VAL_SHORTQTY, intLoopCounter + 1, Convert.ToDouble(DSINVOICEDTL.Tables(2).Rows(intLoopCounter).Item("SHORTQTY")))
                            .SetText(ENUMINVOICEDETAILS.VAL_CUSTOMERQTY, intLoopCounter + 1, Convert.ToDouble(DSINVOICEDTL.Tables(2).Rows(intLoopCounter).Item("CUSTOMERQTY")))
                            .SetText(ENUMINVOICEDETAILS.VAL_REMARKS, intLoopCounter + 1, DSINVOICEDTL.Tables(2).Rows(intLoopCounter).Item("REMARKS").ToString.Trim)
                            color = DSINVOICEDTL.Tables(2).Rows(intLoopCounter).Item("STATUS").ToString.Trim
                            .BlockMode = True
                            .Row = intLoopCounter + 1
                            .Row2 = intLoopCounter + 1
                            .Col = 1
                            .Col2 = Me.fspr_Dtls.MaxCols
                            If color = "SUBMITTED" Then
                                .BackColor = Drawing.Color.Yellow
                            End If
                            If color = "APPROVED" Then
                                .BackColor = Drawing.Color.Green
                            End If
                            If color = "REJECTED" Then
                                .BackColor = Drawing.Color.Red
                            End If
                            .BlockMode = False
                            color = ""
                        Next
                    End With
                End If
            End If


        Catch ex As Exception
            RaiseException(ex)
        Finally
            If sqlCmd.Connection.State = ConnectionState.Open Then sqlCmd.Connection.Close()
            sqlCmd.Connection.Dispose()
            sqlCmd.Dispose()
            DSINVOICEDTL.Clear()
        End Try
    End Sub

#End Region

#Region "GRID EVENTS"

    Private Sub InitializeSpread_InvoiceDtls()
        Try
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            With Me.fspr_Dtls
                .MaxRows = 0
                .MaxCols = ENUMINVOICEDETAILS.VAL_REMARKS
                .set_RowHeight(0, 20)
                .Row = 0 : .Col = ENUMINVOICEDETAILS.VAL_INVOICENO : .Text = "INVOICE NO" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMINVOICEDETAILS.VAL_ITEMCODE : .Text = "ITEM CODE" : .set_ColWidth(.Col, 20) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMINVOICEDETAILS.VAL_ITEMDESC : .Text = "ITEM DESCRIPTION" : .set_ColWidth(.Col, 40) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMINVOICEDETAILS.VAL_INVOICEQTY : .Text = "INVOICE QTY" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMINVOICEDETAILS.VAL_SHORTQTY : .Text = "SHORT RECEIPT" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMINVOICEDETAILS.VAL_CUSTOMERQTY : .Text = "CUSTOMER RECEIPT" : .set_ColWidth(.Col, 10) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = 0 : .Col = ENUMINVOICEDETAILS.VAL_REMARKS : .Text = "REMARKS" : .set_ColWidth(.Col, 25) : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
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
            With fspr_Dtls
                .MaxRows = .MaxRows + 1

                .Row = .MaxRows : .Col = ENUMINVOICEDETAILS.VAL_INVOICENO : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMINVOICEDETAILS.VAL_ITEMCODE : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMINVOICEDETAILS.VAL_ITEMDESC : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMINVOICEDETAILS.VAL_INVOICEQTY : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeFloatMin = 0 : .TypeFloatDecimalPlaces = 4 : .BorderStyle = FPSpreadADO.BorderStyleConstants.BorderStyleFixedSingle : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                .Row = .MaxRows : .Col = ENUMINVOICEDETAILS.VAL_CUSTOMERQTY : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeFloatMin = 0 : .TypeFloatDecimalPlaces = 4 : .BorderStyle = FPSpreadADO.BorderStyleConstants.BorderStyleFixedSingle : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)

                If grpcmdbtns.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or grpcmdbtns.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                    .Row = .MaxRows : .Col = ENUMINVOICEDETAILS.VAL_SHORTQTY : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .Lock = False : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeFloatMin = 0 : .TypeFloatDecimalPlaces = 4 : .BorderStyle = FPSpreadADO.BorderStyleConstants.BorderStyleFixedSingle : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                    .Row = .MaxRows : .Col = ENUMINVOICEDETAILS.VAL_REMARKS : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .Lock = False : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                End If

                If grpcmdbtns.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    .Row = .MaxRows : .Col = ENUMINVOICEDETAILS.VAL_SHORTQTY : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeFloatMin = 0 : .TypeFloatDecimalPlaces = 4 : .BorderStyle = FPSpreadADO.BorderStyleConstants.BorderStyleFixedSingle : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                    .Row = .MaxRows : .Col = ENUMINVOICEDETAILS.VAL_REMARKS : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Lock = True : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .SetCellBorder(.Col, .Row, .Col, .Row, FPSpreadADO.CellBorderIndexConstants.CellBorderIndexOutline, &H0, FPSpreadADO.CellBorderStyleConstants.CellBorderStyleSolid)
                End If

            End With

            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub fspr_Dtls_EditChange(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles fspr_Dtls.EditChange
        Try
            If Me.fspr_Dtls.ActiveCol = ENUMINVOICEDETAILS.VAL_SHORTQTY Then
                Edit_CustomerQty()
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub Edit_CustomerQty()
        Try
            With Me.fspr_Dtls
                intLoopCounter = fspr_Dtls.ActiveRow

                VAL_INVOICEQTY = Nothing
                .GetText(ENUMINVOICEDETAILS.VAL_INVOICEQTY, intLoopCounter, VAL_INVOICEQTY)

                VAL_SHORTQTY = Nothing
                .GetText(ENUMINVOICEDETAILS.VAL_SHORTQTY, intLoopCounter, VAL_SHORTQTY)

                If Convert.ToDouble(VAL_INVOICEQTY) < Convert.ToDouble(VAL_SHORTQTY) Then
                    MsgBox("Short Receipt can not be greater than Invoice Qty.!!", MsgBoxStyle.Information, ResolveResString(100))
                    .SetText(ENUMINVOICEDETAILS.VAL_SHORTQTY, intLoopCounter, Convert.ToDouble(0.0))
                    .SetText(ENUMINVOICEDETAILS.VAL_CUSTOMERQTY, intLoopCounter, Convert.ToDouble(VAL_INVOICEQTY))
                Else
                    VAL_CUSTOMERQTY = Nothing
                    VAL_CUSTOMERQTY = (Convert.ToDouble(VAL_INVOICEQTY) - Convert.ToDouble(VAL_SHORTQTY))

                    .SetText(ENUMINVOICEDETAILS.VAL_CUSTOMERQTY, intLoopCounter, Convert.ToDouble(VAL_CUSTOMERQTY))
                End If
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub LockGrid()
        With fspr_Dtls
            For intLoopCounter = 1 To .MaxRows
                .Row = intLoopCounter : .Col = ENUMINVOICEDETAILS.VAL_INVOICENO : .Lock = True
                .Row = intLoopCounter : .Col = ENUMINVOICEDETAILS.VAL_ITEMCODE : .Lock = True
                .Row = intLoopCounter : .Col = ENUMINVOICEDETAILS.VAL_ITEMDESC : .Lock = True
                .Row = intLoopCounter : .Col = ENUMINVOICEDETAILS.VAL_INVOICEQTY : .Lock = True
                .Row = intLoopCounter : .Col = ENUMINVOICEDETAILS.VAL_SHORTQTY : .Lock = True
                .Row = intLoopCounter : .Col = ENUMINVOICEDETAILS.VAL_CUSTOMERQTY : .Lock = True
                .Row = intLoopCounter : .Col = ENUMINVOICEDETAILS.VAL_REMARKS : .Lock = True
            Next
        End With
    End Sub

    Private Sub txt_ContractCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt_ContractCode.KeyDown
        Try
            If grpcmdbtns.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                If e.KeyCode = Keys.F1 Then
                    cmdContract_Click(sender, e)
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtCustomerCode_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Try
            If grpcmdbtns.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If e.KeyCode = Keys.F1 Then
                    cmdcustcodehelp_Click(sender, e)
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtInvoiceNo_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtInvoiceNo.KeyDown
        Try
            If grpcmdbtns.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If e.KeyCode = Keys.F1 Then
                    BtnHelpInvoiceNo_Click(sender, e)
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
#End Region


End Class