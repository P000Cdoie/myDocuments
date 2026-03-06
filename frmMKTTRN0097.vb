Imports System.Data
Imports System.Data.SqlClient


Public Class frmMKTTRN0097

    Private Sub CMdHelpToolno_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMdHelpToolno.Click
        Try
            Dim strHelp() As String
            Dim strQuery As String
            If chkDelLoadingslip.CheckState = CheckState.Checked Then
                strQuery = "SELECT DISTINCT X.SLIP_NO,Y.DOC_NO,X.PICK_LIST_NO,CONVERT(VARCHAR(10),X.SLIPDATE,103) AS SLIP_DATE,CASE WHEN Y.MODE='61' THEN 'ONLINE' WHEN Y.MODE='63' THEN 'OFFLINE' END AS MODE  FROM LOADINGSLIP  X  " & _
                " INNER JOIN TRIGGER_FILE_MARUTI_SCANDATA  Y   " & _
                " ON X.UNIT_CODE=Y.UNIT_CODE AND X.CUSTPART_NO =Y.FGPART_NO   AND X.PICK_LIST_NO=Y.PICKLIST_NO AND X.SEQ_NO=Y.PSN AND X.SLIP_NO=Y.LoadingSlip_No " & _
                " WHERE X.UNIT_CODE='" & gstrUNITID & "' AND X.CUSTOMER_CODE='C0000037' AND Y.ISLOADINGSLIPGENERATED =1 AND X.INVOICENO IS NULL" & _
                " ORDER BY Y.DOC_NO,X.SLIP_NO ASC"
            Else
                strQuery = "SELECT DISTINCT A.DOC_NO,convert(varchar(10), A.ENTERED_DATE,103) as Upload_Date,A.CUST_CODE FROM TRIGGER_FILE_MARUTI A WHERE A.UNIT_CODE='" & gstrUNITID & "' AND A.ISINVGENERATED='N'  AND A.CUST_CODE='C0000037' " & _
                           " AND DOC_NO NOT IN (SELECT DISTINCT DOC_NO FROM  TRIGGER_FILE_MARUTI_SCANDATA(NOLOCK) WHERE UNIT_CODE='" & gstrUNITID & "' )  ORDER BY A.DOC_NO,A.CUST_CODE "
            End If
            strHelp = CtlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQuery, "Rout Operation Detail")
            If UBound(strHelp) = -1 Then Exit Sub
            If UBound(strHelp) > 0 Then
                If Trim(strHelp(0)) = "0" Or Trim(strHelp(0)) = String.Empty Then
                    MessageBox.Show("Nothing available to show.", "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Hand)
                    txtDocNo.Text = String.Empty
                    If chkDelLoadingslip.CheckState = CheckState.Checked Then
                        chkDelLoadingslip.CheckState = CheckState.Unchecked
                        lblLoadingslipNo.Visible = False
                        lblLoadingslipNo.Text = String.Empty

                    End If
                    Exit Sub
                Else
                    If chkDelLoadingslip.CheckState = CheckState.Checked Then
                        lblLoadingslipNo.Visible = True
                        lblLoadingslipNo.Text = strHelp(1).ToString()
                    Else
                        lblLoadingslipNo.Visible = False
                        lblLoadingslipNo.Text = String.Empty
                    End If
                    txtDocNo.Text = strHelp(0).ToString()
                   
                End If
            End If
           
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtDocNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDocNo.TextChanged
        Try
            txtDocNo_Validating(Nothing, Nothing)
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtDocNo_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDocNo.Validating
        Try
            If txtDocNo.Text.Trim.Length = 0 Then
                DGVDocDetail.DataSource = Nothing
                Exit Sub
            End If

            Dim Qselect As String = String.Empty
            If chkDelLoadingslip.CheckState = CheckState.Checked Then
                Qselect = "SELECT DOC_NO,LOADINGSLIP_NO,PICKLIST_NO, PSN,CHASSIS,FGPART_NO,QTY,CONVERT(VARCHAR(15),SLIP_DATE,103) AS SLIP_DATE,CASE WHEN MODE='61' THEN 'ONLINE' WHEN MODE='63' THEN 'OFFLINE' END ON_OFF_MODE " & _
                " FROM TRIGGER_FILE_MARUTI_SCANDATA(NOLOCK) WHERE UNIT_CODE='" & gstrUNITID & "'  AND LOADINGSLIP_NO=" & txtDocNo.Text.ToString.Trim & " AND DOC_NO=" & lblLoadingslipNo.Text.ToString.Trim & ""
            Else
                Qselect = "Select PICKLIST_NO, PSN,CHASSIS,MODEL_CODE,ACHV_DATE,PART_NO,CONVERT(VARCHAR(15),ENTERED_DATE,103) AS Upload_Date,Uploaded_FileName,Mset,CASE WHEN Front_RearType='R' THEN 'REAR' WHEN Front_RearType='F' THEN 'FRONT' END  Front_RearType,CASE WHEN LEFT(Mode,2)='61' THEN 'ONLINE' WHEN LEFT(Mode,2)='63' THEN 'OFFLINE' END ON_OFF_MODE" & _
                      " from trigger_file_maruti(NOLOCK) where unit_code='" & gstrUNITID & "' and DOC_NO=" & txtDocNo.Text.ToString.Trim & ""
            End If
            
            Dim da As SqlDataAdapter = New SqlDataAdapter(Qselect, SqlConnectionclass.GetConnection)
            Dim dt As New DataTable
            da.Fill(dt)
            If dt.Rows.Count > 0 Then
                DGVDocDetail.DataSource = dt
                cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = True
                da.Dispose()
            Else
                DGVDocDetail.DataSource = Nothing
                MessageBox.Show("Nothing To Show", "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Dim mintFormIndex As Integer
    Private Sub frmMKTTRN0097_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Try
            mdifrmMain.CheckFormName = mintFormIndex
            System.Windows.Forms.Application.DoEvents()
            frmModules.NodeFontBold(Me.Tag) = True
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub frmMKTTRN0097_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        Try
            frmModules.NodeFontBold(Me.Tag) = False
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub frmMKTTRN0097_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            Dim KeyCode As Short = e.KeyCode
            Dim Shift As Short = e.KeyData \ &H10000
            If Shift <> 0 Then Exit Sub
            If KeyCode = System.Windows.Forms.Keys.F4 Then Call ctlFormHeader1_Click(ctlFormHeader1, New System.EventArgs()) : Exit Sub
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub frmMKTTRN0097_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Call FitToClient(Me, GbxMain, ctlFormHeader1, cmdButtons)
            Me.MdiParent = mdifrmMain

            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = False

            CMdHelpToolno.Enabled = True
            txtDocNo.Focus()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        Try
            Call ShowHelp("underconstruction.htm")
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

   
    
    Private Sub cmdButtons_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles cmdButtons.ButtonClick
        Try
            Select Case e.Button
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                    DGVDocDetail.DataSource = Nothing
                    txtDocNo.Text = String.Empty
                    cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = False
                    cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                    txtDocNo.Focus()
                    lblLoadingslipNo.Text = String.Empty
                    lblLoadingslipNo.Visible = False
                    chkDelLoadingslip.CheckState = CheckState.Unchecked
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                    Me.Close()

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION, 60095) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Dim srtSql As String = String.Empty
                        If chkDelLoadingslip.CheckState = CheckState.Checked Then
                            If lblLoadingslipNo.Text.ToString.Trim.Length = 0 Or txtDocNo.Text.ToString.Trim.Length = 0 Then
                                MessageBox.Show("SLip No and Doc No required Before Deletion.", "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                Exit Sub
                            End If
                            srtSql = " DELETE FROM LOADINGSLIP WHERE UNIT_CODE='" & gstrUNITID & "' AND SLIP_NO=" & txtDocNo.Text.ToString.Trim & " AND Customer_Code='C0000037' AND InvoiceNo IS NULL AND ACT_INV_NO IS NULL ;" & _
                            " UPDATE TRIGGER_FILE_MARUTI_SCANDATA SET IsLoadingSlipGenerated=0,LoadingSlip_No=NULL,Slip_date=NULL,upd_Dt=GETDATE(),Upd_By='" & mP_User & "' " & _
                            " WHERE UNIT_CODE='" & gstrUNITID & "' AND DOC_NO=" & lblLoadingslipNo.Text.ToString.Trim & " AND LoadingSlip_No=" & txtDocNo.Text.ToString.Trim & " AND IsLoadingSlipGenerated=1"
                        Else
                            srtSql = " DELETE FROM TRIGGER_FILE_MARUTI WHERE UNIT_CODE='" & gstrUNITID & "' AND DOC_NO='" & txtDocNo.Text.ToString.Trim & "'"
                        End If

                        SqlConnectionclass.CloseGlobalConnection()
                        SqlConnectionclass.OpenGlobalConnection()

                        Dim cmd As SqlCommand = Nothing
                        cmd = New System.Data.SqlClient.SqlCommand()
                        cmd.Connection = SqlConnectionclass.GetConnection
                        cmd.Transaction = cmd.Connection.BeginTransaction
                        Try
                            With cmd
                                .CommandText = String.Empty
                                .Parameters.Clear()
                                .CommandTimeout = 0
                                .CommandType = CommandType.Text
                                .CommandText = srtSql

                                If cmd.ExecuteNonQuery() Then
                                    cmd.Transaction.Commit()
                                    MessageBox.Show("Document Deleted Successfully.", "eMPRO", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                    DGVDocDetail.DataSource = Nothing
                                    txtDocNo.Text = String.Empty
                                    cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                                    cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                                    txtDocNo.Focus()
                                    If chkDelLoadingslip.CheckState Then
                                        lblLoadingslipNo.Text = String.Empty
                                        lblLoadingslipNo.Visible = False
                                        chkDelLoadingslip.CheckState = CheckState.Unchecked
                                    End If
                                Else
                                    cmd.Transaction.Rollback()
                                End If
                            End With
                        Catch ex As Exception
                            cmd.Transaction.Dispose()
                            MessageBox.Show(ex.Message, "eMPRO")
                        End Try

                    Else
                        DGVDocDetail.DataSource = Nothing
                        txtDocNo.Text = String.Empty
                        txtDocNo.Focus()
                    End If

            End Select


        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
End Class